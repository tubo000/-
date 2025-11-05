# main_application.py (バッチ処理・軽量化 最終版)
import os
import sys
import pandas as pd
import win32com.client as win32
import threading 
import tkinter as tk
from tkinter import Frame, messagebox, simpledialog, ttk 
import pythoncom 
import re 
import traceback 
import os.path
import datetime 
import queue 
import sqlite3 
from datetime import timedelta 
# 外部モジュールのインポート
import gui_elements
import gui_search_window 
import utils 

# 既存の内部処理関数をインポート
from config import INPUT_QUESTION_CSV, MASTER_ANSWERS_PATH, OUTPUT_EVAL_PATH, NUM_RECORDS, TARGET_FOLDER_PATH, SCRIPT_DIR
from extraction_core import extract_skills_data
from evaluator_core import run_triple_csv_validation, get_question_data_from_csv
from email_processor import get_mail_data_from_outlook_in_memory
from email_processor import has_unprocessed_mail 
# 📌 修正: remove_processed_category, mark_email_as_processed はもう使わない
# from email_processor import remove_processed_category, PROCESSED_CATEGORY_NAME
# from email_processor import mark_email_as_processed
from config import DATABASE_NAME 


# --- グローバル変数の定義をファイル先頭に移動 ---
root = None
main_elements = {}
# ----------------------------------------------------

def open_outlook_email_by_id(entry_id: str):
    """Entry IDを使用してOutlookデスクトップアプリでメールを開く関数。（GUI版）"""
    if not entry_id:
        messagebox.showerror("エラー", "Entry IDが指定されていません。")
        return
    try:
        pythoncom.CoInitialize() 
        try:
            outlook_app = win32.GetActiveObject("Outlook.Application")
        except:
            outlook_app = win32.Dispatch("Outlook.Application")
        namespace = outlook_app.GetNamespace("MAPI")
        olItem = namespace.GetItemFromID(entry_id)
        if olItem:
            olItem.Display()
        else:
            messagebox.showerror("エラー", "指定された Entry ID のメールが見つかりませんでした。")
    except Exception as e:
        messagebox.showerror("Outlook連携エラー", f"Outlook連携中にエラーが発生しました: {e}\nOutlookが起動しているか確認してください。")
    finally:
        pythoncom.CoUninitialize() 


def interactive_id_search_test():
    pass


def reorder_output_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    fixed_leading_cols = [
        'メールURL', '受信日時', '件名', '名前', '信頼度スコア', 
        '本文(テキスト形式)', '本文(ファイル含む)', 'Attachments'
    ]
    fixed_leading_cols = [col for col in fixed_leading_cols if col in df.columns]
    remaining_cols = [col for col in df.columns.tolist() if col not in fixed_leading_cols]
    return df.reindex(columns=fixed_leading_cols + remaining_cols, fill_value='N/A')

# ----------------------------------------------------
# 抽出処理ロジック (0日対応 + バッチ処理 + 安全なマーク付け)
# ----------------------------------------------------
def actual_run_extraction_logic(root, main_elements, target_email, folder_path, read_mode, read_days, status_label):
    
    thread_id = threading.get_ident()
    try:
        pythoncom.CoInitialize()
    except Exception:
        pass 
        
    # --- ▼▼▼【修正】3つの変数をすべてここで初期化 ▼▼▼ ---
    total_new_records_saved = 0 # DBに保存した総件数を初期化
    total_items_skipped = 0     # スキップIDを保存した総件数を初期化
    total_items_marked = 0      # ★★★ マーク付けした総件数を初期化 (これ) ★★★
    # --- ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲ ---
    
    outlook_app = None # マーク付け用に保持
    namespace = None
    
    try:
        days_ago = None
        if read_days.strip():
            try:
                days_ago = int(read_days)
                if days_ago < 0: 
                    raise ValueError("日数は0以上の整数である必要があります")
            except ValueError:
                messagebox.showerror("入力エラー", "期間指定は **0以上** の整数で指定してください。\n(空欄の場合は全期間)")
                status_label.config(text="状態: 抽出失敗 (期間入力不正)。")
                return 

        if days_ago == 0:
             mode_text = "未処理 (今日のみ)"
        elif days_ago is not None and days_ago > 0 :
             mode_text = f"未処理 (過去{days_ago}日)"
        else:
             mode_text = "未処理 (全期間)"
            
        status_label.config(text=f"状態: {target_email} アカウントからメール取得中 ({mode_text})...")
        root.update_idletasks() 

        # email_processor からジェネレータ（バッチ）を受け取る
        df_mail_data_generator = get_mail_data_from_outlook_in_memory(
            folder_path, 
            target_email, 
            read_mode=read_mode, 
            days_ago=days_ago,
            main_elements=main_elements 
        )
        
        status_label.config(text="状態: 抽出コアロジック実行中...")
        root.update_idletasks()

        # 📌 修正: Outlook接続はもう不要
        # try:
        #     outlook_app = win32.GetActiveObject("Outlook.Application")
        # except:
        #     outlook_app = win32.Dispatch("Outlook.Application")
        # namespace = outlook_app.GetNamespace("MAPI")

        db_path = os.path.abspath(DATABASE_NAME) 
        
        batch_number = 0
        for batch_data in df_mail_data_generator:
            batch_number += 1
            
            df_mail_data_batch = batch_data.get("data_batch", pd.DataFrame())
            skip_ids_batch = batch_data.get("skip_ids", [])
            
            if df_mail_data_batch.empty and not skip_ids_batch:
                continue 

            status_label.config(text=f"状態: バッチ{batch_number} を処理中...")
            root.update_idletasks()
            
            df_extracted = pd.DataFrame()
            if not df_mail_data_batch.empty:
                df_extracted = extract_skills_data(df_mail_data_batch)

            conn = None
            # newly_saved_ids = [] # 📌 修正: マーク付けしないので不要
            
            try:
                conn = sqlite3.connect(db_path)
                
                # --- 1. 抽出データをDB保存 ---
                if not df_extracted.empty:
                    # ... (df_output の整形処理 ... は変更なし)
                    df_output = df_extracted.copy()
                    date_key_df = df_mail_data_batch[['EntryID', '受信日時']].copy()
                    if '受信日時' in df_output.columns:
                        df_output.drop(columns=['受信日時'], inplace=True, errors='ignore')
                    df_output = pd.merge(df_output, date_key_df, on='EntryID', how='left')
                    if 'EntryID' in df_output.columns and 'メールURL' not in df_output.columns:
                         df_output.insert(0, 'メールURL', df_output.apply(lambda row: f"outlook:{row['EntryID']}", axis=1))
                    df_output = reorder_output_dataframe(df_output)
                    final_drop_list = ['宛先メール', '本文(抽出元結合)'] 
                    final_drop_list = [col for col in df_output.columns if col in final_drop_list]
                    df_output = df_output.drop(columns=final_drop_list, errors='ignore')
                    if 'EntryID' not in df_output.columns:
                         raise KeyError("抽出結果バッチに EntryID が含まれていません。")
                    entry_ids_in_this_batch = df_output['EntryID'].tolist()
                    df_output.set_index('EntryID', inplace=True)
                    
                    try:
                        existing_ids_set = set(pd.read_sql_query("SELECT EntryID FROM emails", conn)['EntryID'].tolist())
                    except pd.io.sql.DatabaseError:
                        existing_ids_set = set() 
                        
                    new_ids = [eid for eid in entry_ids_in_this_batch if eid not in existing_ids_set]
                    df_new = df_output.loc[new_ids]
                    
                    if not df_new.empty:
                        df_new.to_sql('emails', conn, if_exists='append', index=True) 
                        newly_saved_count = len(df_new)
                        total_new_records_saved += newly_saved_count 
                        # newly_saved_ids = df_new.index.tolist() # 📌 修正: マーク付けしないので不要
                        print(f"INFO: {newly_saved_count} 件の新規レコードをDBに追記しました。(累計: {total_new_records_saved} 件)")
                        
                        db_flag = main_elements.get("db_has_new_data_var")
                        if db_flag: db_flag.set(True) 
                        
                        search_button = main_elements.get("search_button")
                        if search_button and str(search_button.cget('state')) == tk.DISABLED:
                            q = main_elements.get("gui_queue")
                            if q: q.put("ENABLE_SEARCH_BUTTON")

                # --- 2. スキップIDをDB保存 ---
                if skip_ids_batch:
                    try:
                        existing_skip_ids = set(pd.read_sql_query("SELECT EntryID FROM skipped_ids", conn)['EntryID'].tolist())
                    except pd.io.sql.DatabaseError:
                        existing_skip_ids = set()
                        conn.execute("CREATE TABLE IF NOT EXISTS skipped_ids (EntryID TEXT PRIMARY KEY)")
                    
                    # --- ▼▼▼【修正】バグ2対策: 重複を除外してからDB保存 ▼▼▼ ---
                    unique_ids_in_batch = set(skip_ids_batch)
                    new_skip_ids = [eid for eid in unique_ids_in_batch if eid not in existing_skip_ids] 
                    # --- ▲▲▲ 修正ここまで ▲▲▲ ---
                    
                    if new_skip_ids:
                         df_skip = pd.DataFrame(new_skip_ids, columns=['EntryID'])
                         df_skip.set_index('EntryID', inplace=True)
                         df_skip.to_sql('skipped_ids', conn, if_exists='append', index=True)
                         total_items_skipped += len(new_skip_ids)
                         print(f"INFO: {len(new_skip_ids)} 件のスキップIDをDBに追記しました。(累計: {total_items_skipped} 件)")
                         
                         # スキップIDが保存されただけでも「旗」を立てる
                         db_flag = main_elements.get("db_has_new_data_var")
                         if db_flag: db_flag.set(True)

            except Exception as e:
                print(f"❌ データベース書き込み中にエラー発生: {e}")
                messagebox.showerror("DB書込エラー", f"データベースへの書き込み中にエラーが発生しました。\n詳細: {e}")
            finally:
                if conn: conn.close()
            
            # --- 3.【重要】DB保存後/スキップ対象にマーク付け ---　削除
        # --- ループ終了後の最終処理 ---
        if total_new_records_saved == 0 and total_items_skipped == 0:
            status_label.config(text="状態: 処理対象のメールがありませんでした。")
            q = main_elements.get("gui_queue")
            if q: q.put("EXTRACTION_NO_ITEMS_FOUND")
            return 

        final_message = f"抽出処理が正常に完了しました。\n"
        if total_new_records_saved > 0:
             final_message += f"合計 {total_new_records_saved} 件の新規レコードが '{DATABASE_NAME}' に保存されました。\n"
        if total_items_skipped > 0:
             final_message += f"合計 {total_items_skipped} 件のメールが「スキップ対象」としてDBに記録されました。\n"
        # 📌 修正: total_items_marked のメッセージを削除
        q = main_elements.get("gui_queue")
        if q:
            q.put(f"EXTRACTION_COMPLETE:{total_new_records_saved}:{final_message}") 
        
        status_label.config(text=f"状態: 処理完了。{total_new_records_saved} 件保存済み。")
        
        search_button = main_elements.get("search_button")
        if search_button and total_new_records_saved > 0:
             search_button.config(state=tk.NORMAL)
        
    except Exception as e:
        status_label.config(text=f"状態: エラー発生 - {e}")
        error_message_for_user = f"抽出処理中に予期せぬエラーが発生しました。\n詳細: {e}"
        q = main_elements.get("gui_queue")
        if q:
            q.put(f"EXTRACTION_ERROR:{error_message_for_user}")
        traceback.print_exc()
        
    finally:
        q = main_elements.get("gui_queue")
        if q:
            q.put("EXTRACTION_COMPLETE_ENABLE_BUTTON") 
        
        # 📌 修正: Outlookオブジェクトの解放は不要
        # namespace = None
        # outlook_app = None
        
        pythoncom.CoUninitialize()
# ----------------------------------------------------
# 抽出ボタンコールバック (ボタン無効化)
# ----------------------------------------------------
def run_extraction_callback():
    run_button = main_elements.get("run_button")
    if run_button is None:
        print("警告: run_button が main_elements に見つかりません。") 
        return 
    if str(run_button.cget('state')) == tk.NORMAL:
        run_button.config(state=tk.DISABLED)
        run_extraction_thread(root, main_elements, main_elements["extract_days_var"])

def run_extraction_thread(root, main_elements, extract_days_var):
    account_email = main_elements["account_entry"].get().strip()
    folder_path = main_elements["folder_entry"].get().strip()
    status_label = main_elements["status_label"]
    read_mode = "unprocessed"
    read_days = extract_days_var.get()
    if not account_email or not folder_path:
        messagebox.showerror("入力エラー", "メールアカウントとフォルダパスの入力は必須です。")
        run_button = main_elements.get("run_button")
        if run_button:
            try:
                 if run_button.winfo_exists(): run_button.config(state=tk.NORMAL)
            except: pass
        return
    thread = threading.Thread(target=lambda: actual_run_extraction_logic(root, main_elements, account_email, folder_path, read_mode, read_days, status_label))
    thread.start()

# ----------------------------------------------------
# 削除処理ロジック
# ----------------------------------------------------
def delete_processed_records(days_ago: int, db_path: str, table_name: str) -> str:
    """
    指定された日数に基づき、指定されたテーブル内の古いレコードを削除する。
    """
    try:
        days_ago = int(days_ago)
        if days_ago < 0:
             raise ValueError("日数は0以上の整数で指定してください。")
    except ValueError:
        return f"エラー({table_name}): 日数設定が不正です (0以上の整数で指定)。"

    today = datetime.date.today()
    
    # 📌 'skipped_ids' テーブルには日付列がないため、0日(全削除)以外はスキップ
    if table_name == "skipped_ids" and days_ago != 0:
        return f"INFO(skipped_ids): スキップIDは日付指定削除の対象外です。"

    if days_ago == 0:
        where_clause = "" 
        target_message = f"テーブル '{table_name}' のすべての取り込み記録"
    else:
        # --- ▼▼▼【ロジック修正】「N日前も含む」ように変更 ▼▼▼ ---
        # (N-1)日前の0時をカットオフにする (例: N=1 なら 0日前(今日)の0時)
        cutoff_date = today - timedelta(days=(days_ago - 1)) 
        cutoff_datetime = datetime.datetime.combine(cutoff_date, datetime.time.min) 
        cutoff_str = cutoff_datetime.strftime('%Y-%m-%d %H:%M:%S')
        # (例: N=1 の場合、'今日'の0時より前 = 昨日 23:59 までが削除対象)
        where_clause = f"WHERE \"受信日時\" < '{cutoff_str}'"
        target_message = f"テーブル '{table_name}' の '{days_ago}日前以前' の記録"
        # --- ▲▲▲ 修正ここまで ▲▲▲ ---
    deleted_count = 0
    if not os.path.exists(db_path):
        return f"INFO({table_name}): データベースファイルが見つかりません。スキップします。"

    conn = None
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()

        count_sql = f"SELECT COUNT(*) FROM {table_name} {where_clause}"
        cursor.execute(count_sql)
        deleted_count = cursor.fetchone()[0]

        if deleted_count > 0:
            delete_sql = f"DELETE FROM {table_name} {where_clause}"
            cursor.execute(delete_sql)
            conn.commit() 
            return f"{target_message} ({deleted_count}件) を削除しました。"
        else:
            return f"{target_message} は見つかりませんでした。削除は行われませんでした。"
    except sqlite3.Error as e: 
        if conn: conn.rollback() 
        if "no such table" in str(e): 
             return f"INFO({table_name}): テーブルが見つかりません。スキップします。"
        return f"エラー({table_name}): DBファイルの処理中にエラーが発生しました ({e})"
    except Exception as e: 
         if conn: conn.rollback()
         return f"エラー({table_name}): 予期せぬエラーが発生しました ({e})"
    finally:
        if conn: conn.close()

def run_deletion_thread(root, main_elements):
    thread = threading.Thread(target=lambda: actual_run_file_deletion_logic(main_elements))
    thread.start()

def actual_run_file_deletion_logic(main_elements):
    
    # 📌 修正: Outlookに接続しないので COM初期化は不要
    # try:
    #     pythoncom.CoInitialize() 
    # except Exception:
    #     pass 
    
    try:
        days_entry = main_elements["delete_days_entry"] 
        status_label = main_elements["status_label"]
        # 📌 修正: reset_category_var は削除 (GUIからも削除済み想定)
        # reset_category_var = main_elements["reset_category_var"]
        days_input = days_entry.get().strip()
        db_path = os.path.abspath(DATABASE_NAME) 

        try:
            days_ago = int(days_input)
            if days_ago < 0: 
                raise ValueError("日数は0以上の整数を指定してください。")
        except ValueError as e:
            messagebox.showerror("入力エラー", f"削除日数の入力が不正です: {e}\n(0以上の整数で指定)")
            status_label.config(text="状態: 削除失敗 (入力不正)。")
            return 

        # 📌 修正: reset_category_flag は削除
        # reset_category_flag = reset_category_var.get()

        if days_ago == 0:
             confirm_prompt = f"🚨 **警告:** データベース内の**すべてのレコード**（抽出データとスキップ履歴）を削除します。\n"
        else:
             confirm_prompt = f"🚨 **警告:** データベース内の **{days_ago}日より古いレコード**を削除します。\n"
        
        # 📌 修正: カテゴリ解除のチェックボックスに関連するロジックを削除
        # if reset_category_flag:
        #     ...
        # else:
        #     confirm_prompt += "\n**本当に実行しますか？**"
        confirm_prompt += "\n**本当に実行しますか？**"

        confirm = messagebox.askyesno("最終確認", confirm_prompt, icon='warning')
        if not confirm:
            status_label.config(text="状態: 削除処理キャンセル。")
            return 

        status_label.config(text=f"状態: DBレコード削除試行中...")
        root = status_label.winfo_toplevel() 
        root.update_idletasks() 

        db_exists = os.path.exists(db_path) 
        delete_result_message = "" 
        delete_result_message_skipped = "" # スキップID用
        db_processed = False 
        db_had_error = False 

        if db_exists:
            try:
                # 抽出済み(emails)テーブルから削除
                delete_result_message = delete_processed_records(days_ago, db_path, "emails")
                # スキップ済み(skipped_ids)テーブルから削除
                delete_result_message_skipped = delete_processed_records(days_ago, db_path, "skipped_ids")
                
                db_processed = True 
                if "エラー:" in delete_result_message or "エラー:" in delete_result_message_skipped:
                    db_had_error = True 
                    status_label.config(text="状態: DB削除エラー。")
                
            except NameError:
                 messagebox.showerror("内部エラー", "レコード削除関数(delete_processed_records)が見つかりません。")
                 status_label.config(text="状態: 内部エラー。")
                 return
            except Exception as db_del_err:
                 delete_result_message = f"DBレコード削除中に予期せぬエラーが発生しました。\n{db_del_err}" 
                 db_had_error = True
                 messagebox.showerror("DB削除エラー", delete_result_message) 
                 status_label.config(text="状態: DB削除エラー。")
                 db_processed = True 
        else:
            delete_result_message = f"データベースファイル '{os.path.basename(db_path)}' が見つかりませんでした。DBレコード削除はスキップされました。"

        # 📌 修正: カテゴリ解除の処理ブロックをすべて削除
        
        final_msg = delete_result_message + "\n" + delete_result_message_skipped.replace("INFO: ", "")
        
        # 📌 修正: カテゴリ解除のメッセージを削除
                 
        msg_title = "処理完了"
        msg_icon = 'info'
        final_status_text = "状態: 削除処理完了。"
        
        # 📌 修正: カテゴリ解除のエラー分岐を削除
        if db_had_error:
             msg_title = "処理完了 (DB削除エラー)"
             msg_icon = 'warning' 
             final_status_text = "状態: DB削除エラー。"
        elif "INFO:" in delete_result_message:
             msg_title = "処理スキップ"
             msg_icon = 'info'
             final_status_text = "状態: DBファイルなし。"
             
        if msg_icon == 'info': messagebox.showinfo(msg_title, final_msg)
        elif msg_icon == 'warning': messagebox.showwarning(msg_title, final_msg)
        status_label.config(text=final_status_text) 
    
    except Exception as outer_err:
         try:
              status_label.config(text="状態: 削除スレッドで重大なエラー。")
         except: pass 
         
    finally:
        pass
        # 📌 修正: Outlookに接続しないので COM解放は不要

# ----------------------------------------------------
# メイン実行関数 (GUI起動)
# ----------------------------------------------------
def main():
    global root, main_elements
    
    root = tk.Tk()
    root.title("Outlook Mail Search Tool")
    window_width = 800
    window_height = 600 
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    center_x = int(screen_width/2 - window_width/2)
    center_y = int(screen_height/2 - window_height/2)
    root.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
    
    def on_main_window_close():
        root.destroy() 
    root.protocol("WM_DELETE_WINDOW", on_main_window_close)

    delete_days_var = tk.StringVar(value="14") 
    extract_days_var = tk.StringVar(value="1") 
    # reset_category_var = tk.BooleanVar(value=False) # 📌 修正: 削除
    gui_queue = queue.Queue()
    db_has_new_data_var = tk.BooleanVar(value=False)
    
    saved_account, saved_folder = utils.load_config_csv() 
    if not saved_folder: saved_folder = TARGET_FOLDER_PATH 

    main_frame = Frame(root)
    main_frame.pack(padx=10, pady=10, fill='both', expand=True)
    
    top_button_frame = ttk.Frame(main_frame)
    top_button_frame.pack(fill='x', padx=10, pady=(10, 0))
    top_button_frame.grid_columnconfigure(0, weight=1) 
    top_button_frame.grid_columnconfigure(1, weight=0) 
    settings_button = ttk.Button(top_button_frame, text="⚙ 設定")
    settings_button.grid(row=0, column=1, padx=(0, 5), pady=5, sticky='e')

    setting_frame = ttk.LabelFrame(main_frame, text="アカウント/フォルダ設定")
    setting_frame.pack(padx=10, pady=(0, 10), fill='x')
    setting_frame.grid_columnconfigure(1, weight=1)
    ttk.Label(setting_frame, text="アカウントメール:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
    account_entry = ttk.Entry(setting_frame, width=50)
    account_entry.insert(0, saved_account)
    account_entry.grid(row=0, column=1, padx=5, pady=5, sticky='ew')
    ttk.Label(setting_frame, text="対象フォルダパス:").grid(row=1, column=0, padx=5, pady=5, sticky='w')
    folder_entry = ttk.Entry(setting_frame, width=50)
    folder_entry.insert(0, saved_folder)
    folder_entry.grid(row=1, column=1, padx=5, pady=5, sticky='ew')
    
    process_frame = ttk.LabelFrame(main_frame, text="メールデータ抽出/検索")
    process_frame.pack(padx=10, pady=10, fill='x')
    process_frame.grid_columnconfigure(0, weight=1)
    process_frame.grid_columnconfigure(1, weight=1)
    days_frame = ttk.Frame(process_frame)
    days_frame.grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky='w')
    ttk.Label(days_frame, text="未処理メールの検索期間 (N日前まで):").pack(side=tk.LEFT)
    extract_days_entry = ttk.Entry(days_frame, textvariable=extract_days_var, width=10)
    extract_days_entry.pack(side=tk.LEFT, padx=5)
    ttk.Label(days_frame, text="日 (0=今日, 空欄=全期間)").pack(side=tk.LEFT)
    run_button = ttk.Button(process_frame, text="抽出実行") 
    run_button.grid(row=1, column=0, padx=5, pady=5, sticky='ew')
    search_button = ttk.Button(process_frame, text="検索一覧 (結果表示)", state=tk.DISABLED)
    search_button.grid(row=1, column=1, padx=5, pady=5, sticky='ew')
    
    delete_frame = ttk.LabelFrame(main_frame, text="メール/レコード管理")
    delete_frame.pack(padx=10, pady=(10, 5), fill='x')
    delete_frame.grid_columnconfigure(0, weight=0)
    delete_frame.grid_columnconfigure(1, weight=0)
    delete_frame.grid_columnconfigure(2, weight=0)
    delete_frame.grid_columnconfigure(3, weight=1) 
    # --- ▼▼▼【修正案】▼▼▼ ---
    ttk.Label(delete_frame, text="N日前以前のレコードを削除:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
    
    delete_days_entry = ttk.Entry(delete_frame, textvariable=delete_days_var, width=10)
    delete_days_entry.grid(row=0, column=1, padx=5, pady=5, sticky='w') 
    
    # (補足説明も変更)
    ttk.Label(delete_frame, text="日 (0=全削除, 1=昨日も含む)").grid(row=0, column=2, padx=(0, 10), pady=5, sticky='w') 
    # --- ▲▲▲ 修正ここまで ▲▲▲ ---
    delete_button = ttk.Button(delete_frame, text="レコード削除実行")
    delete_button.grid(row=1, column=0, columnspan=4, padx=5, pady=5, sticky='ew') 
    
    # --- ▼▼▼【修正】カテゴリ解除のチェックボックスを削除 ▼▼▼ ---
    # reset_category_checkbox = ttk.Checkbutton(
    #     delete_frame, 
    #     text="処理済みマークを解除する", 
    #     variable=reset_category_var
    # )
    # reset_category_checkbox.grid(row=2, column=0, columnspan=4, padx=5, pady=(15, 5), sticky='w') 
    # --- ▲▲▲ 削除ここまで ▲▲▲ ---
    
    status_label = ttk.Label(main_frame, text="状態: 待機中", relief=tk.SUNKEN, anchor='w')
    status_label.pack(side=tk.BOTTOM, fill='x', padx=10, pady=(5, 0))
    
    # --- グローバル辞書にキーを追加 ---
    main_elements["account_entry"] = account_entry
    main_elements["folder_entry"] = folder_entry
    main_elements["status_label"] = status_label
    main_elements["search_button"] = search_button
    main_elements["delete_days_entry"] = delete_days_entry
    main_elements["extract_days_entry"] = extract_days_entry
    main_elements["settings_button"] = settings_button
    # main_elements["reset_category_var"] = reset_category_var # 📌 修正: 削除
    main_elements["extract_days_var"] = extract_days_var
    main_elements["run_button"] = run_button
    main_elements["gui_queue"] = gui_queue
    main_elements["db_has_new_data_var"] = db_has_new_data_var
    
    settings_button.config(command=open_settings_callback)
    run_button.config(command=run_extraction_callback)
    search_button.config(command=open_search_callback)
    delete_button.config(command=lambda: run_deletion_thread(root, main_elements))

    output_file_abs_path = os.path.abspath(DATABASE_NAME) 
    
    if os.path.exists(output_file_abs_path):
        search_button.config(state=tk.NORMAL)
        status_label.config(text="状態: 抽出結果ファイルあり。検索一覧が利用可能です。")

    # --- 起動時の未処理メールチェック (COM初期化追加 + チェッカー) ---
    def check_unprocessed_async(account_email, folder_path, q, initial_days_value):
        thread_id = threading.get_ident()
        try:
            pythoncom.CoInitialize()
        except Exception:
            pass
        
        try: 
            output_path_exists = os.path.exists(output_file_abs_path)
            days_to_check_val = None
            try:
                if initial_days_value is not None and str(initial_days_value).strip():
                     days_to_check_val = int(initial_days_value) 
                     if days_to_check_val < 0:
                          days_to_check_val = None 
            except (ValueError, TypeError):
                 days_to_check_val = None 

            try:
                # 📌 修正: DB基準の has_unprocessed_mail を呼び出す
                unprocessed_count = has_unprocessed_mail(folder_path, account_email, days_to_check=days_to_check_val)
                
                if unprocessed_count > 0:
                    final_message = f"状態: {unprocessed_count}件の未処理メールがあります (直近300件を確認)。"
                else:
                    if output_path_exists:
                        final_message = "状態: 抽出結果ファイルあり。未処理メールはありません。"
                    else:
                        final_message = "状態: 対象のメールはありません" 
                q.put(final_message)
            except Exception as e:
                error_msg = f"状態: バックグラウンドチェックエラー - {e}"
                q.put(error_msg)
                if not output_path_exists:
                    q.put("状態: 待機中（チェックエラー）。")
                    
        except Exception as outer_err:
             q.put("状態: 未処理チェックで重大なエラー。")
             
        finally:
             pythoncom.CoUninitialize()
             
    def check_queue():
        try:
            message = gui_queue.get(block=False)
            
            if message == "EXTRACTION_COMPLETE_ENABLE_BUTTON":
                run_button = main_elements.get("run_button")
                if run_button:
                    try:
                        if run_button.winfo_exists():
                            run_button.config(state=tk.NORMAL)
                    except tk.TclError:
                        pass 
            
            elif message == "ENABLE_SEARCH_BUTTON":
                search_button = main_elements.get("search_button")
                if search_button:
                    try:
                        if search_button.winfo_exists():
                            search_button.config(state=tk.NORMAL)
                            status_label.config(text="状態: DBにデータが保存されました。検索可能です。")
                    except tk.TclError:
                        pass
            
            elif message.startswith("EXTRACTION_COMPLETE:"):
                try:
                    parts = message.split(":", 2)
                    total_saved = parts[1]
                    final_message = parts[2]
                    search_window = main_elements.get("search_window")
                    active_parent = search_window if (search_window and search_window.winfo_exists()) else root
                    messagebox.showinfo("完了", final_message, parent=active_parent)
                    status_label.config(text=f"状態: 処理完了。{total_saved} 件保存済み。")
                except Exception as e:
                    print(f"完了メッセージの表示エラー: {e}")

            elif message == "EXTRACTION_NO_ITEMS_FOUND":
                search_window = main_elements.get("search_window")
                active_parent = search_window if (search_window and search_window.winfo_exists()) else root
                messagebox.showinfo("完了", "処理対象のメールがありませんでした。", parent=active_parent)
                status_label.config(text="状態: 処理対象のメールがありませんでした。")

            elif message.startswith("EXTRACTION_ERROR:"):
                error_details = message.split(":", 1)[1]
                search_window = main_elements.get("search_window")
                active_parent = search_window if (search_window and search_window.winfo_exists()) else root
                messagebox.showerror("エラー", error_details, parent=active_parent)
                status_label.config(text="状態: エラー発生。")

            else:
                status_label.config(text=message)
                 
        except queue.Empty:
            pass
        finally:
            try:
                 if root and root.winfo_exists(): root.after(100, check_queue)
            except tk.TclError: pass

    initial_extract_days = None
    if "extract_days_var" in main_elements:
         try: initial_extract_days = main_elements["extract_days_var"].get()
         except tk.TclError: pass 
              
    threading.Thread(target=lambda: check_unprocessed_async(saved_account, saved_folder, gui_queue, initial_extract_days), daemon=True).start()
    
    root.after(100, check_queue)
    root.mainloop()

# ----------------------------------------------------
# 外部コールバック (main 関数外に移動)
# ----------------------------------------------------
def open_settings_callback():
    if root and main_elements:
        gui_elements.open_settings_window(
            root, main_elements["account_entry"], main_elements["status_label"]
        )

def open_search_callback():
    if not root or not main_elements: return
    
    db_path = os.path.abspath(DATABASE_NAME)
    if not os.path.exists(db_path):
        messagebox.showwarning("警告", f"データベース ('{DATABASE_NAME}') が見つかりません。\n先に抽出を実行してください。")
        return
        
    try:
        root.withdraw() 
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='emails';")
        if cursor.fetchone() is None:
             conn.close()
             messagebox.showerror("エラー", f"データベースに 'emails' テーブルが見つかりません。")
             try: root.deiconify()
             except tk.TclError: pass
             return
             
        # --- フリーズ対策: 軽量読み込み (ORDER BY 追加) ---
        cursor.execute("PRAGMA table_info(emails)")
        all_columns = [info[1] for info in cursor.fetchall()]
        heavy_columns = ['本文(テキスト形式)', '本文(ファイル含む)']
        light_columns = [col for col in all_columns if col not in heavy_columns]
        if not light_columns:
             conn.close()
             messagebox.showerror("エラー", "データベースの列構造が不明か、主要な列がありません。")
             try: root.deiconify()
             except tk.TclError: pass
             return
        light_columns_sql = ", ".join([f'"{col}"' for col in light_columns])
        query = f"SELECT {light_columns_sql} FROM emails ORDER BY \"受信日時\" DESC"
        df_for_gui = pd.read_sql_query(query, conn)
        conn.close()
        
        # --- 「旗」の処理 (検索一覧が開くことを記録) ---
        db_flag = main_elements.get("db_has_new_data_var")
        if db_flag:
            db_flag.set(False) 

        search_app = gui_search_window.App(
            root, 
            data_frame=df_for_gui,
            open_email_callback=open_outlook_email_by_id,
            db_has_new_data_var=db_flag 
        ) 
        
        main_elements["search_window"] = search_app 
        
        search_app.wait_window() 
        
    except Exception as e:
        messagebox.showerror("検索ウィンドウ起動エラー", f"検索一覧の表示中に予期せぬエラーが発生しました。\n詳細: {e}")
        traceback.print_exc()
    finally:
         main_elements["search_window"] = None 
         try:
             if root and root.winfo_exists():
                  root.deiconify() 
         except tk.TclError:
              pass 
         except Exception as e_final:
              print(f"警告: メインウィンドウ復元中に予期せぬエラー: {e_final}")


if __name__ == "__main__":
    main()