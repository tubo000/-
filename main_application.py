# main_application.py (ログ出力削除・COM初期化修正版)
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
from email_processor import get_mail_data_from_outlook_in_memory, DATABASE_NAME 
from email_processor import has_unprocessed_mail 
from email_processor import remove_processed_category, PROCESSED_CATEGORY_NAME

root = None
main_elements = {}
# ----------------------------------------------------
# ユーティリティ関数群 (Outlook連携、DF処理)
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
# 抽出処理ロジック (COM初期化は維持 + 0日対応)
# ----------------------------------------------------
def actual_run_extraction_logic(root, main_elements, target_email, folder_path, read_mode, read_days, status_label):
    
    thread_id = threading.get_ident()
    
    try:
        pythoncom.CoInitialize()
    except Exception:
        pass 
    # --- ▼▼▼【この1行を追加】▼▼▼ ---
    total_new_records_saved = 0 # DBに保存した総件数を初期化
    # --- ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲ ---     
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
                return # finally が実行される

        if days_ago == 0:
             mode_text = "未処理 (今日のみ)"
        elif days_ago is not None and days_ago > 0 :
             mode_text = f"未処理 (過去{days_ago}日)"
        else:
             mode_text = "未処理 (全期間)"
            
        status_label.config(text=f"状態: {target_email} アカウントからメール取得中 ({mode_text})...")
        root.update_idletasks() 

        df_mail_data = get_mail_data_from_outlook_in_memory(
            folder_path, 
            target_email, 
            read_mode=read_mode, 
            days_ago=days_ago,
            main_elements=main_elements # 📌 キューを渡すために main_elements を渡す
        )
        # --- ▼▼▼【テキスト追加】▼▼▼ ---
        status_label.config(text="状態: 抽出コアロジック実行中...")
        root.update_idletasks()
        # --- ▲▲▲ テキスト追加 ▲▲▲ ---
        # バッチごとにループ
        for df_mail_data_batch in df_mail_data:
            
            if df_mail_data_batch.empty:
                # このバッチでは新規抽出がなかった
                status_label.config(text="状態: 処理対象のメールがありませんでした。")
                # (email_processor 側で5秒待機は実行済み)
                continue 

            # バッチ処理中は「N件を整形・保存中」と表示（「コアロジック実行中」を上書き）
            status_label.config(text=f"状態: {len(df_mail_data_batch)}件のデータをDB保存中...")
            root.update_idletasks()
            df_extracted = extract_skills_data(df_mail_data_batch)
            
            # --- データベース書き込み処理 (バッチごと) ---
            df_output = df_extracted.copy()
            
            # df_mail_data_batch から EntryID と 受信日時 をマージする
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
            
            db_path = os.path.abspath(DATABASE_NAME) 
            conn = None
            try:
                conn = sqlite3.connect(db_path)
                if 'EntryID' not in df_output.columns:
                     raise KeyError("抽出結果バッチに EntryID が含まれていません。")
                     
                entry_ids_in_this_batch = df_output['EntryID'].tolist()
                df_output.set_index('EntryID', inplace=True)
                
                # DBから「最新」の既存IDを読み込む (バッチごとに)
                try:
                    existing_ids_set = set(pd.read_sql_query("SELECT EntryID FROM emails", conn)['EntryID'].tolist())
                except pd.io.sql.DatabaseError:
                    existing_ids_set = set() 
                    
                new_ids = [eid for eid in entry_ids_in_this_batch if eid not in existing_ids_set]
                df_new = df_output.loc[new_ids]
                
                if not df_new.empty:
                    df_new.to_sql('emails', conn, if_exists='append', index=True) 
                    newly_saved_count = len(df_new)
                    total_new_records_saved += newly_saved_count # 総件数を加算
                    print(f"INFO: {newly_saved_count} 件の新規レコードをDBに追記しました。(累計: {total_new_records_saved} 件)")
                    # --- ▼▼▼【ここに追加】(4行) ▼▼▼ ---
                    # 「新しいデータがあるよ」という旗を立てる
                    db_flag = main_elements.get("db_has_new_data_var")
                    if db_flag:
                        db_flag.set(True) 
                    # --- ▲▲▲ 追加ここまで ▲▲▲ ---
                    # --- ▼▼▼【ここが修正箇所】▼▼▼ ---
                    # 最初のDB書き込みが成功したら、検索ボタンを有効化するよう通知
                    search_button = main_elements.get("search_button")
                    # ボタンがまだ無効 (DISABLED) の場合のみ、有効化を依頼
                    if search_button and str(search_button.cget('state')) == tk.DISABLED:
                        q = main_elements.get("gui_queue")
                        if q:
                            q.put("ENABLE_SEARCH_BUTTON") # 検索ボタン有効化を依頼
                    # --- ▲▲▲ 追加ここまで ▲▲▲ ---
                
            except Exception as e:
                print(f"❌ データベース書き込み中にエラー発生: {e}")
                messagebox.showerror("DB書込エラー", f"データベースへの書き込み中にエラーが発生しました。\n詳細: {e}")
            finally:
                if conn: conn.close()
            # --- データベース書き込みここまで (バッチごと) ---
        
        # --- ▲▲▲ 修正ここまで ▲▲▲ ---
        
        
        # --- ループ終了後の最終処理 ---
        if total_new_records_saved == 0:
            status_label.config(text="状態: 処理対象のメールがありませんでした。")
            
            # --- ▼▼▼【修正】ポップアップの代わりにキューに送る ▼▼▼ ---
            q = main_elements.get("gui_queue")
            if q:
                q.put("EXTRACTION_NO_ITEMS_FOUND")
            # --- ▲▲▲ 修正ここまで ▲▲▲ ---
            return # finally が実行される

        # --- ▼▼▼【修正】ポップアップの代わりにキューに送る ▼▼▼ ---
        # (成功メッセージもキューに送る)
        final_message = f"抽出処理が正常に完了し、\n合計 {total_new_records_saved} 件の新規レコードが '{DATABASE_NAME}' に保存されました。"
        q = main_elements.get("gui_queue")
        if q:
            # 完了メッセージと件数を送る
            q.put(f"EXTRACTION_COMPLETE:{total_new_records_saved}:{final_message}") 
        # --- ▲▲▲ 修正ここまで ▲▲▲ ---

        # (古い messagebox.showinfo は削除)
        
        status_label.config(text=f"状態: 処理完了。{total_new_records_saved} 件保存済み。")
        search_button = main_elements.get("search_button")
        if search_button: search_button.config(state=tk.NORMAL)
        
    except Exception as e:
        status_label.config(text=f"状態: エラー発生 - {e}")
        # --- ▼▼▼【修正】エラーもキューに送る ▼▼▼ ---
        error_message_for_user = f"抽出処理中に予期せぬエラーが発生しました。\n詳細: {e}"
        q = main_elements.get("gui_queue")
        if q:
            q.put(f"EXTRACTION_ERROR:{error_message_for_user}")
        # --- ▲▲▲ 修正ここまで ▲▲▲ ---
        traceback.print_exc()
        
    finally:
        q = main_elements.get("gui_queue")
        if q:
            q.put("EXTRACTION_COMPLETE_ENABLE_BUTTON") # ボタン有効化 (これは変更なし)
        
        pythoncom.CoUninitialize()
# ----------------------------------------------------
# 抽出ボタンコールバック (ボタン無効化)
# ----------------------------------------------------
def run_extraction_callback():
    run_button = main_elements.get("run_button")
    
    # 1回目のクリックで None になる問題の対策
    if run_button is None:
        print("警告: run_button が main_elements に見つかりません。") 
        return 

    if str(run_button.cget('state')) == tk.NORMAL:
        run_button.config(state=tk.DISABLED)
        # print("INFO: 抽出実行ボタンを無効化。処理開始...") # ログ削除
        run_extraction_thread(root, main_elements, main_elements["extract_days_var"])
    # else: # ログ削除
    #     print(f"INFO: 抽出処理が既に実行中か、ボタンが無効です。")

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
                 # print("INFO: 入力エラーのため抽出実行ボタンを有効化しました。") # ログ削除
            except: pass
        return

    thread = threading.Thread(target=lambda: actual_run_extraction_logic(root, main_elements, account_email, folder_path, read_mode, read_days, status_label))
    thread.start()

# ----------------------------------------------------
# 削除処理ロジック (COM初期化追加)
# ----------------------------------------------------
def run_deletion_thread(root, main_elements):
    thread = threading.Thread(target=lambda: actual_run_file_deletion_logic(main_elements))
    thread.start()

# ----------------------------------------------------------------------
# 💡 【最終版】 レコード削除関数 (SQLite専用)
# ----------------------------------------------------------------------
def delete_processed_records(days_ago: int, db_path: str) -> str:
    try:
        days_ago = int(days_ago)
        if days_ago < 0:
             raise ValueError("日数は0以上の整数で指定してください。")
    except ValueError:
        return "エラー: 日数設定が不正です (0以上の整数で指定)。"
    today = datetime.date.today()
    if days_ago == 0:
        where_clause = "" 
        target_message = "すべての取り込み記録"
    else:
        cutoff_date = today - timedelta(days=days_ago)
        cutoff_datetime = datetime.datetime.combine(cutoff_date, datetime.time.min) 
        cutoff_str = cutoff_datetime.strftime('%Y-%m-%d %H:%M:%S')
        where_clause = f"WHERE \"受信日時\" < '{cutoff_str}'"
        target_message = f"'{cutoff_date.strftime('%Y年%m月%d日')}' より古い取り込み記録"
    deleted_count = 0
    if not os.path.exists(db_path):
        return f"INFO: データベースファイルが見つかりません ({os.path.basename(db_path)})。スキップします。"
    conn = None
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        count_sql = f"SELECT COUNT(*) FROM emails {where_clause}"
        cursor.execute(count_sql)
        deleted_count = cursor.fetchone()[0]
        if deleted_count > 0:
            delete_sql = f"DELETE FROM emails {where_clause}"
            cursor.execute(delete_sql)
            conn.commit() 
            return f"{target_message} ({deleted_count}件) を削除しました。"
        else:
            return f"{target_message} は見つかりませんでした。削除は行われませんでした。"
    except sqlite3.Error as e: 
        if conn: conn.rollback() 
        # print(f"❌ DBエラー発生: {e}") # ログ削除
        return f"エラー: DBファイルの処理中にエラーが発生しました ({e})"
    except Exception as e: 
         if conn: conn.rollback()
         # print(f"❌ 予期せぬエラー発生: {e}") # ログ削除
         return f"エラー: 予期せぬエラーが発生しました ({e})"
    finally:
        if conn: conn.close()

# ----------------------------------------------------
# 削除スレッド本体 (COM初期化追加 + 0日対応)
# ----------------------------------------------------
def actual_run_file_deletion_logic(main_elements):
    
    try:
        pythoncom.CoInitialize() 
    except Exception:
        pass 
    
    try:
        days_entry = main_elements["delete_days_entry"] 
        status_label = main_elements["status_label"]
        reset_category_var = main_elements["reset_category_var"]
        days_input = days_entry.get().strip()
        db_path = os.path.abspath(DATABASE_NAME) 

        try:
            days_ago = int(days_input)
            if days_ago < 0: 
                raise ValueError("日数は0以上の整数を指定してください。")
        except ValueError as e:
            messagebox.showerror("入力エラー", f"削除日数の入力が不正です: {e}\n(0以上の整数で指定)")
            status_label.config(text="状態: 削除失敗 (入力不正)。")
            return # finally が実行される

        reset_category_flag = reset_category_var.get()

        if days_ago == 0:
             confirm_prompt = f"🚨 **警告:** データベース内の**すべてのレコード**を削除します。\n"
        else:
             confirm_prompt = f"🚨 **警告:** データベース内の **{days_ago}日より古いレコード**を削除します。\n"
        if reset_category_flag:
            if days_ago == 0:
                 confirm_prompt += f"また、Outlookメールの『{PROCESSED_CATEGORY_NAME}』マークを**すべて解除**します。\n\n**本当に実行しますか？**"
            else:
                 confirm_prompt += f"また、Outlookメールの『{PROCESSED_CATEGORY_NAME}』マークを **{days_ago}日より古いメールから解除**します。\n\n**本当に実行しますか？**"
        else:
            confirm_prompt += "\n**本当に実行しますか？**"

        confirm = messagebox.askyesno("最終確認", confirm_prompt, icon='warning')
        if not confirm:
            status_label.config(text="状態: 削除処理キャンセル。")
            return # finally が実行される

        status_label.config(text=f"状態: DBレコード削除試行中...")
        root = status_label.winfo_toplevel() 
        root.update_idletasks() 

        db_exists = os.path.exists(db_path) 
        delete_result_message = "" 
        db_processed = False 
        db_had_error = False 

        if db_exists:
            try:
                delete_result_message = delete_processed_records(days_ago, db_path)
                db_processed = True 
                if "エラー:" in delete_result_message:
                    db_had_error = True 
                    status_label.config(text="状態: DB削除エラー。")
                elif "INFO:" in delete_result_message:
                     pass # print(f"INFO: {delete_result_message}") # ログ削除
                else:
                     pass # print(f"INFO: {delete_result_message}") # ログ削除
            except NameError:
                 messagebox.showerror("内部エラー", "レコード削除関数(delete_processed_records)が見つかりません。")
                 status_label.config(text="状態: 内部エラー。")
                 return # finally が実行される
            except Exception as db_del_err:
                 delete_result_message = f"DBレコード削除中に予期せぬエラーが発生しました。\n{db_del_err}" 
                 db_had_error = True
                 messagebox.showerror("DB削除エラー", delete_result_message) 
                 status_label.config(text="状態: DB削除エラー。")
                 db_processed = True 
        else:
            delete_result_message = f"データベースファイル '{os.path.basename(db_path)}' が見つかりませんでした。DBレコード削除はスキップされました。"
            # print(f"INFO: {delete_result_message}") # ログ削除

        reset_count = 0
        category_reset_error = None
        if reset_category_flag:
            status_label.config(text=f"状態: Outlookカテゴリ解除中...")
            root.update_idletasks()
            try:
                reset_days_param = None if days_ago == 0 else days_ago 
                reset_count = remove_processed_category(
                    main_elements["account_entry"].get().strip(),
                    main_elements["folder_entry"].get().strip(),
                    days_ago=reset_days_param
                )
                # print(f"INFO: Outlookカテゴリリセット {reset_count} 件完了。") # ログ削除
            except NameError:
                 category_reset_error = "カテゴリ解除関数(remove_processed_category)が見つかりません。"
                 # print(f"❌ {category_reset_error}") # ログ削除
                 status_label.config(text="状態: カテゴリ解除エラー (内部エラー)。")
            except Exception as e:
                 category_reset_error = f"Outlookカテゴリの解除中にエラーが発生しました。\n詳細: {e}"
                 # print(f"❌ {category_reset_error}") # ログ削除
                 status_label.config(text="状態: カテゴリ解除エラー。")

        final_msg = delete_result_message 
        if reset_category_flag:
            final_msg += f"\nOutlookカテゴリリセット: {reset_count} 件完了"
                 
        msg_title = "処理完了"
        msg_icon = 'info'
        final_status_text = "状態: 削除処理完了。"
        if category_reset_error:
             final_msg += f"\n\n**警告:** {category_reset_error}"
             msg_title = "処理完了 (カテゴリ解除エラー)"
             msg_icon = 'warning'
             if db_had_error: final_status_text = "状態: DB削除エラー、カテゴリ解除エラー。"
             elif not db_exists: final_status_text = "状態: カテゴリ解除エラー (DBスキップ)。"
             else: final_status_text = "状態: DB削除完了、カテゴリ解除エラー。"
        elif db_had_error:
             msg_title = "処理完了 (DB削除エラー)"
             msg_icon = 'warning' 
             final_status_text = "状態: DB削除エラー。"
        elif not db_exists and reset_category_flag:
             msg_title = "処理完了 (カテゴリ解除のみ)"
             final_status_text = "状態: カテゴリ解除完了 (DBスキップ)。"
        elif "INFO:" in delete_result_message and not reset_category_flag:
             msg_title = "処理スキップ"
             msg_icon = 'info'
             final_status_text = "状態: DBファイルなし。"
             
        if msg_icon == 'info': messagebox.showinfo(msg_title, final_msg)
        elif msg_icon == 'warning': messagebox.showwarning(msg_title, final_msg)
        status_label.config(text=final_status_text) 
    
    except Exception as outer_err:
         # print(f"❌ 削除スレッド全体で予期せぬエラー: {outer_err}\n{traceback.format_exc()}") # ログ削除
         try:
              status_label.config(text="状態: 削除スレッドで重大なエラー。")
         except: pass 
         
    finally:
        pythoncom.CoUninitialize() 

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
    reset_category_var = tk.BooleanVar(value=False) 
    gui_queue = queue.Queue()
    # --- ▼▼▼【ここに追加】(2行) ▼▼▼ ---
    db_has_new_data_var = tk.BooleanVar(value=False) # DBが更新されたかの「旗」
    # --- ▲▲▲ 追加ここまで ▲▲▲ ---
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
    ttk.Label(delete_frame, text="N日前より古いレコード削除:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
    delete_days_entry = ttk.Entry(delete_frame, textvariable=delete_days_var, width=10)
    delete_days_entry.grid(row=0, column=1, padx=5, pady=5, sticky='w') 
    ttk.Label(delete_frame, text="日 (0=全削除)").grid(row=0, column=2, padx=(0, 10), pady=5, sticky='w') 
    delete_button = ttk.Button(delete_frame, text="レコード削除実行")
    delete_button.grid(row=1, column=0, columnspan=4, padx=5, pady=5, sticky='ew') 
    reset_category_checkbox = ttk.Checkbutton(
        delete_frame, 
        text="処理済みマークを解除する", 
        variable=reset_category_var
    )
    reset_category_checkbox.grid(row=2, column=0, columnspan=4, padx=5, pady=(15, 5), sticky='w') 
    
    status_label = ttk.Label(main_frame, text="状態: 待機中", relief=tk.SUNKEN, anchor='w')
    status_label.pack(side=tk.BOTTOM, fill='x', padx=10, pady=(5, 0))
    
    main_elements = {
        "account_entry": account_entry,
        "folder_entry": folder_entry,
        "status_label": status_label,
        "search_button": search_button,
        "delete_days_entry": delete_days_entry, 
        "extract_days_entry": extract_days_entry,
        "settings_button": settings_button, 
        "reset_category_var": reset_category_var, 
        "extract_days_var": extract_days_var,
        "run_button": run_button,
        "gui_queue": gui_queue,
        # --- ▼▼▼【ここに追加】(1行) ▼▼▼ ---
        "db_has_new_data_var": db_has_new_data_var
        # --- ▲▲▲ 追加ここまで ▲▲▲ ---
    }
    
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
        # print(f"\n[CHECKER] Thread {thread_id} (Async Check) STARTING...") # ログ削除
        try:
            pythoncom.CoInitialize()
            # print(f"[CHECKER] Thread {thread_id} (Async Check) CoInitialize() CALLED.") # ログ削除
        except Exception:
            pass
        
        try: 
            output_path_exists = os.path.exists(output_file_abs_path)
            days_to_check_val = None
            try:
                if initial_days_value is not None and str(initial_days_value).strip():
                     days_to_check_val = int(initial_days_value) 
                     if days_to_check_val < 0:
                          # print("警告: 起動時チェック - 初期日数が負のため無視します。") # ログ削除
                          days_to_check_val = None 
            except (ValueError, TypeError) as e:
                 # print(f"警告: 起動時チェック - 初期日数 '{initial_days_value}' の変換に失敗: {e}。全期間チェックします。") # ログ削除
                 days_to_check_val = None 

            try:
                unprocessed_count = has_unprocessed_mail(folder_path, account_email, days_to_check=days_to_check_val)
                
                if unprocessed_count > 0:
                    final_message = f"状態: {unprocessed_count}件の新規未処理メールがあります"
                else:
                    if output_path_exists:
                        final_message = "状態: 抽出結果ファイルあり。未処理メールはありません。"
                    else:
                        final_message = "状態: 対象のメールはありません" 
                q.put(final_message)
            except Exception as e:
                error_msg = f"状態: バックグラウンドチェックエラー - {e}"
                q.put(error_msg)
                # print(f"未処理チェックスレッドでエラーが発生: {e}") # ログ削除
                if not output_path_exists:
                    q.put("状態: 待機中（チェックエラー）。")
                    
        except Exception as outer_err:
             # print(f"❌ 未処理チェックスレッド全体で予期せぬエラー: {outer_err}\n{traceback.format_exc()}") # ログ削除
             q.put("状態: 未処理チェックで重大なエラー。")
             
        finally:
             # print(f"[CHECKER] Thread {thread_id} (Async Check) CoUninitialize() CALLED.") # ログ削除
             pythoncom.CoUninitialize()
             
    def check_queue():
        try:
            message = gui_queue.get(block=False)
            
            # --- ▼▼▼【ここから修正】▼▼▼ ---
            
            # 1. 完了メッセージを処理
            if message.startswith("EXTRACTION_COMPLETE:"):
                try:
                    parts = message.split(":", 2)
                    total_saved = parts[1]
                    final_message = parts[2]
                    
                    # アクティブなウィンドウを特定
                    search_window = main_elements.get("search_window")
                    active_parent = search_window if (search_window and search_window.winfo_exists()) else root
                    
                    messagebox.showinfo("完了", final_message, parent=active_parent)
                    status_label.config(text=f"状態: 処理完了。{total_saved} 件保存済み。")
                except Exception as e:
                    print(f"完了メッセージの表示エラー: {e}")

            # 2. 0件メッセージを処理
            elif message == "EXTRACTION_NO_ITEMS_FOUND":
                search_window = main_elements.get("search_window")
                active_parent = search_window if (search_window and search_window.winfo_exists()) else root
                messagebox.showinfo("完了", "処理対象のメールがありませんでした。", parent=active_parent)
                status_label.config(text="状態: 処理対象のメールがありませんでした。")

            # 3. エラーメッセージを処理
            elif message.startswith("EXTRACTION_ERROR:"):
                error_details = message.split(":", 1)[1]
                search_window = main_elements.get("search_window")
                active_parent = search_window if (search_window and search_window.winfo_exists()) else root
                messagebox.showerror("エラー", error_details, parent=active_parent)
                status_label.config(text="状態: エラー発生。")

            # 4. ボタン有効化メッセージを処理
            elif message == "EXTRACTION_COMPLETE_ENABLE_BUTTON":
                run_button = main_elements.get("run_button")
                if run_button:
                    try:
                        if run_button.winfo_exists():
                            run_button.config(state=tk.NORMAL)
                    except tk.TclError:
                        pass 
            
            # 5. 【★ここに追加★】 検索ボタン有効化メッセージを処理
            elif message == "ENABLE_SEARCH_BUTTON":
                search_button = main_elements.get("search_button")
                if search_button:
                    try:
                        if search_button.winfo_exists():
                            search_button.config(state=tk.NORMAL)
                            status_label.config(text="状態: DBにデータが保存されました。検索可能です。")
                    except tk.TclError:
                        pass
            
            # 6. それ以外のメッセージ (スキャン途中経過など)
            else:
                status_label.config(text=message)
            # --- ▲▲▲ 修正ここまで ▲▲▲ ---
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
             
        # --- ▼▼▼【修正】フリーズ対策: 軽量読み込み ▼▼▼
        # 必要な「軽い」列だけを読み込む
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
        # --- ▼▼▼【修正】SQLクエリに ORDER BY を追加 ▼▼▼ ---
        query = f"SELECT {light_columns_sql} FROM emails ORDER BY \"受信日時\" DESC"
        # print(f"DEBUG: Loading light data with query: {query}") # ログ
        df_for_gui = pd.read_sql_query(query, conn)
        # --- ▲▲▲ 修正ここまで ▲▲▲ ---
        conn.close()
        # --- ▼▼▼【削除】Pandas側でのソート処理を削除 ▼▼▼ ---
        # if not df_for_gui.empty and '受信日時' in df_for_gui.columns:
        #     df_for_gui['受信日時'] = pd.to_datetime(df_for_gui['受信日時'], errors='coerce')
        #     df_for_gui = df_for_gui.sort_values(by='受信日時', ascending=False, na_position='last').reset_index(drop=True)
        # --- ▲▲▲ 削除ここまで ▲▲▲ ---
        # --- ▼▼▼【ここに追加/修正】(6行) ▼▼▼ ---
        # 1. これから開くウィンドウに「旗」を渡す
        db_flag = main_elements.get("db_has_new_data_var")
        
        # 2. 「今から最新データで開く」ので、旗を False (更新なし) に倒す
        if db_flag:
            db_flag.set(False) 

        search_app = gui_search_window.App(
            root, 
            data_frame=df_for_gui,
            open_email_callback=open_outlook_email_by_id,
            db_has_new_data_var=db_flag # ← ★ 引数に db_flag を追加 ★
        ) 
        # --- ▲▲▲ 修正ここまで ▲▲▲ ---
        # --- ▼▼▼【ここに追加】▼▼▼ ---
        # どのウィンドウがアクティブか main_elements に記録
        main_elements["search_window"] = search_app 
        # --- ▲▲▲ 追加ここまで ▲▲▲ ---
        
        search_app.wait_window() 
        
    except Exception as e:
        messagebox.showerror("検索ウィンドウ起動エラー", f"検索一覧の表示中に予期せぬエラーが発生しました。\n詳細: {e}")
        traceback.print_exc()
    finally:
         # --- ▼▼▼【ここに追加】▼▼▼ ---
         # 検索ウィンドウが閉じたら記録を消す
         main_elements["search_window"] = None
         # --- ▲▲▲ 追加ここまで ▲▲▲ ---
         try:
             if root and root.winfo_exists():
                  root.deiconify() 
         except tk.TclError:
              pass 
         except Exception as e_final:
              print(f"警告: メインウィンドウ復元中に予期せぬエラー: {e_final}")


if __name__ == "__main__":
    main()