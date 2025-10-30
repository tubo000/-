# main_application.py (COM初期化 + チェッカー追加版)
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
# ----------------------------------------------------
# ユーティリティ関数群 (Outlook連携、DF処理)
# ----------------------------------------------------

def open_outlook_email_by_id(entry_id: str):
    """Entry IDを使用してOutlookデスクトップアプリでメールを開く関数。（GUI版）"""
    # --- ▼▼▼ チェッカー ▼▼▼ ---
    thread_id = threading.get_ident()
    print(f"\n[CHECKER] Thread {thread_id} (MainThread/OpenEmail) STARTING...")
    # --- ▲▲▲ チェッカー ▲▲▲ ---
    
    if not entry_id:
        messagebox.showerror("エラー", "Entry IDが指定されていません。")
        return

    try:
        pythoncom.CoInitialize() # ★ 維持
        print(f"[CHECKER] Thread {thread_id} (MainThread/OpenEmail) CoInitialize() CALLED.")
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
        print(f"[CHECKER] Thread {thread_id} (MainThread/OpenEmail) CoUninitialize() CALLED.")
        pythoncom.CoUninitialize() # ★ 維持


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
# 抽出処理ロジック (COM初期化は維持 + チェッカー)
# ----------------------------------------------------
def actual_run_extraction_logic(root, main_elements, target_email, folder_path, read_mode, read_days, status_label):
    
    # --- ▼▼▼ チェッカー ▼▼▼ ---
    thread_id = threading.get_ident()
    print(f"\n[CHECKER] Thread {thread_id} (Extraction) STARTING...")
    # --- ▲▲▲ チェッカー ▲▲▲ ---
    
    try:
        pythoncom.CoInitialize()
        print(f"[CHECKER] Thread {thread_id} (Extraction) CoInitialize() CALLED.")
    except Exception as e:
        print(f"[CHECKER] Thread {thread_id} (Extraction) CoInitialize() FAILED: {e}")
        pass 
        
    try:
        days_ago = None
        if read_days.strip():
            try:
                days_ago = int(read_days)
                # ▼▼▼【注意】このコードでは 0 はエラーになる ▼▼▼
                if days_ago < 1: raise ValueError("日数は1以上の整数を指定してください。")
            except ValueError:
                messagebox.showerror("入力エラー", "期間指定は1以上の整数で指定してください。")
                status_label.config(text="状態: 抽出失敗 (期間入力不正)。")
                return # finally が実行される

        if days_ago is not None:
            mode_text = f"未処理 (過去{days_ago}日)"
        else:
            mode_text = "未処理 (全期間)"
            
        status_label.config(text=f"状態: {target_email} アカウントからメール取得中 ({mode_text})...")
        root.update_idletasks() 

        # 内部関数 (CoInitialize なし) を呼び出す
        df_mail_data = get_mail_data_from_outlook_in_memory(
            folder_path, 
            target_email, 
            read_mode=read_mode, 
            days_ago=days_ago 
        )
        
        if df_mail_data.empty:
            status_label.config(text="状態: 処理対象のメールがありませんでした。")
            messagebox.showinfo("完了", "処理対象のメールがありませんでした。")
            return # finally が実行される

        status_label.config(text="状態: 抽出コアロジック実行中...")
        root.update_idletasks()
        df_extracted = extract_skills_data(df_mail_data)
        
        # --- データベース書き込み処理 ---
        df_output = df_extracted.copy()
        date_key_df = df_mail_data[['EntryID', '受信日時']].copy()
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
                 raise KeyError("抽出結果に EntryID が含まれていません。データベースに保存できません。")
            entry_ids_in_current_extraction = df_output['EntryID'].tolist()
            df_output.set_index('EntryID', inplace=True)
            try:
                existing_ids_set = set(pd.read_sql_query("SELECT EntryID FROM emails", conn)['EntryID'].tolist())
            except pd.io.sql.DatabaseError:
                existing_ids_set = set() 
            new_ids = [eid for eid in entry_ids_in_current_extraction if eid not in existing_ids_set]
            df_new = df_output.loc[new_ids]
            update_ids = [eid for eid in entry_ids_in_current_extraction if eid in existing_ids_set]
            df_update = df_output.loc[update_ids]
            
            # --- デバッグ表示 (DB書き込み) ---
            print("-" * 30)
            print(f"DEBUG(DB Write): データベースパス: {db_path}")
            print(f"DEBUG(DB Write): 今回抽出したメールの EntryID 件数: {len(entry_ids_in_current_extraction)}")
            print(f"DEBUG(DB Write): 既存DB内の EntryID 件数: {len(existing_ids_set)}")
            print(f"DEBUG(DB Write): 新規と判定された EntryID 件数: {len(new_ids)}")
            if not df_new.empty:
                 print(f"DEBUG(DB Write): これからDBに追記する {len(df_new)} 件の EntryID (先頭5件): {df_new.index.tolist()[:5]}")
            else:
                 print("DEBUG(DB Write): DBに追記する新規データはありません。")
            print("-" * 30)
            
            if not df_new.empty:
                df_new.to_sql('emails', conn, if_exists='append', index=True) 
                print(f"INFO: データベースに {len(df_new)} 件の新規レコードを追加しました。")
            if not df_update.empty:
                print(f"INFO: {len(df_update)} 件の既存レコードが見つかりましたが、更新はスキップされました。")
        except Exception as e:
            print(f"❌ データベース書き込み中にエラー発生: {e}")
            messagebox.showerror("DB書込エラー", f"データベースへの書き込み中にエラーが発生しました。\n詳細: {e}")
        finally:
            if conn: conn.close()
        # ----------------------------------------------------

        messagebox.showinfo("完了", f"抽出処理が正常に完了し、\n'{DATABASE_NAME}' に保存されました。\n検索一覧ボタンを押して結果を確認してください。")
        status_label.config(text=f"状態: 処理完了。DB保存済み。")
        search_button = main_elements.get("search_button")
        if search_button: search_button.config(state=tk.NORMAL)
        
    except Exception as e:
        status_label.config(text=f"状態: エラー発生 - {e}")
        messagebox.showerror("エラー", f"抽出処理中に予期せぬエラーが発生しました。\n詳細: {e}")
        traceback.print_exc()
        
    finally:
        # --- ボタン有効化のメッセージをキューに入れる ---
        q = main_elements.get("gui_queue")
        if q:
            q.put("EXTRACTION_COMPLETE_ENABLE_BUTTON") 
        
        # --- ▼▼▼ チェッカー ▼▼▼ ---
        print(f"[CHECKER] Thread {thread_id} (Extraction) CoUninitialize() CALLED.")
        # --- ▲▲▲ チェッカー ▲▲▲ ---
        pythoncom.CoUninitialize() # ★ スレッド終了時に実行

# ----------------------------------------------------
# 抽出ボタンコールバック (ボタン無効化 + チェッカー)
# ----------------------------------------------------
def run_extraction_callback():
    """抽出実行ボタンが押されたときの処理"""
    
    # --- ▼▼▼【ここからチェッカー】▼▼▼ ---
    thread_id = threading.get_ident()
    print("\n" + "="*40)
    print(f"[CHECKER] Thread {thread_id} (MainThread): 'run_extraction_callback' が呼び出されました。")
    
    run_button = main_elements.get("run_button")
    
    if run_button is None:
        print("DEBUG: ★★★ 原因判明 ★★★ -> main_elements.get(\"run_button\") の結果が None です。")
        print(f"DEBUG: 現在の main_elements のキー: {list(main_elements.keys())}")
        print("="*40 + "\n")
        return # ボタンがなければここで終了
    else:
        print(f"DEBUG: -> main_elements.get(\"run_button\") は Button オブジェクト '{run_button}' を取得しました。")

    current_state = None
    try:
        current_state = str(run_button.cget('state')) 
        print(f"DEBUG: ボタンの現在の状態 (state) は: '{current_state}' です。")
    except Exception as e:
        print(f"DEBUG: ボタンの状態取得中にエラー: {e}")
            
    print("DEBUG: これから if 文の判定に入ります...")
    print("="*40 + "\n")
    # --- ▲▲▲【チェッカーここまで】▲▲▲ ---
    
    # str() で比較
    if run_button and str(run_button.cget('state')) == tk.NORMAL:
        run_button.config(state=tk.DISABLED)
        print("INFO: 抽出実行ボタンを無効化。処理開始...")
        run_extraction_thread(root, main_elements, main_elements["extract_days_var"])
    else:
        print(f"INFO: 抽出処理が既に実行中か、ボタンが無効です。(if文がFalseと判定 / state='{current_state}')")

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
                 print("INFO: 入力エラーのため抽出実行ボタンを有効化しました。")
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
    # ... (この関数は変更なし、main_application.py 内に定義されている前提) ...
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
        # 修正: エラーではなく情報メッセージを返す
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
        print(f"❌ DBエラー発生: {e}")
        return f"エラー: DBファイルの処理中にエラーが発生しました ({e})"
    except Exception as e: 
         if conn: conn.rollback()
         print(f"❌ 予期せぬエラー発生: {e}")
         return f"エラー: 予期せぬエラーが発生しました ({e})"
    finally:
        if conn: conn.close()

# ----------------------------------------------------
# 削除スレッド本体 (COM初期化追加 + チェッカー)
# ----------------------------------------------------
def actual_run_file_deletion_logic(main_elements):
    
    # --- ▼▼▼【COM初期化 追加】▼▼▼ ---
    thread_id = threading.get_ident()
    print(f"\n[CHECKER] Thread {thread_id} (Deletion) STARTING...")
    try:
        pythoncom.CoInitialize()
        print(f"[CHECKER] Thread {thread_id} (Deletion) CoInitialize() CALLED.")
    except Exception as e:
        print(f"[CHECKER] Thread {thread_id} (Deletion) CoInitialize() FAILED: {e}")
        pass
    # --- ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲ ---
    
    try: # メインのロジックを try で囲み、finally で CoUninitialize する
        days_entry = main_elements["delete_days_entry"] 
        status_label = main_elements["status_label"]
        reset_category_var = main_elements["reset_category_var"]
        days_input = days_entry.get().strip()
        db_path = os.path.abspath(DATABASE_NAME) 

        try:
            days_ago = int(days_input)
            # ▼▼▼【注意】このコードでは 0 はエラーになる ▼▼▼
            if days_ago < 1: 
                raise ValueError("日数は1以上の整数を指定してください。")
        except ValueError as e:
            # ▼▼▼【注意】エラーメッセージも「1以上」のまま ▼▼▼
            messagebox.showerror("入力エラー", f"削除日数の入力が不正です: {e}\n(1以上の整数で指定)")
            status_label.config(text="状態: 削除失敗 (入力不正)。")
            return # finally が実行される

        reset_category_flag = reset_category_var.get()

        if days_ago == 0: # この条件は通らない
             confirm_prompt = f"🚨 **警告:** データベース内の**すべてのレコード**を削除します。\n"
        else:
             confirm_prompt = f"🚨 **警告:** データベース内の **{days_ago}日より古いレコード**を削除します。\n"
        if reset_category_flag:
            if days_ago == 0: pass # 通らない
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
                elif "INFO:" in delete_result_message: # INFOメッセージの場合
                     print(f"INFO: {delete_result_message}")
                     # db_had_error は False のまま
                else:
                     print(f"INFO: {delete_result_message}") 
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
            print(f"INFO: {delete_result_message}") 

        reset_count = 0
        category_reset_error = None
        if reset_category_flag:
            status_label.config(text=f"状態: Outlookカテゴリ解除中...")
            root.update_idletasks()
            try:
                # ▼▼▼【注意】days_ago=0 の場合はスキップする古いロジック ▼▼▼
                reset_days_param = days_ago if days_ago > 0 else None 
                if reset_days_param is not None:
                    reset_count = remove_processed_category(
                        main_elements["account_entry"].get().strip(),
                        main_elements["folder_entry"].get().strip(),
                        days_ago=reset_days_param
                    )
                    print(f"INFO: Outlookカテゴリリセット {reset_count} 件完了。") 
                else:
                     print("INFO: days_ago=0 のため、Outlookカテゴリの解除はスキップされました。")
            except NameError:
                 category_reset_error = "カテゴリ解除関数(remove_processed_category)が見つかりません。"
                 print(f"❌ {category_reset_error}")
                 status_label.config(text="状態: カテゴリ解除エラー (内部エラー)。")
            except Exception as e:
                 category_reset_error = f"Outlookカテゴリの解除中にエラーが発生しました。\n詳細: {e}"
                 print(f"❌ {category_reset_error}")
                 status_label.config(text="状態: カテゴリ解除エラー。")

        final_msg = delete_result_message 
        if reset_category_flag:
            # ▼▼▼【注意】スキップメッセージ分岐が残っている ▼▼▼
            if reset_days_param is not None:
                 final_msg += f"\nOutlookカテゴリリセット: {reset_count} 件完了"
            else:
                 final_msg += "\n(Outlookカテゴリの解除はスキップされました)"
                 
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
        
        # データベースファイルが見つからなかった INFO メッセージの場合
        elif "INFO:" in delete_result_message and not reset_category_flag:
             msg_title = "処理スキップ"
             msg_icon = 'info'
             final_status_text = "状態: DBファイルなし。"
             
        if msg_icon == 'info': messagebox.showinfo(msg_title, final_msg)
        elif msg_icon == 'warning': messagebox.showwarning(msg_title, final_msg)
        status_label.config(text=final_status_text) 
    
    except Exception as outer_err:
         print(f"❌ 削除スレッド全体で予期せぬエラー: {outer_err}\n{traceback.format_exc()}")
         try:
              status_label.config(text="状態: 削除スレッドで重大なエラー。")
         except: pass 
         
    finally:
        # --- ▼▼▼【COM終了 追加】▼▼▼ ---
        print(f"[CHECKER] Thread {thread_id} (Deletion) CoUninitialize() CALLED.")
        pythoncom.CoUninitialize() 
        # --- ▲▲▲▲▲▲▲▲▲▲▲▲▲ ---

# ----------------------------------------------------
# メイン実行関数 (GUI起動)
# ----------------------------------------------------
root = None
main_elements = {}

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
    # ▼▼▼【注意】GUIのラベルが「0=今日」になっていない ▼▼▼
    ttk.Label(days_frame, text="日 (空欄の場合は全期間)").pack(side=tk.LEFT)
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
    # ▼▼▼【注意】GUIのラベルが「0=全削除」になっていない ▼▼▼
    ttk.Label(delete_frame, text="日").grid(row=0, column=2, padx=(0, 10), pady=5, sticky='w') 
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
    
    # --- main_elements にグローバル変数として代入 ---
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
        "run_button": run_button, # 抽出ボタンも追加
        "gui_queue": gui_queue
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
        # --- ▼▼▼【COM初期化 追加】▼▼▼ ---
        thread_id = threading.get_ident()
        print(f"\n[CHECKER] Thread {thread_id} (Async Check) STARTING...")
        try:
            pythoncom.CoInitialize()
            print(f"[CHECKER] Thread {thread_id} (Async Check) CoInitialize() CALLED.")
        except Exception as e:
            print(f"[CHECKER] Thread {thread_id} (Async Check) CoInitialize() FAILED: {e}")
            pass
        # --- ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲ ---
        
        try: # メインのロジックを try で囲む
            output_path_exists = os.path.exists(output_file_abs_path)
            days_to_check_val = None
            try:
                if initial_days_value is not None and str(initial_days_value).strip():
                     days_to_check_val = int(initial_days_value) 
                     if days_to_check_val < 0:
                          print("警告: 起動時チェック - 初期日数が負のため無視します。")
                          days_to_check_val = None 
            except (ValueError, TypeError) as e:
                 print(f"警告: 起動時チェック - 初期日数 '{initial_days_value}' の変換に失敗: {e}。全期間チェックします。")
                 days_to_check_val = None 

            try:
                # 内部関数 (CoInitialize なし) を呼び出す
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
                print(f"未処理チェックスレッドでエラーが発生: {e}")
                if not output_path_exists:
                    q.put("状態: 待機中（チェックエラー）。")
                    
        except Exception as outer_err:
             print(f"❌ 未処理チェックスレッド全体で予期せぬエラー: {outer_err}\n{traceback.format_exc()}")
             q.put("状態: 未処理チェックで重大なエラー。")
             
        finally:
             # --- ▼▼▼【COM終了 追加】▼▼▼ ---
             print(f"[CHECKER] Thread {thread_id} (Async Check) CoUninitialize() CALLED.")
             pythoncom.CoUninitialize()
             # --- ▲▲▲▲▲▲▲▲▲▲▲▲▲ ---
             
    def check_queue():
        try:
            message = gui_queue.get(block=False)
            
            if message == "EXTRACTION_COMPLETE_ENABLE_BUTTON":
                run_button = main_elements.get("run_button")
                if run_button:
                    try:
                        if run_button.winfo_exists():
                            run_button.config(state=tk.NORMAL)
                            print("INFO: 抽出実行ボタンを有効化しました (via Queue)。")
                    except tk.TclError:
                        pass 
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
             
        df_for_gui = pd.read_sql_query("SELECT * FROM emails", conn)
        conn.close()

        search_app = gui_search_window.App(
            root, 
            data_frame=df_for_gui,
            open_email_callback=open_outlook_email_by_id 
        ) 
        search_app.wait_window() 
        
    except Exception as e:
        messagebox.showerror("検索ウィンドウ起動エラー", f"検索一覧の表示中に予期せぬエラーが発生しました。\n詳細: {e}")
        traceback.print_exc()
    finally:
         try:
             if root and root.winfo_exists():
                  root.deiconify() 
         except tk.TclError:
              pass 
         except Exception as e_final:
              print(f"警告: メインウィンドウ復元中に予期せぬエラー: {e_final}")


if __name__ == "__main__":
    main()