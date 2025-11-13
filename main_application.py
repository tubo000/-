
# main_application.py (★ マルチプロセス化 修正版 ★)
# (UIフリーズを根本的に解決する)
# 2025-11-13 更新

import os
import sys
import pandas as pd
import win32com.client as win32
# --- ▼▼▼ ★★★ 修正: threading -> multiprocessing ★★★ ▼▼▼
import threading # (open_outlook_email_by_id のスレッド起動のためだけに残す)
import multiprocessing # 👈 ★ 1. multiprocessing をインポート
import tkinter as tk
from tkinter import Frame, messagebox, simpledialog, ttk 
import pythoncom 
import re 
import traceback 
import os.path
import datetime 
import queue # 👈 ★ 2. queue は check_queue の .Empty 例外のためだけに残す
# --- ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲ ---
import sqlite3 
from datetime import timedelta 

# 外部モジュールのインポート
import gui_elements
import gui_search_window 
import utils 

# 既存の内部処理関数をインポート
from config import INPUT_QUESTION_CSV, MASTER_ANSWERS_PATH, OUTPUT_EVAL_PATH, NUM_RECORDS, TARGET_FOLDER_PATH, SCRIPT_DIR, MASTER_COLUMNS
# --- ▼▼▼ ★★★ 修正: worker_process から関数をインポート ★★★ ▼▼▼
# (このファイルから削除された関数を、新しいファイルからインポート)
try:
    import worker_process
except ImportError:
    messagebox.showerror("起動エラー", "worker_process.py が見つかりません。")
    sys.exit()
# --- ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲ ---

from email_processor import has_unprocessed_mail 
from config import DATABASE_NAME 

# --- グローバル変数の定義 ---
root = None
main_elements = {}
# ----------------------------------------------------

def open_outlook_email_by_id(entry_id: str):
    """ (v5のコードと変更なし) """
    if not entry_id:
        messagebox.showerror("エラー", "Entry IDが指定されていません。")
        return
    def _open_worker():
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
                q = main_elements.get("gui_queue")
                if q:
                    q.put(f"ERROR:指定された Entry ID のメールが見つかりませんでした。")
        except Exception as e:
            q = main_elements.get("gui_queue")
            if q:
                q.put(f"ERROR:Outlook連携エラー: {e}\nOutlookが起動しているか確認してください。")
        finally:
            pythoncom.CoUninitialize()
    threading.Thread(target=_open_worker, daemon=True).start()

def interactive_id_search_test():
    pass

# --- ▼▼▼ ★★★ 修正: 以下の4関数を「削除」 ★★★ ▼▼▼
# reorder_output_dataframe (worker_process.py に移動)
# actual_run_extraction_logic (worker_process.py に移動)
# delete_processed_records (worker_process.py に移動)
# actual_run_file_deletion_logic (worker_process.py に移動)
# --- ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲ ---

# ----------------------------------------------------
# 抽出ボタンコールバック (★ multiprocessing 版 ★)
# ----------------------------------------------------
def run_extraction_callback():
    run_button = main_elements.get("run_button")
    stop_button = main_elements.get("stop_button") 
    stop_flag = main_elements.get("stop_extraction_flag") # 👈 (multiprocessing.Event)
    
    if run_button is None:
        print("警告: run_button が main_elements に見つかりません。") 
        return 
    if str(run_button.cget('state')) == tk.NORMAL:
        run_button.config(state=tk.DISABLED)  
        stop_button.config(state=tk.NORMAL)   
        stop_flag.clear() # 👈 (Event を False に)
        
        run_extraction_process(root, main_elements, main_elements["extract_days_var"], main_elements["gui_queue"], stop_flag)
    else:
        pass 
        
def run_extraction_process(root, main_elements, extract_days_var, gui_queue, stop_flag):
    account_email = main_elements["account_entry"].get().strip()
    folder_path = main_elements["folder_entry"].get().strip()
    read_mode = "unprocessed"
    read_days = extract_days_var.get()
    
    if not account_email or not folder_path:
        gui_queue.put("ERROR:メールアカウントとフォルダパスの入力は必須です。")
        gui_queue.put("EXTRACTION_COMPLETE_ENABLE_BUTTON") 
        return
        
    # ★ 修正: threading.Thread -> multiprocessing.Process
    # (★ target は worker_process.py の関数を呼び出す)
    process = multiprocessing.Process(
        target=worker_process.actual_run_extraction_logic, 
        args=(
            { # 👈 ★ main_elements はプロセス間で安全に渡せないため、辞書として渡す
                "delete_days_entry": main_elements["delete_days_entry"].get() 
            }, 
            account_email, folder_path, 
            read_mode, read_days, 
            gui_queue, # 👈 (multiprocessing.Queue)
            stop_flag  # 👈 (multiprocessing.Event)
        )
    )
    process.start()

# ----------------------------------------------------
# 削除処理ロジック (★ multiprocessing 版 ★)
# ----------------------------------------------------

def run_deletion_thread(root, main_elements, gui_queue): # 👈 ★ 引数に gui_queue を追加
    # ★ 修正: threading.Thread -> multiprocessing.Process
    
    # ★ main_elements から必要な「値」だけを取り出す
    elements_for_worker = {
        "delete_days_entry": main_elements["delete_days_entry"].get()
    }
    
    process = multiprocessing.Process(
        target=worker_process.actual_run_file_deletion_logic, 
        args=(
            elements_for_worker, # 👈 ★ 辞書を渡す
            gui_queue # 👈 (multiprocessing.Queue)
        )
    )
    process.start()

# ----------------------------------------------------
# メイン実行関数 (GUI起動)
# ----------------------------------------------------
def main():
    global root, main_elements
    
    # --- ▼▼▼ ★★★ 修正: multiprocessing の初期化 ★★★ ▼▼▼
    # (EXE化する際は、main() の「直後」に置く)
    # multiprocessing.freeze_support() 
    # (↑ これは if __name__ == "__main__": の直下に置くのが正しい)
    # (main.py側で呼ばれるため、ここでは不要)
    # --- ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲ ---
    
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
        run_button = main_elements.get("run_button")
        stop_flag = main_elements.get("stop_extraction_flag")
        is_running = False 
        if run_button and str(run_button.cget('state')) == tk.DISABLED:
            is_running = True
        if is_running and stop_flag:
            print("INFO: ×ボタン検知。バックグラウンド処理に停止を要求します...")
            q = main_elements.get("gui_queue")
            if q: q.put("STATUS: 終了処理中...バックグラウンド処理の完了を待っています。")
            
            stop_flag.set() # 👈 (multiprocessing.Event を True に)
            main_elements["is_shutting_down"] = True
        else:
            print("INFO: 処理は実行されていません。すぐに終了します。")
            root.destroy() 
            
    root.protocol("WM_DELETE_WINDOW", on_main_window_close)

    delete_days_var = tk.StringVar(value="14") 
    extract_days_var = tk.StringVar(value="1") 
    db_has_new_data_var = tk.BooleanVar(value=False)
    
    # --- ▼▼▼ ★★★ 修正: threading -> multiprocessing ★★★ ▼▼▼
    gui_queue = multiprocessing.Queue()
    stop_extraction_flag = multiprocessing.Event() 
    # --- ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲ ---
    
    saved_account, saved_folder = utils.load_config_csv() 
    if not saved_folder: saved_folder = TARGET_FOLDER_PATH 

    main_frame = Frame(root)
    main_frame.pack(padx=10, pady=10, fill='both', expand=True)
    
    # ( ... (GUIのウィジェット配置 v34 と変更なし) ... )
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
    process_frame.grid_columnconfigure(2, weight=1) 
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
    stop_button = ttk.Button(process_frame, text="抽出中断", state=tk.DISABLED)
    stop_button.grid(row=1, column=2, padx=5, pady=5, sticky='ew')
    delete_frame = ttk.LabelFrame(main_frame, text="メール/レコード管理")
    delete_frame.pack(padx=10, pady=(10, 5), fill='x')
    delete_frame.grid_columnconfigure(0, weight=0)
    delete_frame.grid_columnconfigure(1, weight=0)
    delete_frame.grid_columnconfigure(2, weight=0)
    delete_frame.grid_columnconfigure(3, weight=1) 
    ttk.Label(delete_frame, text="N日前以前のレコードを削除:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
    delete_days_entry = ttk.Entry(delete_frame, textvariable=delete_days_var, width=10)
    delete_days_entry.grid(row=0, column=1, padx=5, pady=5, sticky='w') 
    ttk.Label(delete_frame, text="日 (0=全削除, 1=昨日も含む)").grid(row=0, column=2, padx=(0, 10), pady=5, sticky='w') 
    delete_button = ttk.Button(delete_frame, text="レコード削除実行")
    delete_button.grid(row=1, column=0, columnspan=4, padx=5, pady=5, sticky='ew') 
    status_label = ttk.Label(main_frame, text="状態: 待機中", relief=tk.SUNKEN, anchor='w')
    status_label.pack(side=tk.BOTTOM, fill='x', padx=10, pady=(5, 0))
    
    # --- グローバル辞書 (変更なし) ---
    main_elements["account_entry"] = account_entry
    main_elements["folder_entry"] = folder_entry
    main_elements["status_label"] = status_label
    main_elements["search_button"] = search_button
    main_elements["delete_days_entry"] = delete_days_entry
    main_elements["extract_days_entry"] = extract_days_entry
    main_elements["settings_button"] = settings_button
    main_elements["extract_days_var"] = extract_days_var
    main_elements["run_button"] = run_button
    main_elements["gui_queue"] = gui_queue
    main_elements["db_has_new_data_var"] = db_has_new_data_var
    main_elements["stop_extraction_flag"] = stop_extraction_flag 
    main_elements["stop_button"] = stop_button               
    main_elements["delete_button"] = delete_button 
    main_elements["is_shutting_down"] = False 
    
    # --- ボタンの動作を割り当て (変更なし) ---
    settings_button.config(command=open_settings_callback)
    run_button.config(command=run_extraction_callback)
    search_button.config(command=open_search_callback)
    
    def deletion_callback_with_confirm():
        """ (v5のコードと変更なし) """
        days_input = delete_days_var.get().strip()
        try:
            days_ago = int(days_input)
            if days_ago < 0: raise ValueError()
        except ValueError:
            messagebox.showerror("入力エラー", "削除日数は 0以上 の整数で指定してください。")
            return
        if days_ago == 0:
             confirm_prompt = f"🚨 警告: データベース内のすべてのレコードを削除します。\n"
        else:
             confirm_prompt = f"🚨 警告: データベース内の {days_ago}日(を含む)より古いレコードを削除します。\n"
        confirm_prompt += "\n本当に実行しますか？"
        
        confirm = messagebox.askyesno("最終確認", confirm_prompt, icon='warning')
        if not confirm:
            status_label.config(text="状態: 削除処理キャンセル。")
            return
        
        run_button.config(state=tk.DISABLED)
        delete_button.config(state=tk.DISABLED)
        # ★ 修正: run_deletion_thread に gui_queue を渡す
        run_deletion_thread(root, main_elements, gui_queue)

    delete_button.config(command=deletion_callback_with_confirm)

    output_file_abs_path = os.path.abspath(DATABASE_NAME) 
    stop_button.config(command=lambda: stop_extraction_flag.set())
    if os.path.exists(output_file_abs_path):
        search_button.config(state=tk.NORMAL)
        status_label.config(text="状態: 抽出結果ファイルあり。検索一覧が利用可能です。")

    # --- 起動時の未処理メールチェック ---
    def check_unprocessed_async(account_email, folder_path, q, initial_days_value):
        """ (v5のコードと変更なし) """
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
                unprocessed_count = has_unprocessed_mail(folder_path, account_email, days_to_check=days_to_check_val)
                if unprocessed_count > 0:
                    final_message = f"状態: {unprocessed_count}件の未処理メールがあります (直近300件を確認)。"
                else:
                    if output_path_exists:
                        final_message = "状態: 抽出結果ファイルあり。未処理メールはありません。"
                    else:
                        final_message = "状態: 対象のメールはありません" 
                q.put(f"STATUS:{final_message}") 
            except Exception as e:
                error_msg = f"状態: バックグラウンドチェックエラー - {e}"
                q.put(f"STATUS:{error_msg}") 
                if not output_path_exists:
                    q.put("STATUS:状態: 待機中（チェックエラー）。")
        except Exception as outer_err:
             q.put("STATUS:状態: 未処理チェックで重大なエラー。")
        finally:
             pythoncom.CoUninitialize()
             
# main_application.py の check_queue 関数

    def check_queue():
        """ メインスレッドでキューを監視し、GUIを安全に更新する """
        try:
            message = gui_queue.get(block=False)
            
            if message.startswith("STATUS:"):
                status_label.config(text=message[len("STATUS:"):])
            elif message.startswith("ERROR:"):
                messagebox.showerror("エラー", message[len("ERROR:"):])
            elif message.startswith("DB_ERROR:"):
                messagebox.showerror("DB書込エラー", message[len("DB_ERROR:"):])

            # ▼▼▼ ★★★ ここから修正 ★★★ ▼▼▼
            # 'MSGBOX:' ではなく、'MSGBOX_...' の各タイプを処理する

            elif message.startswith("MSGBOX_INFO:"):
                try:
                    # "MSGBOX_INFO:" の次から分割 (e.g., "タイトル:本文")
                    parts = message[len("MSGBOX_INFO:"):].split(":", 1)
                    title = parts[0]
                    msg_body = parts[1]
                    messagebox.showinfo(title, msg_body)
                except Exception:
                    # 予期せぬ形式なら、プレフィックスだけ取って表示
                    messagebox.showinfo("情報", message[len("MSGBOX_INFO:"):])

            elif message.startswith("MSGBOX_WARNING:"):
                try:
                    parts = message[len("MSGBOX_WARNING:"):].split(":", 1)
                    title = parts[0]
                    msg_body = parts[1]
                    messagebox.showwarning(title, msg_body)
                except Exception:
                    messagebox.showwarning("警告", message[len("MSGBOX_WARNING:"):])

            elif message.startswith("MSGBOX_ERROR:"):
                try:
                    parts = message[len("MSGBOX_ERROR:"):].split(":", 1)
                    title = parts[0]
                    msg_body = parts[1]
                    messagebox.showerror(title, msg_body)
                except Exception:
                    messagebox.showerror("エラー", message[len("MSGBOX_ERROR:"):])

            # (古い MSGBOX: の分岐は削除、または上記に吸収される)
            # ▲▲▲ ★★★ 修正ここまで ★★★ ▲▲▲

            elif message == "EXTRACTION_COMPLETE_ENABLE_BUTTON":
                run_button.config(state=tk.NORMAL)
                stop_button.config(state=tk.DISABLED)
                delete_button.config(state=tk.NORMAL) 
                if main_elements.get("is_shutting_down") == True:
                    print("INFO: バックグラウンド処理が安全に停止しました。ウィンドウを閉じます。")
                    root.destroy()
            elif message == "DELETION_COMPLETE_ENABLE_BUTTON":
                run_button.config(state=tk.NORMAL)
                stop_button.config(state=tk.DISABLED)
                delete_button.config(state=tk.NORMAL)
            elif message == "ENABLE_SEARCH_BUTTON":
                search_button.config(state=tk.NORMAL)
                db_has_new_data_var.set(True) 
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
                
        except queue.Empty:
            pass
        except Exception as e:
            print(f"CRITICAL: check_queue でエラー: {e}")
        finally:
            try:
                if root and root.winfo_exists(): root.after(100, check_queue)
            except tk.TclError: pass
    initial_extract_days = None
    if "extract_days_var" in main_elements:
         try: initial_extract_days = main_elements["extract_days_var"].get()
         except tk.TclError: pass 
              
    threading.Thread(target=lambda: check_unprocessed_async(saved_account, saved_folder, gui_queue, initial_extract_days), daemon=True).start()
    
    root.after(100, check_queue) # 👈 ★ キューの監視を開始
    root.mainloop()

# ----------------------------------------------------
# 外部コールバック
# ----------------------------------------------------
def open_settings_callback():
    if root and main_elements:
        gui_elements.open_settings_window(
            root, main_elements["account_entry"], main_elements["status_label"]
        )

def open_search_callback():
    """ (v34のコードと変更なし) """
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
        
        db_flag = main_elements.get("db_has_new_data_var")
        
        search_app = gui_search_window.App(
            root, 
            main_elements, 
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
             if root and root.winfo_exists() and not main_elements.get("is_shutting_down", False):
                  root.deiconify()
         except tk.TclError:
              pass 
         except Exception as e_final:
              print(f"警告: メインウィンドウ復元中に予期せぬエラー: {e_final}")


# (★ main.py から __name__ == "__main__" ブロックを移動)
if __name__ == "__main__":
    # --- ▼▼▼ ★★★ 修正: multiprocessing の初期化 ★★★ ▼▼▼
    multiprocessing.freeze_support() 
    # --- ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲ ---
    main()
