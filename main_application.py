# main_application.py (GUI簡略化 + 循環インポート修正版)
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
from datetime import timedelta # timedelta をインポート
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
    """出力データフレームの列順を調整し、'受信日時'と本文カラムを左側に固定する。（ローカル定義）"""
    fixed_leading_cols = [
        'メールURL', '受信日時', '件名', '名前', '信頼度スコア', 
        '本文(テキスト形式)', '本文(ファイル含む)', 'Attachments'
    ]
    # 📌 修正: fixed_leading_cols が df に存在する列のみを対象にする
    fixed_leading_cols = [col for col in fixed_leading_cols if col in df.columns]
    remaining_cols = [col for col in df.columns.tolist() if col not in fixed_leading_cols]
    return df.reindex(columns=fixed_leading_cols + remaining_cols, fill_value='N/A')

# ----------------------------------------------------
# 抽出処理ロジック (📌 修正)
# ----------------------------------------------------

def actual_run_extraction_logic(root, main_elements, target_email, folder_path, read_mode, read_days, status_label):
    
    try:
        pythoncom.CoInitialize()
    except Exception:
        pass 
        
    try:
        days_ago = None
        # 📌 修正: 日数入力欄 (read_days) が空でなければ整数に変換
        if read_days.strip():
            try:
                days_ago = int(read_days)
                if days_ago < 1: raise ValueError
            except ValueError:
                messagebox.showerror("入力エラー", "期間指定は1以上の整数で指定してください。")
                status_label.config(text="状態: 抽出失敗 (期間入力不正)。")
                return

        # 📌 修正: モードテキストを「未処理」固定に変更
        if days_ago is not None:
            mode_text = f"未処理 (過去{days_ago}日)"
        else:
            mode_text = "未処理 (全期間)"
            
        status_label.config(text=f"状態: {target_email} アカウントからメール取得中 ({mode_text})...")
        
        # 読み込みモードと日数を渡す (read_mode は "unprocessed" が渡される)
        df_mail_data = get_mail_data_from_outlook_in_memory(
            folder_path, 
            target_email, 
            read_mode=read_mode, 
            days_ago=days_ago 
        )
        
        if df_mail_data.empty:
            status_label.config(text="状態: 処理対象のメールがありませんでした。")
            messagebox.showinfo("完了", "処理対象のメールがありませんでした。")
            return

        status_label.config(text="状態: 抽出コアロジック実行中...")
        df_extracted = extract_skills_data(df_mail_data)
        
        # --- データベース書き込み処理 (変更なし) ---
        
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

            # 📌 修正: set_index する前に EntryID を保持しておく
            entry_ids_in_current_extraction = df_output['EntryID'].tolist()

            df_output.set_index('EntryID', inplace=True)
            
            try:
                # 📌 修正: 既存IDをセット(set)で取得して高速化
                existing_ids_set = set(pd.read_sql_query("SELECT EntryID FROM emails", conn)['EntryID'].tolist())
            except pd.io.sql.DatabaseError:
                existing_ids_set = set() 

            # 📌 修正: 新規IDをリスト内包表記とセットを使って効率的に特定
            new_ids = [eid for eid in entry_ids_in_current_extraction if eid not in existing_ids_set]
            
            # 📌 修正: df_output から新規IDの行だけを抽出
            df_new = df_output.loc[new_ids] # .loc を使ってインデックスで抽出
            
            # --- デバッグ表示 ---
            print("-" * 30)
            print(f"DEBUG: データベースパス: {db_path}")
            print(f"DEBUG: 今回抽出したメールの EntryID 件数: {len(entry_ids_in_current_extraction)}")
            print(f"DEBUG: 既存DB内の EntryID 件数: {len(existing_ids_set)}")
            print(f"DEBUG: 新規と判定された EntryID 件数: {len(new_ids)}")
            if not df_new.empty:
                 print(f"DEBUG: これからDBに追記する {len(df_new)} 件の EntryID (先頭5件): {df_new.index.tolist()[:5]}")
            else:
                 print("DEBUG: DBに追記する新規データはありません。")
            print("-" * 30)
            # --- デバッグ表示ここまで ---
                 
            # 更新データの抽出 (ロジックは変更なし、デバッグ用に追加)
            update_ids = [eid for eid in entry_ids_in_current_extraction if eid in existing_ids_set]
            df_update = df_output.loc[update_ids]
            
            if not df_new.empty:
                # 📌 index=True で EntryID をカラムとして追記
                df_new.to_sql('emails', conn, if_exists='append', index=True) 
                print(f"INFO: データベースに {len(df_new)} 件の新規レコードを追加しました。")

            if not df_update.empty:
                print(f"INFO: {len(df_update)} 件の既存レコードが見つかりましたが、更新はスキップされました。")

        except Exception as e:
            print(f"❌ データベース書き込み中にエラー発生: {e}")
            messagebox.showerror("DB書込エラー", f"データベースへの書き込み中にエラーが発生しました。\n詳細: {e}") # GUIにもエラー表示
        finally:
            if conn:
                conn.close()
        # ----------------------------------------------------

        messagebox.showinfo("完了", f"抽出処理が正常に完了し、\n'{DATABASE_NAME}' に保存されました。\n検索一覧ボタンを押して結果を確認してください。")
        status_label.config(text=f"状態: 処理完了。DB保存済み。")
        
        search_button = main_elements.get("search_button")
        if search_button:
            search_button.config(state=tk.NORMAL)
        
    except Exception as e:
        status_label.config(text=f"状態: エラー発生 - {e}")
        messagebox.showerror("エラー", f"抽出処理中に予期せぬエラーが発生しました。\n詳細: {e}")
        traceback.print_exc()
        
    finally:
        pythoncom.CoUninitialize()

# 📌 修正: run_extraction_thread の引数を変更
def run_extraction_thread(root, main_elements, extract_days_var):
    account_email = main_elements["account_entry"].get().strip()
    folder_path = main_elements["folder_entry"].get().strip()
    status_label = main_elements["status_label"]
    
    # 📌 修正: read_mode を "unprocessed" に固定
    read_mode = "unprocessed"
    # 📌 修正: 引数の StringVar から値を取得
    read_days = extract_days_var.get() 
    
    if not account_email or not folder_path:
        messagebox.showerror("入力エラー", "メールアカウントとフォルダパスの入力は必須です。")
        return

    # 📌 修正: スレッドに渡す引数を変更
    thread = threading.Thread(target=lambda: actual_run_extraction_logic(root, main_elements, account_email, folder_path, read_mode, read_days, status_label))
    thread.start()

# ----------------------------------------------------
# ファイル内のレコード削除ロジック (変更なし)
# ----------------------------------------------------
def run_deletion_thread(root, main_elements):
    """GUIをブロックしないよう、DBレコード削除を別スレッドで実行するラッパー。"""
    thread = threading.Thread(target=lambda: actual_run_file_deletion_logic(main_elements))
    thread.start()

# main_application.py (L280 付近の actual_run_file_deletion_logic 関数)

# (もし delete_processed_records を別ファイル utils.py に置いた場合)
# from utils import delete_processed_records

def actual_run_file_deletion_logic(main_elements):
    
    days_entry = main_elements["delete_days_entry"] 
    status_label = main_elements["status_label"]
    reset_category_var = main_elements["reset_category_var"]
    
    days_input = days_entry.get().strip()
    # データベースのフルパスを取得
    db_path = os.path.abspath(DATABASE_NAME) 
    
    try:
        # 日数を整数に変換 (0以上)
        days_ago = int(days_input)
        if days_ago < 0: 
            raise ValueError("日数は0以上の整数を指定してください。")
    except ValueError as e:
        messagebox.showerror("入力エラー", f"削除日数の入力が不正です: {e}")
        status_label.config(text="状態: 削除失敗 (入力不正)。")
        return

    # --- DBファイルの存在チェックは delete_processed_records 内で行う ---

    reset_category_flag = reset_category_var.get()

    # --- 確認メッセージの作成 ---
    if days_ago == 0:
         confirm_prompt = f"🚨 警告: データベース内のすべてのレコードを削除します。\n"
    else:
         confirm_prompt = f"🚨 警告: データベース内の {days_ago}日より古いレコードを削除します。\n"

    if reset_category_flag:
        confirm_prompt += f"また、Outlookメールの『{PROCESSED_CATEGORY_NAME}』マークも解除します。\n\n本当に実行しますか？"
    else:
        confirm_prompt += "\n本当に実行しますか？"

    # --- 削除実行の最終確認 ---
    confirm = messagebox.askyesno("確認", confirm_prompt)
    if not confirm:
        status_label.config(text="状態: 削除処理キャンセル。")
        return

    # --- 削除処理の開始 ---
    status_label.config(text=f"状態: DBレコード削除中...")
    # messageboxを表示する前に親ウィンドウを取得
    root = status_label.winfo_toplevel() 
    root.update_idletasks() # ステータスラベルの更新を即時反映

    # ----------------------------------------------------
    # ▼▼▼▼▼▼▼▼▼▼▼▼▼▼ 削除が必要なコード ▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    # deleted_count = 0 # 不要になる
    # reset_count = 0 # reset_count はカテゴリ解除部分で初期化
    # try:
    #     conn = sqlite3.connect(db_path)
    #     cursor = conn.cursor()
    #     cutoff_date_dt = datetime.datetime.now() - datetime.timedelta(days=days_ago)
    #     cutoff_date_str = cutoff_date_dt.strftime('%Y-%m-%d %H:%M:%S')
    #     cursor.execute(f"SELECT COUNT(*) FROM emails WHERE \"{DATE_COLUMN}\" < ?", (cutoff_date_str,))
    #     deleted_count = cursor.fetchone()[0]
    #     cursor.execute(f"DELETE FROM emails WHERE \"{DATE_COLUMN}\" < ?", (cutoff_date_str,))
    #     conn.commit()
    #     # カテゴリリセットの呼び出しはこの try ブロックの外に移動
    # except Exception as e:
    #     messagebox.showerror("削除エラー", f"データベース処理中にエラーが発生しました。\n詳細: {e}")
    #     status_label.config(text="状態: 削除エラー。")
    #     # 🔴 エラー時に finally が実行されない可能性がある
    # finally:
    #     if 'conn' in locals() and conn:
    #         conn.close()
    # ▲▲▲▲▲▲▲▲▲▲▲▲▲▲ 削除が必要なコード ▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    # ----------------------------------------------------
    
    # ----------------------------------------------------
    # ▼▼▼▼▼▼▼▼▼▼▼▼▼ 新しく追加するコード ▼▼▼▼▼▼▼▼▼▼▼▼▼
    # 新しい削除関数を呼び出す
    delete_result_message = delete_processed_records(days_ago, db_path)
    
    # 削除関数がエラーメッセージを返したかチェック
    if "エラー:" in delete_result_message:
        messagebox.showerror("削除エラー", delete_result_message)
        status_label.config(text="状態: 削除エラー。")
        return # DB削除が失敗したらここで終了
    # ▲▲▲▲▲▲▲▲▲▲▲▲▲ 新しく追加するコード ▲▲▲▲▲▲▲▲▲▲▲▲▲
    # ----------------------------------------------------

    # --- カテゴリマークのリセット (DB削除が成功した場合のみ実行) ---
    reset_count = 0
    if reset_category_flag:
        status_label.config(text=f"状態: Outlookカテゴリ解除中...")
        root.update_idletasks()
        try:
            # days_ago=0 の場合は None を渡して全期間のカテゴリ解除をするか、
            # days_ago > 0 の時だけ実行するかを決定する必要があります。
            # ここでは days_ago > 0 の時だけ解除するようにします（安全策）。
            reset_days = days_ago if days_ago > 0 else None 
            if reset_days is not None: # days_ago=0 (全削除) の場合はカテゴリ解除しない
                reset_count = remove_processed_category(
                    main_elements["account_entry"].get().strip(), 
                    main_elements["folder_entry"].get().strip(), 
                    days_ago=reset_days 
                )
            else:
                 print("INFO: days_ago=0 のため、Outlookカテゴリの解除はスキップされました。")

        except Exception as e:
             # DB削除は成功しているので、カテゴリ解除のエラーのみ報告
             messagebox.showerror("カテゴリ解除エラー", f"Outlookカテゴリの解除中にエラーが発生しました。\nDBレコードの削除は完了しています。\n詳細: {e}")
             status_label.config(text="状態: DB削除完了、カテゴリ解除エラー。")
             # ここで return せず、DB削除成功のメッセージは表示する
             # return # カテゴリ解除エラーでもDB削除は完了している

    # --- 最終結果メッセージ ---
    final_msg = delete_result_message # DB削除結果
    if reset_category_flag:
        # days_ago=0 でスキップした場合も考慮
        if reset_days is not None:
             final_msg += f"\nOutlookカテゴリリセット: {reset_count} 件完了"
        else:
             final_msg += "\n(Outlookカテゴリの解除はスキップされました)"
        
    messagebox.showinfo("処理完了", final_msg)
    status_label.config(text="状態: 削除処理完了。")
# ----------------------------------------------------------------------
# 💡 【最終版】 レコード削除関数 (SQLite専用)
# ----------------------------------------------------------------------
def delete_processed_records(days_ago: int, db_path: str) -> str:
    """
    指定された日数に基づき、SQLiteデータベース内の古いレコードを削除する。
    0: すべてのレコードを削除。
    1以上: N日前より古いレコードを削除（N日前の0時0分より前）。
    """
    try:
        days_ago = int(days_ago)
        if days_ago < 0:
             raise ValueError("日数は0以上の整数で指定してください。")
    except ValueError:
        return "エラー: 日数設定が不正です (0以上の整数で指定)。"

    # --- 削除対象の日付を計算 ---
    today = datetime.date.today()
    
    if days_ago == 0:
        # 0日の場合: すべて削除
        cutoff_datetime = datetime.datetime.combine(today + timedelta(days=1), datetime.time.min) # 比較用メッセージのため
        where_clause = "" # WHERE句なし
        target_message = "すべての取り込み記録"
    else:
        # N日前の場合
        cutoff_date = today - timedelta(days=days_ago)
        cutoff_datetime = datetime.datetime.combine(cutoff_date, datetime.time.min) # N日前の00:00:00
        cutoff_str = cutoff_datetime.strftime('%Y-%m-%d %H:%M:%S')
        # DBのカラム名 '受信日時' を使用
        where_clause = f"WHERE \"受信日時\" < '{cutoff_str}'"
        target_message = f"'{cutoff_date.strftime('%Y年%m月%d日')}' より古い取り込み記録"

    deleted_count = 0

    # --- SQLite DBからのレコード削除 ---
    if not os.path.exists(db_path):
        return f"エラー: データベースファイルが見つかりません ({os.path.basename(db_path)})"

    conn = None
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()

        # 1. 削除対象の件数を取得
        count_sql = f"SELECT COUNT(*) FROM emails {where_clause}"
        cursor.execute(count_sql)
        deleted_count = cursor.fetchone()[0]

        # 2. 削除を実行 (件数が0より大きい場合のみ)
        if deleted_count > 0:
            delete_sql = f"DELETE FROM emails {where_clause}"
            cursor.execute(delete_sql)
            conn.commit() # 変更を確定
            return f"{target_message} ({deleted_count}件) を削除しました。"
        else:
            return f"{target_message} は見つかりませんでした。削除は行われませんでした。"

    except sqlite3.Error as e: # DBエラー
        if conn:
            conn.rollback() # エラー時は元に戻す
        print(f"❌ DBエラー発生: {e}")
        return f"エラー: DBファイルの処理中にエラーが発生しました ({e})"
    except Exception as e: # その他のエラー
         if conn:
            conn.rollback()
         print(f"❌ 予期せぬエラー発生: {e}")
         return f"エラー: 予期せぬエラーが発生しました ({e})"
    finally:
        if conn:
            conn.close() # 接続を閉じる
# ----------------------------------------------------
# メイン実行関数 (GUI起動) (📌 GUI修正版)
# ----------------------------------------------------

def main():
    """
    アプリケーションのメインウィンドウを作成し、実行する。
    """
    root = tk.Tk()
    root.title("Outlook Mail Search Tool")
    
    # ----------------------------------------------------
    # 📌 修正: ウィンドウの高さを 600 に変更
    # ----------------------------------------------------
    window_width = 800
    window_height = 600 # UIが減ったため高さを縮小
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    center_x = int(screen_width/2 - window_width/2)
    center_y = int(screen_height/2 - window_height/2)
    root.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
    
    def on_main_window_close():
        root.destroy() 
    root.protocol("WM_DELETE_WINDOW", on_main_window_close)

    # --- 共有変数 ---
    # 📌 修正: read_mode_var を削除
    delete_days_var = tk.StringVar(value="14") 
    extract_days_var = tk.StringVar(value="14") # 抽出用（未処理N日）
    reset_category_var = tk.BooleanVar(value=False) 
    gui_queue = queue.Queue() # スレッド通信用
    
    saved_account, saved_folder = utils.load_config_csv() 
    if not saved_folder: saved_folder = TARGET_FOLDER_PATH 

    main_frame = Frame(root)
    main_frame.pack(padx=10, pady=10, fill='both', expand=True)
    
    # --- 設定ボタン (変更なし) ---
    top_button_frame = ttk.Frame(main_frame)
    top_button_frame.pack(fill='x', padx=10, pady=(10, 0))
    top_button_frame.grid_columnconfigure(0, weight=1) 
    top_button_frame.grid_columnconfigure(1, weight=0) 
    
    settings_button = ttk.Button(top_button_frame, text="⚙ 設定")
    settings_button.grid(row=0, column=1, padx=(0, 5), pady=5, sticky='e')

    # --- アカウント/フォルダ設定 (変更なし) ---
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
    
    # ----------------------------------------------------
    # 📌 修正: 処理/抽出関連のUI (ラジオボタンを削除)
    # ----------------------------------------------------
    process_frame = ttk.LabelFrame(main_frame, text="メールデータ抽出/検索")
    process_frame.pack(padx=10, pady=10, fill='x')
    process_frame.grid_columnconfigure(0, weight=1)
    process_frame.grid_columnconfigure(1, weight=1)
    
    # 📌 削除: mode_frame (ラジオボタン) を削除
    
    # 📌 修正: 期間指定を row=0 に移動
    days_frame = ttk.Frame(process_frame)
    days_frame.grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky='w')
    
    # 📌 修正: ラベルを「未処理メール」用に変更
    ttk.Label(days_frame, text="未処理メールの検索期間 (N日前まで):").pack(side=tk.LEFT)
    
    extract_days_entry = ttk.Entry(days_frame, textvariable=extract_days_var, width=10)
    extract_days_entry.pack(side=tk.LEFT, padx=5)
    ttk.Label(days_frame, text="日 (空欄の場合は全期間)").pack(side=tk.LEFT)
    
    # 📌 修正: ボタンを row=1 に移動
    run_button = ttk.Button(process_frame, text="抽出実行")
    run_button.grid(row=1, column=0, padx=5, pady=5, sticky='ew')
    
    search_button = ttk.Button(process_frame, text="検索一覧 (結果表示)", state=tk.DISABLED)
    search_button.grid(row=1, column=1, padx=5, pady=5, sticky='ew')
    
    # ----------------------------------------------------
    # 3. 削除機能のセクション (変更なし)
    # ----------------------------------------------------
    delete_frame = ttk.LabelFrame(main_frame, text="メール/レコード管理")
    delete_frame.pack(padx=10, pady=(10, 5), fill='x')
    
    delete_frame.grid_columnconfigure(0, weight=0)
    delete_frame.grid_columnconfigure(1, weight=0)
    delete_frame.grid_columnconfigure(2, weight=0)
    delete_frame.grid_columnconfigure(3, weight=1) 

    ttk.Label(delete_frame, text="N日前より古いレコード削除:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
    
    delete_days_entry = ttk.Entry(delete_frame, textvariable=delete_days_var, width=10)
    delete_days_entry.grid(row=0, column=1, padx=5, pady=5, sticky='w') 
    ttk.Label(delete_frame, text="日").grid(row=0, column=2, padx=(0, 10), pady=5, sticky='w') 
    
    delete_button = ttk.Button(delete_frame, text="レコード削除実行")
    delete_button.grid(row=1, column=0, columnspan=4, padx=5, pady=5, sticky='ew')
    
    reset_category_checkbox = ttk.Checkbutton(
        delete_frame, 
        text="処理済みマークを解除する", 
        variable=reset_category_var
    )
    reset_category_checkbox.grid(row=2, column=0, columnspan=4, padx=5, pady=(15, 5), sticky='w') 
    
    # --- ステータスラベル (変更なし) ---
    status_label = ttk.Label(main_frame, text="状態: 待機中", relief=tk.SUNKEN, anchor='w')
    status_label.pack(side=tk.BOTTOM, fill='x', padx=10, pady=(5, 0))
    
    # ----------------------------------------------------
    # main_elements の定義 (📌 修正)
    # ----------------------------------------------------
    main_elements = {
        "account_entry": account_entry,
        "folder_entry": folder_entry,
        "status_label": status_label,
        "search_button": search_button,
        "delete_days_entry": delete_days_entry, 
        "extract_days_entry": extract_days_entry, # 抽出用のEntry
        "settings_button": settings_button, 
        "reset_category_var": reset_category_var, 
        "extract_days_var": extract_days_var, # 抽出用のStringVar
    }
    
    # ----------------------------------------------------
    # コールバック関数の定義 (📌 修正)
    # ----------------------------------------------------
    def open_settings_callback():
        gui_elements.open_settings_window(
            root, main_elements["account_entry"], main_elements["status_label"]
        )
    
    def run_extraction_callback():
        # 📌 修正: extract_days_var (StringVar) を渡す
        run_extraction_thread(root, main_elements, main_elements["extract_days_var"])
        
    def open_search_callback():
        db_path = os.path.abspath(DATABASE_NAME)
        if not os.path.exists(db_path):
            messagebox.showwarning("警告", f"データベース ('{DATABASE_NAME}') が見つかりません。\n先に抽出を実行してください。")
            return
            
        try:
            root.withdraw() 
            
            conn = sqlite3.connect(db_path)
            df_for_gui = pd.read_sql_query("SELECT * FROM emails", conn)
            conn.close()

            # ----------------------------------------------------
            # 📌 修正 (循環インポートエラー対策):
            # App作成時に、引数として open_outlook_email_by_id 関数を渡す
            # ----------------------------------------------------
            search_app = gui_search_window.App(
                root, 
                data_frame=df_for_gui,
                open_email_callback=open_outlook_email_by_id # 
            ) 
            search_app.wait_window() # 検索ウィンドウが閉じられるのを待つ
            
        except Exception as e:
            messagebox.showerror("検索ウィンドウ起動エラー", f"検索一覧の表示中に予期せぬエラーが発生しました。\n詳細: {e}")
            traceback.print_exc()
        finally:
            try:
                if root.winfo_exists():
                    root.deiconify() # メインウィンドウを復元
            except tk.TclError:
                pass 
    
    # ----------------------------------------------------
    # ボタンにコマンドを設定
    # ----------------------------------------------------
    settings_button.config(command=open_settings_callback)
    run_button.config(command=run_extraction_callback)
    search_button.config(command=open_search_callback)
    delete_button.config(command=lambda: run_deletion_thread(root, main_elements))

    # ----------------------------------------------------
    # 起動時の処理 (変更なし)
    # ----------------------------------------------------
    output_file_abs_path = os.path.abspath(DATABASE_NAME) 
    
    if os.path.exists(output_file_abs_path):
        search_button.config(state=tk.NORMAL)
        status_label.config(text="状態: 抽出結果ファイルあり。検索一覧が利用可能です。")
    
    # 起動時の未処理メールチェック
    def check_unprocessed_async(account_email, folder_path, q):
        output_path_exists = os.path.exists(output_file_abs_path)
        
        try:
            unprocessed_count = has_unprocessed_mail(folder_path, account_email)
            
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
    
    def check_queue():
        try:
            message = gui_queue.get(block=False)
            status_label.config(text=message)
        except queue.Empty:
            pass
        finally:
            root.after(100, check_queue)

    threading.Thread(target=lambda: check_unprocessed_async(saved_account, saved_folder, gui_queue), daemon=True).start()
    
    root.after(100, check_queue)
    
    root.mainloop()

if __name__ == "__main__":
    main()