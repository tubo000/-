# main_application.py (GUI統合とメイン実行フロー - 最終統合版)

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

# 外部モジュールのインポート
import gui_elements
import gui_search_window 
import utils 

# 既存の内部処理関数をインポート
# config.py からの定数インポートは維持
from config import INPUT_QUESTION_CSV, MASTER_ANSWERS_PATH, OUTPUT_EVAL_PATH, NUM_RECORDS, TARGET_FOLDER_PATH, SCRIPT_DIR
from extraction_core import extract_skills_data
from evaluator_core import run_triple_csv_validation, get_question_data_from_csv
# email_processor.py からのインポート (件名・本文・日時が揃ったデータを取得)
from email_processor import get_mail_data_from_outlook_in_memory, OUTPUT_FILENAME 


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
    """GUIのメニューなどから呼び出されるEntry ID検索機能。"""
    test_entry_id = simpledialog.askstring("Entry ID テスト", 
                                          "テスト用の Entry ID をペーストしてください:", 
                                          initialvalue="")
    if test_entry_id:
        open_outlook_email_by_id(test_entry_id.strip())
    else:
        messagebox.showinfo("テストスキップ", "Entry ID が入力されなかったため、テストをスキップします。")


def reorder_output_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """出力データフレームの列順を調整し、'受信日時'と本文カラムを左側に固定する。"""
    fixed_leading_cols = [
        'メールURL', '受信日時', '件名', '名前', '信頼度スコア', 
        '本文(テキスト形式)', '本文(ファイル含む)', 'Attachments'
    ]
    fixed_leading_cols = [col for col in fixed_leading_cols if col in df.columns]
    remaining_cols = [col for col in df.columns.tolist() if col not in fixed_leading_cols]
    return df.reindex(columns=fixed_leading_cols + remaining_cols, fill_value='N/A')

# ----------------------------------------------------
# 抽出処理ロジック
# ----------------------------------------------------

def actual_run_extraction_logic(root, main_elements, target_email, folder_path, read_mode, read_days, status_label):
    
    try:
        pythoncom.CoInitialize()
    except Exception:
        pass 
        
    try:
        # 期間指定モードの入力値チェック (維持)
        days_ago = None
        if read_mode == "days":
            try:
                days_ago = int(read_days)
                if days_ago < 1: raise ValueError
            except ValueError:
                messagebox.showerror("入力エラー", "期間指定モードでは、日数を1以上の整数で指定してください。")
                status_label.config(text="状態: 抽出失敗 (期間入力不正)。")
                return

        mode_text = {"all": "全て", "unprocessed": "未処理のみ", "days": f"過去{days_ago}日"}.get(read_mode, "全て")
        status_label.config(text=f"状態: {target_email} アカウントからメール取得中 ({mode_text})...")
        
        # 読み込みモードと日数を渡す
        # df_mail_data には 'EntryID', '件名', '受信日時', '本文(テキスト形式)', '本文(ファイル含む)', 'Attachments' が含まれるようになった
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
        # df_extracted には抽出結果(信頼度スコアやスキル名)が追加される
        df_extracted = extract_skills_data(df_mail_data)
        
        # ----------------------------------------------------
        # 📌 修正ロジックの再確認: df_mail_dataには本文・件名・日時が全て含まれる
        # df_extractedは df_mail_data を基に作られているため、本文/件名は継承されているはず。
        # ここでは、extract_skills_dataによって失われた可能性のある '受信日時' を念のためマージし直す。
        # ----------------------------------------------------
        
        df_output = df_extracted.copy()
        
        # 必要なカラムを元のデータから取得（日時とキー）。件名・本文は df_extracted に残っている前提。
        date_key_df = df_mail_data[['EntryID', '受信日時']].copy()
        
        # 抽出結果と日時情報を EntryID でマージ
        # マージ前に df_output から '受信日時' を削除し、df_mail_data の日時情報で上書きする
        if '受信日時' in df_output.columns:
            df_output.drop(columns=['受信日時'], inplace=True, errors='ignore')
            
        df_output = pd.merge(
            df_output,
            date_key_df,
            on='EntryID',
            how='left' 
        )

        # EntryIDをURLに変換
        if 'EntryID' in df_output.columns and 'メールURL' not in df_output.columns:
             df_output.insert(0, 'メールURL', df_output.apply(lambda row: f"outlook:{row['EntryID']}", axis=1))

        # 列順の整理
        df_output = reorder_output_dataframe(df_output)
        final_drop_list = ['EntryID', '宛先メール', '本文(抽出元結合)'] 
        final_drop_list = [col for col in df_output.columns if col in final_drop_list]
        df_output = df_output.drop(columns=final_drop_list, errors='ignore')
        
        # 受信日時カラムを保護しつつ、他の文字列をエスケープ
        DATE_COLUMN = '受信日時'
        
        for col in df_output.columns:
            # 日付カラムではない、かつオブジェクト型（文字列）のカラムのみエスケープ
            if col != DATE_COLUMN and df_output[col].dtype == object:
                df_output[col] = df_output[col].str.replace(r'^=', r"'=", regex=True)
                
        # Excelファイルへの書き出し
        output_file_abs_path = os.path.abspath(OUTPUT_FILENAME)
        df_output.to_excel(output_file_abs_path, index=False) 

        messagebox.showinfo("完了", f"抽出処理が正常に完了し、\n'{OUTPUT_FILENAME}' に出力されました。\n検索一覧ボタンを押して結果を確認してください。")
        status_label.config(text=f"状態: 処理完了。ファイル出力済み。")
        
        # 検索ボタンを有効化
        search_button = main_elements.get("search_button")
        if search_button:
            search_button.config(state=tk.NORMAL)
        
    except Exception as e:
        status_label.config(text=f"状態: エラー発生 - {e}")
        messagebox.showerror("エラー", f"抽出処理中に予期せぬエラーが発生しました。\n詳細: {e}")
        traceback.print_exc()
        
    finally:
        pythoncom.CoInitialize()

def run_extraction_thread(root, main_elements, read_mode_var, extract_days_entry):
    """GUIをブロックしないよう、抽出処理を別スレッドで実行するラッパー。"""
    account_email = main_elements["account_entry"].get().strip()
    folder_path = main_elements["folder_entry"].get().strip()
    status_label = main_elements["status_label"]
    
    read_mode = read_mode_var.get()
    read_days = extract_days_entry.get()
    
    if not account_email or not folder_path:
        messagebox.showerror("入力エラー", "メールアカウントとフォルダパスの入力は必須です。")
        return

    thread = threading.Thread(target=lambda: actual_run_extraction_logic(root, main_elements, account_email, folder_path, read_mode, read_days, status_label))
    thread.start()

# ----------------------------------------------------
# ファイル内のレコード削除ロジック
# ----------------------------------------------------

def run_deletion_thread(root, main_elements):
    """GUIをブロックしないよう、ファイルレコード削除を別スレッドで実行するラッパー。"""
    days_entry = main_elements["delete_days_entry"] 
    status_label = main_elements["status_label"]

    thread = threading.Thread(target=lambda: actual_run_file_deletion_logic(days_entry, status_label))
    thread.start()

def actual_run_file_deletion_logic(days_entry, status_label):
    
    days_input = days_entry.get().strip()
    output_file_path = os.path.abspath(OUTPUT_FILENAME)
    DATE_COLUMN = '受信日時' # 削除基準となるカラム名
    
    try:
        days_ago = int(days_input)
        if days_ago < 1:
            raise ValueError("日数は1以上の整数を指定してください。")
    except ValueError as e:
        messagebox.showerror("入力エラー", f"削除日数の入力が不正です: {e}")
        status_label.config(text="状態: 削除失敗 (入力不正)。")
        return

    if not os.path.exists(output_file_path):
        messagebox.showwarning("警告", f"ファイルが見つかりません。削除処理をスキップします: {OUTPUT_FILENAME}")
        status_label.config(text="状態: ファイルなし。")
        return

    confirm = messagebox.askyesno(
        "確認", 
        f"🚨 警告: ファイル '{OUTPUT_FILENAME}' 内の '{DATE_COLUMN}' が {days_ago}日より古いレコードを削除し、上書き保存します。\n\n本当に実行しますか？"
    )
    if not confirm:
        status_label.config(text="状態: 削除処理キャンセル。")
        return

    status_label.config(text=f"状態: {days_ago}日より古いレコードを削除中...")
    
    try:
        # 1. ファイルを読み込み (Excel出力のため read_excel を使用)
        df = pd.read_excel(output_file_path)
        
        if DATE_COLUMN not in df.columns:
            raise KeyError(f"削除基準となる '{DATE_COLUMN}' カラムがファイルに見つかりません。抽出実行後、ファイルに日付カラムがあるか確認してください。")

        # 2. 削除基準を計算
        cutoff_date = datetime.datetime.now() - datetime.timedelta(days=days_ago)
        
        # 3. フィルタリングと削除
        initial_count = len(df)
        
        # '受信日時' カラムを datetime 型に変換
        df['受信日時_dt'] = pd.to_datetime(df[DATE_COLUMN], errors='coerce') 
        
        # 日付変換に成功し、かつカットオフ日より新しいレコードを保持
        df_kept = df[df['受信日時_dt'].notna() & (df['受信日時_dt'] >= cutoff_date)].copy()
        
        deleted_count = initial_count - len(df_kept)
        
        # 4. ファイルを上書き保存
        df_kept.drop(columns=['受信日時_dt'], errors='ignore', inplace=True) # テンポラリカラムを削除
        df_kept.to_excel(output_file_path, index=False)
        
        messagebox.showinfo("削除完了", f"ファイルから {days_ago}日より古いレコード {deleted_count} 件を削除しました。\n残レコード数: {len(df_kept)} 件")
        status_label.config(text="状態: 削除処理完了。")
        
    except Exception as e:
        messagebox.showerror("削除エラー", f"ファイルレコード削除中にエラーが発生しました。\n詳細: {e}")
        status_label.config(text="状態: 削除エラー。")

# ----------------------------------------------------
# メイン実行関数 (GUI起動)
# ----------------------------------------------------

def main():
    """
    アプリケーションのメインウィンドウを作成し、実行する。
    """
    root = tk.Tk()
    root.title("Outlook Mail Search Tool")
    root.geometry("800x650") 

    # --- 共有変数 ---
    read_mode_var = tk.StringVar(value="all") 
    delete_days_var = tk.StringVar(value="14") 
    extract_days_var = tk.StringVar(value="14") 
    
    # 2. 初期設定データの読み込み
    saved_account, saved_folder = utils.load_config_csv() 
    if not saved_folder: saved_folder = TARGET_FOLDER_PATH 

    # 3. メインフレームと設定フレームの作成
    main_frame = Frame(root)
    main_frame.pack(padx=10, pady=10, fill='both', expand=True)
    
    # 設定ボタン用のフレームを画面のトップに作成
    top_button_frame = ttk.Frame(main_frame)
    top_button_frame.pack(fill='x', padx=10, pady=(10, 0))
    top_button_frame.grid_columnconfigure(0, weight=1) 
    top_button_frame.grid_columnconfigure(1, weight=0) 
    
    # 4. コールバック関数の定義
    
    main_elements = {} 
    
    def open_settings_callback():
        # gui_elements.py の open_settings_window を呼び出す
        gui_elements.open_settings_window(
            root, main_elements["account_entry"], main_elements["status_label"]
        )
    
    # 設定ボタンの作成と配置
    settings_button = ttk.Button(
        top_button_frame, 
        text="⚙ 設定",
        command=open_settings_callback
    )
    settings_button.grid(row=0, column=1, padx=(0, 5), pady=5, sticky='e')

    # 1. アカウント/フォルダ設定
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
    
    # 2. 処理/抽出関連
    process_frame = ttk.LabelFrame(main_frame, text="メールデータ抽出/検索")
    process_frame.pack(padx=10, pady=10, fill='x')
    
    process_frame.grid_columnconfigure(0, weight=1)
    process_frame.grid_columnconfigure(1, weight=1)
    
    # 読み込みモードのラジオボタンフレーム
    mode_frame = ttk.LabelFrame(process_frame, text="読み込みモード")
    mode_frame.grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky='ew')
    
    ttk.Radiobutton(mode_frame, text="全て読み込む (試験用)", variable=read_mode_var, value="all").pack(side=tk.LEFT, padx=10, pady=5)
    ttk.Radiobutton(mode_frame, text="未処理のみ", variable=read_mode_var, value="unprocessed").pack(side=tk.LEFT, padx=10, pady=5)
    # 期間指定モードのラジオボタン
    ttk.Radiobutton(mode_frame, text="期間指定", variable=read_mode_var, value="days").pack(side=tk.LEFT, padx=10, pady=5)

    # 期間日数入力フィールド
    days_frame = ttk.Frame(process_frame)
    days_frame.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky='w')
    ttk.Label(days_frame, text="期間 (N日前まで):").pack(side=tk.LEFT)
    extract_days_entry = ttk.Entry(days_frame, textvariable=extract_days_var, width=10)
    extract_days_entry.pack(side=tk.LEFT, padx=5)
    ttk.Label(days_frame, text="日").pack(side=tk.LEFT)
    
    def run_extraction_callback():
        run_extraction_thread(root, main_elements, read_mode_var, extract_days_entry)
        
    # 抽出実行ボタン
    run_button = ttk.Button(
        process_frame, 
        text="抽出実行", 
        command=run_extraction_callback
    )
    run_button.grid(row=2, column=0, padx=5, pady=5, sticky='ew')
    
    # 検索一覧ボタン (前回同様に無効化から開始)
    def open_search_callback():
        output_file_abs_path = os.path.abspath(OUTPUT_FILENAME)
        if not os.path.exists(output_file_abs_path):
            messagebox.showwarning("警告", f"抽出結果ファイル ('{OUTPUT_FILENAME}') が見つかりません。\n先に抽出を実行してください。")
            return
        try:
            root.withdraw() 
            gui_search_window.main()
        except Exception as e:
            messagebox.showerror("検索ウィンドウ起動エラー", f"検索一覧の表示中に予期せぬエラーが発生しました。\n詳細: {e}")
            traceback.print_exc()
        finally:
            root.deiconify() 
    
    search_button = ttk.Button(
        process_frame, 
        text="検索一覧 (結果表示)", 
        command=open_search_callback, 
        state=tk.DISABLED # 初期状態は無効
    )
    search_button.grid(row=2, column=1, padx=5, pady=5, sticky='ew')
    
    # 3. 削除機能のセクション
    delete_frame = ttk.LabelFrame(main_frame, text="レコード削除（ファイル）")
    delete_frame.pack(padx=10, pady=(10, 5), fill='x')
    
    delete_frame.grid_columnconfigure(1, weight=1)
    
    ttk.Label(delete_frame, text="N日前より古いレコードを削除:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
    
    delete_days_entry = ttk.Entry(delete_frame, textvariable=delete_days_var, width=10)
    delete_days_entry.grid(row=0, column=1, padx=5, pady=5, sticky='w')
    ttk.Label(delete_frame, text="日").grid(row=0, column=2, padx=(0, 10), pady=5, sticky='w')
    
    # 削除実行ボタン
    ttk.Button(
        delete_frame, 
        text="削除実行", 
        command=lambda: run_deletion_thread(root, main_elements) 
    ).grid(row=1, column=0, columnspan=3, padx=5, pady=5, sticky='ew')
    
    # 4. ステータスラベル
    status_label = ttk.Label(main_frame, text="状態: 待機中", relief=tk.SUNKEN, anchor='w')
    status_label.pack(side=tk.BOTTOM, fill='x', padx=10, pady=(5, 0))
    
    # 5. 全要素を格納する辞書
    main_elements = {
        "account_entry": account_entry,
        "folder_entry": folder_entry,
        "status_label": status_label,
        "search_button": search_button,
        "delete_days_entry": delete_days_entry, 
        "extract_days_entry": extract_days_entry, 
        "settings_button": settings_button, 
    }
    
    # 6. アプリケーションの開始
    root.mainloop()

if __name__ == "__main__":
    main()