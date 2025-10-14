# main_application.py (GUI統合とメイン実行フロー)

import os
import sys
import pandas as pd
import win32com.client as win32
import threading 
import tkinter as tk
from tkinter import Frame, messagebox, simpledialog 
import pythoncom 

# 外部モジュールのインポート
import gui_elements
import gui_search_window
import utils 

# 既存の内部処理関数をインポート
from config import INPUT_QUESTION_CSV, MASTER_ANSWERS_PATH, OUTPUT_EVAL_PATH, NUM_RECORDS, TARGET_FOLDER_PATH, SCRIPT_DIR
from data_generation import generate_raw_data, export_dataframes_to_tsv
from extraction_core import extract_skills_data
from evaluator_core import run_triple_csv_validation, get_question_data_from_csv
from email_processor import run_email_extraction, get_mail_data_from_outlook_in_memory, OUTPUT_FILENAME


# ----------------------------------------------------
# ユーティリティ関数群 (GUI/コンソール連携用)
# ----------------------------------------------------

def open_outlook_email_by_id(entry_id: str):
    """Entry IDを使用してOutlookデスクトップアプリでメールを開く関数。"""
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
            messagebox.showinfo("成功", f"メールを正常に開きました: {getattr(olItem, 'Subject', '件名なし')}")
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


def actual_run_extraction_logic(root, main_elements, target_email, folder_path, status_label):
    """
    GUIから呼び出される抽出のコアロジック (時間のかかる処理)
    ファイル出力後、検索UIを起動するロジックを削除する。
    """
    # COM初期化
    try:
        pythoncom.CoInitialize()
    except Exception:
        pass 
        
    try:
        status_label.config(text=f"状態: {target_email} アカウントからメール取得中...", fg="blue")
        
        df_mail_data = get_mail_data_from_outlook_in_memory(folder_path, target_email)
        
        if df_mail_data.empty:
            status_label.config(text="状態: 処理対象のメールがありませんでした。", fg="green")
            return

        status_label.config(text="状態: 抽出コアロジック実行中...", fg="blue")
        df_extracted = extract_skills_data(df_mail_data)
        
        # 最終出力
        status_label.config(text="状態: 結果ファイルを書き出し中...", fg="blue")
        df_output = df_extracted.copy()
        output_file_abs_path = os.path.abspath(OUTPUT_FILENAME)

        # EntryIDをURLに変換（Python完結の生URL方式）
        df_output.insert(0, 'メールURL', df_output.apply(lambda row: f"outlook:{row['EntryID']}", axis=1))

        # 最終出力列の整理
        df_output = df_output.drop(columns=['EntryID', '宛先メール', '本文(テキスト形式)'], errors='ignore')

        # pandasでベースデータ(.xlsx)を生成
        df_output.to_excel(output_file_abs_path, index=False)
        messagebox.showinfo("完了", f"抽出処理が正常に完了し、\n'{OUTPUT_FILENAME}' に出力されました。")
        status_label.config(text=f"状態: 処理完了。ファイル出力済み。", fg="green")
        
        # ★★★ 修正: 検索ボタンを有効化 ★★★
        search_button = main_elements.get("search_button")
        if search_button:
            search_button.config(state=tk.NORMAL)
            

        
        # 以前のコードにあった gui_search_window.open_search_window(root) の呼び出しも削除

    except Exception as e:
        status_label.config(text=f"状態: エラー発生 - {e}", fg="red")
        messagebox.showerror("エラー", f"抽出処理中にエラーが発生しました: {e}")
        
    finally:
        pythoncom.CoUninitialize()


def run_extraction_thread(root, main_elements):
    """GUIをブロックしないよう、抽出処理を別スレッドで実行するラッパー。"""
    account_email = main_elements["account_entry"].get().strip()
    folder_path = main_elements["folder_entry"].get().strip()
    status_label = main_elements["status_label"]
    
    if not account_email or not folder_path:
        messagebox.showerror("入力エラー", "メールアカウントとフォルダパスの入力は必須です。")
        return

    thread = threading.Thread(target=lambda: actual_run_extraction_logic(root, main_elements, account_email, folder_path, status_label))
    thread.start()
    
# ----------------------------------------------------
# メイン実行関数 (GUI起動)
# ----------------------------------------------------

def main():
    """
    アプリケーションのメインウィンドウを作成し、実行する。
    """
    # 1. メインウィンドウの設定
    root = tk.Tk()
    root.title("Outlook Mail Search Tool")
    root.geometry("800x600")

    # 2. 初期設定データの読み込み
    saved_account, saved_folder = utils.load_config_csv() 

    # 3. メインフレームと設定フレームの作成
    main_frame = Frame(root)
    main_frame.pack(padx=10, pady=10, fill='both', expand=True)
    
    setting_frame = Frame(main_frame)
    setting_frame.pack(padx=10, pady=10, fill='x')

    # 4. コールバック関数の定義
    
    main_elements = {} 
    
    def open_settings_callback():
        gui_elements.open_settings_window(
            root, main_elements["account_entry"], main_elements["status_label"]
        )

    def open_search_callback():
        gui_search_window.open_search_window(root)
        
    def run_extraction_callback():
        run_extraction_thread(root, main_elements)

    # 5. GUI要素の作成
    elements_dict = gui_elements.create_main_window_elements(
        root,
        setting_frame=setting_frame,
        saved_account=saved_account,
        saved_folder=saved_folder,
        run_extraction_callback=run_extraction_callback,
        open_settings_callback=open_settings_callback,
        open_search_callback=open_search_callback
    )
    
    # 辞書の内容を main_elements にコピー
    main_elements.update(elements_dict)
    
    # 6. アプリケーションの開始
    root.mainloop()

if __name__ == "__main__":
    main()