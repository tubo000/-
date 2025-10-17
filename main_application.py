# main_application.py (GUI統合とメイン実行フロー)

import os
import sys
import pandas as pd
import win32com.client as win32
import threading 
import tkinter as tk
from tkinter import Frame, messagebox, simpledialog 
import pythoncom 
import re 
import traceback 
import os.path

# 外部モジュールのインポート
import gui_elements
# 📌 修正1: gui_search_window.py をインポート
import gui_search_window 
import utils 

# 既存の内部処理関数をインポート
from config import INPUT_QUESTION_CSV, MASTER_ANSWERS_PATH, OUTPUT_EVAL_PATH, NUM_RECORDS, TARGET_FOLDER_PATH, SCRIPT_DIR
from data_generation import generate_raw_data, export_dataframes_to_tsv
from extraction_core import extract_skills_data
from evaluator_core import run_triple_csv_validation, get_question_data_from_csv
from email_processor import run_email_extraction, get_mail_data_from_outlook_in_memory, OUTPUT_FILENAME


# ----------------------------------------------------
# ユーティリティ関数群 (open_outlook_email_by_id, interactive_id_search_test は維持)
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


# 📌 修正2: 循環参照を避けるため、reorder_output_dataframe をローカルで定義
def reorder_output_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """出力データフレームの列順を調整し、特定の項目を左側に固定する。"""
    fixed_leading_cols = [
        'メールURL', '件名', '名前', '信頼度スコア', 
        '本文(ファイル含む)', '本文(テキスト形式)', 'Attachments'
    ]
    fixed_leading_cols = [col for col in fixed_leading_cols if col in df.columns]
    remaining_cols = [col for col in df.columns.tolist() if col not in fixed_leading_cols]
    return df.reindex(columns=fixed_leading_cols + remaining_cols, fill_value='N/A')


def actual_run_extraction_logic(root, main_elements, target_email, folder_path, status_label):
    
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
        
        # 最終出力前のデータ準備
        df_output = df_extracted.copy()
        
        # EntryIDをURLに変換
        if 'EntryID' in df_output.columns and 'メールURL' not in df_output.columns:
             df_output.insert(0, 'メールURL', df_output.apply(lambda row: f"outlook:{row['EntryID']}", axis=1))

        # 列順の整理
        df_output = reorder_output_dataframe(df_output)
        
        # 不要な列の最終削除 (EntryID, 宛先メールなど)
        final_drop_list = ['EntryID', '宛先メール', '本文(抽出元結合)'] 
        final_drop_list = [col for col in final_drop_list if col in df_output.columns]
        df_output = df_output.drop(columns=final_drop_list, errors='ignore')
        
        # Excel修復ログ (数式) エラー対策
        df_output = df_output.astype(str)
        for col in df_output.columns:
            df_output[col] = df_output[col].str.replace(r'^=', r"'=", regex=True)

        # 📌 修正3: Excelファイルに出力
        output_file_abs_path = os.path.abspath(OUTPUT_FILENAME)
        df_output.to_excel(output_file_abs_path, index=False) 

        messagebox.showinfo("完了", f"抽出処理が正常に完了し、\n'{OUTPUT_FILENAME}' に出力されました。\n検索一覧ボタンを押して結果を確認してください。")
        status_label.config(text=f"状態: 処理完了。ファイル出力済み。", fg="green")
        
        # 📌 修正4: 抽出完了後、検索ボタンを有効化
        search_button = main_elements.get("search_button")
        if search_button:
            search_button.config(state=tk.NORMAL)
        
        
    except Exception as e:
        status_label.config(text=f"状態: エラー発生 - {e}", fg="red")
        messagebox.showerror("エラー", f"抽出処理中に予期せぬエラーが発生しました。\n詳細: {e}")
        
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
        
    # 📌 修正5: 検索一覧ボタンのコールバック関数 - ファイル読み込みでUIを起動
    def open_search_callback():
        output_file_abs_path = os.path.abspath(OUTPUT_FILENAME)
        
        if not os.path.exists(output_file_abs_path):
            messagebox.showwarning("警告", f"抽出結果ファイル ('{OUTPUT_FILENAME}') が見つかりません。\n先に抽出を実行してください。")
            return
            
        try:
            # メインウィンドウを隠す
            root.withdraw() 
            
            # gui_search_window.py の main() 関数を呼び出す
            gui_search_window.main()
            
        except Exception as e:
            messagebox.showerror("検索ウィンドウ起動エラー", f"検索一覧の表示中に予期せぬエラーが発生しました。\n詳細: {e}")
            traceback.print_exc()
        finally:
            # 検索ウィンドウが閉じられたら、元のメインウィンドウを再表示
            root.deiconify() 


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