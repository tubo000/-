# gui_callbacks.py
# 責務: GUIからの入力を受け取り、Outlook接続・抽出・評価のワークフローをスレッドで実行する。（UIフリーズ防止）

import tkinter as tk
from tkinter import messagebox
import pandas as pd 
import os
import threading # スレッド処理用
import pythoncom # 📌 COM初期化/解放用 (Outlook操作で必須)

# 📌 修正: インポートパスを現在のファイル構成に合わせる
from config import SCRIPT_DIR, INTERMEDIATE_CSV_FILE, MASTER_ANSWERS_PATH, OUTPUT_EVAL_PATH
from utils import save_config_csv
from extraction_core import extract_skills_data
# from evaluator_core import run_triple_csv_validation # 評価関数はここでは使用しない
from outlook_api import get_mail_data_from_outlook_in_memory


def _actual_run_extraction_logic(root, account_name, target_folder, status_label, search_button):
    """
    抽出のメインロジック。別スレッドで実行される。
    """
    # 📌 修正1: Tkinterの更新を安全に行うヘルパー関数
    def update_ui_status(text, color, message_type=None, message_text=None):
        """ステータス更新やメッセージボックス表示をメインスレッドに渡す"""
        if message_type == "error":
            root.after(0, lambda: messagebox.showerror("エラー", message_text))
        elif message_type == "info":
            root.after(0, lambda: messagebox.showinfo("完了", message_text))
            # 完了時のみ検索ボタンを有効化 (データができたことを示す)
            root.after(0, lambda: search_button.config(state=tk.NORMAL)) 
        
        # ステータスラベルの更新
        root.after(0, lambda: status_label.config(text=text, fg=color))

    # Outlook/COMオブジェクトを使う前に、必ずカレントスレッドで初期化を行う
    try:
        pythoncom.CoInitialize()
    except Exception as e:
        update_ui_status("状態: COM初期化エラー", "red", "error", f"COM初期化に失敗しました。詳細: {e}")
        return

    try:
        intermediate_path = os.path.join(SCRIPT_DIR, INTERMEDIATE_CSV_FILE)
        
        # 1. Outlook接続とメールデータ取得
        update_ui_status(f" ステップ1/3: Outlookアカウント '{account_name}' からメール取得中...", "blue")
        # 📌 修正: 引数の順番を Outlook API に合わせる
        df_mail_data = get_mail_data_from_outlook_in_memory(target_folder, account_name) 
        
        if df_mail_data.empty:
            update_ui_status("状態: 処理対象のメールがありませんでした。", "green")
            return
            
        # 2. メール本文と添付ファイルの内容結合 (処理のステータス更新)
        update_ui_status(f" ステップ2/3: {len(df_mail_data)}件のメール本文と添付ファイルの内容を結合中...", "blue")
        mail_df = df_mail_data 
        
        # 3. メール本文からスキル情報を抽出
        update_ui_status(" ステップ3/3: スキル情報を抽出中...", "blue")
        extracted_df = extract_skills_data(mail_df)
        
        # 4. 結果の出力
        output_file_abs_path = os.path.join(SCRIPT_DIR, OUTPUT_EVAL_PATH) 
        
        # 📌 修正2: 抽出結果を XLSX ファイルとして出力 (ファイル名・拡張子は環境に合わせて調整)
        extracted_df.to_excel(output_file_abs_path, index=False, encoding='utf-8-sig') 
        
        new_skill_count = len(extracted_df)
        
        update_ui_status(f" 処理完了。{new_skill_count}件のスキルシートを抽出しました。", "green", "info", f"抽出処理が正常に完了し、\n'{os.path.basename(output_file_abs_path)}' に出力されました。")
        
    except RuntimeError as e:
        # RuntimeError (Outlookエラーなど) の処理
        update_ui_status(" 処理中にエラーが発生しました。", "red", "error", f"Outlook接続エラー: {e}")
    except Exception as e:
        # その他の予期せぬエラー処理
        update_ui_status(" 予期せぬエラー。", "red", "error", f"予期せぬエラーが発生しました: \n\n【詳細】{e}")
    finally:
        # COMオブジェクトの使用を終えたら、必ず解放する
        pythoncom.CoUninitialize()


def run_extraction_workflow(root, account_entry, folder_entry, status_label, search_button):
    """
    「抽出を実行」ボタンが押された際のメインワークフロー。
    別スレッドで実行することで、GUIのフリーズを防ぐ。
    """
    
    account_name = account_entry.get().strip()
    target_folder = folder_entry.get().strip()
    
    # 必須入力チェック
    if not account_name or not target_folder:
        error_message = " Outlookアカウントとフォルダパスの入力は必須です。"
        status_label.config(text=error_message, fg="red"); messagebox.showerror("入力エラー", error_message); return
    
    # ユーザー設定の保存
    save_config_csv(account_name, target_folder)        
    
    # 実行前にステータスをリセット
    status_label.config(text=" 状態: 処理開始準備中...", fg="black")
    search_button.config(state=tk.DISABLED) # 処理中は検索ボタンを無効化
    
    # 📌 修正3: メインの抽出ロジックを別スレッドで実行
    extraction_thread = threading.Thread(
        target=_actual_run_extraction_logic, 
        args=(root, account_name, target_folder, status_label, search_button)
    )
    extraction_thread.start()