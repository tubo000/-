# gui_callbacks.py
#抽出ボタンのフロー

import tkinter as tk
from tkinter import messagebox
import pandas as pd 
import os
from gui_config import SCRIPT_DIR ,INTERMEDIATE_CSV_FILE
from gui_utils import save_config_csv
from gui_data_processor import  extract_skills_data , run_evaluation_feedback_and_output
from outlook_api import get_mail_data_from_outlook_in_memory

def run_extraction_workflow(root, account_entry, folder_entry, status_label, search_button):
    """「抽出を実行」ボタンが押された際のメインワークフロー"""
    # ★ ここにあった 'global search_button, status_label' は削除する ★
    
    account_name = account_entry.get().strip(); target_folder = folder_entry.get().strip()
    
    # 必須入力チェック（status_labelは引数で渡されたものを使う）
    if not account_name:
        error_message = " Outlookアカウント（メールアドレスまたは表示名）の入力は必須です。"
        status_label.config(text=error_message, fg="red"); messagebox.showerror("入力エラー", error_message); return
    if not target_folder:
        error_message = " 対象フォルダパスの入力は必須です。"
        status_label.config(text=error_message, fg="red"); messagebox.showerror("入力エラー", error_message); return

    save_config_csv(account_name) # アカウント名を保存
    # ステータス表示の更新（status_labelは引数で渡されたものを使う）
    status_label.config(text=" ステップ1/3: Outlookからメールの取得を開始...", fg="blue"); root.update()
    
    mail_df = pd.DataFrame(); new_skill_count = 0
    
    try:
        # 1. Outlookからメールデータを取得
        mail_df = get_mail_data_from_outlook_in_memory(target_folder, account_name)
        if mail_df.empty:
            message = " 処理完了。スキルシート候補のメールは見つかりませんでした。"
            status_label.config(text=message, fg="green"); messagebox.showinfo("完了", message + f"\n\n(参照フォルダ: {target_folder})")
            if search_button: search_button.config(state=tk.DISABLED); return

        status_label.config(text=" ステップ2/3: 取得したメールデータを一時ファイルに出力中...", fg="blue"); root.update()
        intermediate_path = os.path.join(SCRIPT_DIR, INTERMEDIATE_CSV_FILE)
        mail_df.to_csv(intermediate_path, index=False, encoding='utf-8-sig')

        status_label.config(text=" ステップ3/3: メール本文からスキル情報を抽出中...", fg="blue"); root.update()
        
        extracted_df = extract_skills_data(mail_df)
        skill_count, message = run_evaluation_feedback_and_output(extracted_df)
        new_skill_count = skill_count
        
    except RuntimeError as e:
        status_label.config(text=" 処理中にエラーが発生しました。", fg="red"); messagebox.showerror("Outlook接続エラー", str(e)); return
    except Exception as e:
        status_label.config(text=" 予期せぬエラー。", fg="red"); messagebox.showerror("エラー", f"予期せぬエラーが発生しました: \n\n【詳細】{e}"); return

    final_message = f" 処理が完了しました！\n\n対象メール総数: {len(mail_df)}件\nスキルシート抽出件数: {new_skill_count}件"
    status_label.config(text=final_message, fg="green")
    
    if new_skill_count > 0 and search_button: search_button.config(state=tk.NORMAL) 
    elif search_button: search_button.config(state=tk.DISABLED)

    if new_skill_count > 0:
        messagebox.showinfo("完了", final_message + "\n\n「検索・結果一覧表示」ボタンを押して結果を確認してください。")
    else:
        messagebox.showinfo("完了", final_message)
