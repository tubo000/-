# data_processor.py
#抽出、正規表現、データ管理 フィルタリングのチェック欄の機能
import pandas as pd
import re
import os
import tkinter as tk
from tkinter import messagebox

from config import ITEM_PATTERNS,SCRIPT_DIR, OUTPUT_CSV_FILE,INTERMEDIATE_CSV_FILE
from utils import clean_and_normalize, save_config_csv
from outlook_api import get_mail_data_from_outlook_in_memory 

def extract_skills_data(mail_data_df: pd.DataFrame) -> pd.DataFrame:
    """メール本文からスキルデータを抽出する。（正規表現マッチング）"""
    all_extracted_rows = []
    major_items = ['氏名', '年齢', '単金', '業務_業種', 'スキル_言語', 'スキル_OS'] 
    
    for index, row in mail_data_df.iterrows():
        mail_id = str(row.get('EntryID', f'Row_{index+1}'))
        full_mail_text = str(row.get('本文(テキスト形式)', ''))
        extracted_data = {
            'EntryID': mail_id, '件名': row.get('件名', 'N/A'), '__Source_Mail__': row.get('__Source_Mail__', 'N/A'),
            '本文(テキスト形式)': row.get('本文(テキスト形式)', 'N/A'), 'Attachments': row.get('Attachments', []) 
        }
        reliability_scores = {}
        
        # 項目ごとに、定義されたパターンリストを試行
        for item_key, patterns_list in ITEM_PATTERNS.items():
            base_item_name = item_key 
            best_match_value = None
            best_score = 0
            
            # 各項目について、定義されたパターンリストをスコアが高い順に試行
            for pattern_info in patterns_list:
                pattern = pattern_info['pattern']
                score = pattern_info['score']
                
                # すでに十分なスコアでマッチしている場合はスキップ
                if best_score >= score and best_match_value is not None: continue

                match = re.search(pattern, full_mail_text, re.IGNORECASE | re.MULTILINE)
                if match:
                    extracted_value = match.group(1)
                    cleaned_val = clean_and_normalize(extracted_value, base_item_name)
                    
                    # より高いスコア、または初めての有効な値であれば採用
                    if score > best_score or (best_match_value is None and cleaned_val != 'N/A'):
                        best_match_value = cleaned_val
                        best_score = score
                        
            if best_match_value and best_match_value != 'N/A':
                extracted_data[item_key] = best_match_value
                reliability_scores[item_key] = best_score
        
        # 総合スコア計算（抽出できた項目の平均スコア）
        valid_scores = [s for s in reliability_scores.values() if s > 0]
        overall_score = sum(valid_scores) / len(valid_scores) if valid_scores else 0
        extracted_data['信頼度スコア'] = round(overall_score, 1)
        
        # 抽出されなかった主要項目に 'N/A' を設定
        for key in major_items:
            if key not in extracted_data: extracted_data[key] = 'N/A'
        
        all_extracted_rows.append(extracted_data)
        
    return pd.DataFrame(all_extracted_rows)
def run_evaluation_feedback_and_output(extracted_df: pd.DataFrame):
    """抽出結果を最終CSVファイルに出力する。"""
    try:
        output_path = os.path.join(SCRIPT_DIR, OUTPUT_CSV_FILE)
        # 本文と添付ファイルのカラムを除外
        output_df = extracted_df.drop(columns=['本文(テキスト形式)', 'Attachments'], errors='ignore')
        # 信頼度スコアを末尾に移動して、BOM付きUTF-8でCSVに出力
        output_cols = [col for col in output_df.columns if col != '信頼度スコア'] + ['信頼度スコア']
        output_df = output_df.reindex(columns=output_cols)
        output_df.to_csv(output_path, index=False, encoding='utf-8-sig')
        return len(output_df), f" 処理が完了しました！\n最終結果を '{OUTPUT_CSV_FILE}' に保存しました。"
    except Exception as e:
        return 0, f" 最終CSV出力エラー: {e}"
def run_extraction_workflow(root, account_entry, folder_entry, status_label, search_button):
    """「抽出を実行」ボタンが押された際のメインワークフロー"""
    # ★ 'global search_button, status_label' の行を削除 ★
    
    # account_entry と folder_entry は既に引数で受け取っている
    account_name = account_entry.get().strip(); target_folder = folder_entry.get().strip()
    
    # 必須入力チェック (ここでの status_label は引数で受け取ったもの)
    if not account_name:
        error_message = " Outlookアカウント（メールアドレスまたは表示名）の入力は必須です。"
        status_label.config(text=error_message, fg="red"); messagebox.showerror("入力エラー", error_message); return
    if not target_folder:
        error_message = " 対象フォルダパスの入力は必須です。"
        status_label.config(text=error_message, fg="red"); messagebox.showerror("入力エラー", error_message); return

    save_config_csv(account_name) # アカウント名を保存
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

def toggle_all_checkboxes(vars_dict, select_state, update_func):
    """全てのチェックボックスの状態を切り替え、テーブルを更新する"""
    for var in vars_dict.values():
        var.set(select_state)
    update_func()
#すべてのチェックボックスの機能
def apply_checkbox_filter(df, column_name, selected_items, keyword_list):
    """DataFrameにチェックボックスと手動キーワードによるANDフィルタを適用する。（AND条件）"""
    # 項目が選択されていない場合は、全てのデータをそのまま返す
    if not selected_items and not keyword_list:
        return df
    if column_name not in df.columns:
        return df 
    
    is_match = pd.Series(True, index=df.index) 
    column_series = df[column_name].astype(str)
    
    # 1. チェックボックスフィルタ（OR条件）
    if selected_items:
        # 選択された項目のいずれかにマッチするかどうか (OR条件)
        # パターン: (行頭 or 区切り文字) + 項目 + (区切り文字 or 行末)
        delimiter_chars = r'[\s,、/・]'
        or_patterns = [r'(?:^|' + delimiter_chars + r')' + re.escape(item) + r'(?:' + delimiter_chars + r'|$)' for item in selected_items]
        # 全てのORパターンを結合して一つの正規表現にする
        or_regex = '|'.join(or_patterns)
        
        # OR条件のマッチ結果を取得 (一つでもマッチすればTrue)
        is_or_match = column_series.str.contains(or_regex, na=False, flags=re.IGNORECASE, regex=True)
        is_match = is_match & is_or_match

    # 2. 手動キーワードフィルタ（AND条件）
    if keyword_list:
        # カンマ区切りで入力された各キーワードについて、全てのマッチを要求 (AND条件)
        for keyword in keyword_list:
            escaped_keyword = re.escape(keyword)
            # キーワードが文字列中のどこかに含まれているかチェック (大文字小文字無視)
            keyword_match = column_series.str.contains(escaped_keyword, na=False, flags=re.IGNORECASE, regex=True)
            is_match = is_match & keyword_match # ANDで結合

    return df[is_match]

