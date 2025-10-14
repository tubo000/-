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
    
    # 1. チェックボックスフィルタ（AND条件）
    if selected_items:
        delimiter_chars = r'[\s,、/・]'
        or_patterns = [r'(?:^|' + delimiter_chars + r')' + re.escape(item) + r'(?:' + delimiter_chars + r'|$)' for item in selected_items]
        # 全てのANDパターンを結合して一つの正規表現にする
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

