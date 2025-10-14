# data_processor.py
#正規表現、データ管理 フィルタリングの機能
import pandas as pd
import re
import os
import unicodedata
from gui_config import ITEM_PATTERNS,SCRIPT_DIR, OUTPUT_CSV_FILE
from gui_utils import clean_and_normalize

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

def safe_to_int(value):
    """単金や年齢の文字列を安全に整数に変換するヘルパー関数"""
    if pd.isna(value) or value is None: return None
    value_str = str(value).strip()
    if not value_str: return None 
    try:
        # 文字列のクリーンアップと正規化
        cleaned_str = re.sub(r'[\s　\xa0\u3000]+', '', value_str) 
        normalized_str = unicodedata.normalize('NFKC', cleaned_str)
        # 不要な文字を除去 
        cleaned_str = normalized_str.replace(',', '').replace('万円', '').replace('歳', '').strip()
        # 数字と小数点以外を除去（小数点以下も許可）
        cleaned_str = re.sub(r'[^\d\.]', '', cleaned_str) 
        
        # cleaned_strが空文字列になった場合はNoneを返す
        if not cleaned_str: return None

        # 浮動小数点数（例: 70.5）として解釈し、小数点以下を切り捨てて整数に変換
        return int(float(cleaned_str))
        
    except ValueError:
        return None
    except Exception:
        return None 
    
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
    
    # is_matchは最終的なAND条件の結果を保持する
    is_match = pd.Series(True, index=df.index) 
    column_series = df[column_name].astype(str)
    
    # 1. チェックボックスフィルタ（AND条件）
    if selected_items:
        delimiter_chars = r'[\s,、/・]'
        
        # ★★★ 修正箇所: 選択された各項目をループでAND結合 ★★★
        for item in selected_items:
            escaped_item = re.escape(item)
            
            # パターン: (行頭 or 区切り文字) + 項目 + (区切り文字 or 行末)
            # 選択された項目が、区切り文字に囲まれた単語として含まれるかをチェック
            pattern = r'(?:^|' + delimiter_chars + r')' + escaped_item + r'(?:' + delimiter_chars + r'|$)'

            # 現在のitemのマッチ結果を取得 (大文字小文字無視)
            current_item_match = column_series.str.contains(pattern, na=False, flags=re.IGNORECASE, regex=True)
            
            # これまでの結果と現在のマッチ結果をANDで結合
            is_match = is_match & current_item_match
        # ★★★ 修正箇所 終わり ★★★

    # 2. 手動キーワードフィルタ（AND条件）
    if keyword_list:
        # カンマ区切りで入力された各キーワードについて、全てのマッチを要求 (AND条件)
        for keyword in keyword_list:
            escaped_keyword = re.escape(keyword)
            # キーワードが文字列中のどこかに含まれているかチェック (大文字小文字無視)
            keyword_match = column_series.str.contains(escaped_keyword, na=False, flags=re.IGNORECASE, regex=True)
            is_match = is_match & keyword_match # ANDで結合

    return df[is_match]
