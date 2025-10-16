# extraction_core.py
# 責務: メール本文からの情報抽出（正規表現適用）と、GUIで使用されるデータフィルタリングロジックの提供。

import pandas as pd
import re
from config import MASTER_COLUMNS, ITEM_PATTERNS, PROCESS_KEYWORDS


def clean_and_normalize(value: str, item_name: str) -> str:
    """抽出結果をクリーンアップし、正規化する関数。（ノイズ除去を含む）"""
    # 📌 修正1: そもそも抽出結果が空文字列の場合は、すぐに'N/A'を返す
    if not value or not value.strip(): 
        return 'N/A'
    
    cleaned = value.strip().replace('\xa0', ' ')
    cleaned = re.sub(r'[\s\u3000]+', ' ', cleaned).strip() 
    
    if item_name == '名前' or item_name == '氏名':
        cleaned = re.sub(r'[\(（\[【].*?[\)）\]】]', '', cleaned) 
        cleaned = re.sub(r'[・\_]', ' ', cleaned)
        cleaned = re.sub(r'[-]+$', '', cleaned).strip() 
        cleaned = re.sub(r'\s+', '', cleaned).strip()
    
    if item_name == '年齢':
        numeric_val = re.sub(r'[\D\.,]+', '', cleaned)
        if numeric_val.isdigit():
            age = int(numeric_val)
            if 18 <= age <= 100:
                return str(age) 
            else:
                return 'N/A' 
        return 'N/A' 
    
    if item_name == '単金':
        cleaned = value.strip()
        numeric_val = re.sub(r'[\D\.,]+', '', cleaned)
        
        if numeric_val.isdigit():
            num = int(numeric_val)
            return str(num * 10000)
        return 'N/A' 

    if item_name in ['マネジメント経験人数', '人数']:
        return re.sub(r'[\D\.,]+', '', cleaned) 
        
    if item_name in ['スキルor言語', 'OS', 'データベース', 'フレームワーク/ライブラリ', '開発ツール']:
        cleaned = re.sub(r'^【\s*言\s*語\s*】|^【\s*DB\s*】|^【\s*OS\s*】', '', cleaned, flags=re.IGNORECASE)
        cleaned = cleaned.strip() 
        
        # 連続する区切り文字を単一のカンマに変換
        cleaned = re.sub(r'[・、/\\|,;]+', ',', cleaned) 
        # カンマの周りの空白を削除し、文字列の先頭・末尾のカンマを削除
        cleaned = re.sub(r'\s*,\s*', ',', cleaned).strip(',') 
        
        # 📌 修正2: クリーニング後に空になった場合も 'N/A' を返す
        if not cleaned:
            return 'N/A'
        
    return cleaned


def extract_skills_data(mail_data_df: pd.DataFrame) -> pd.DataFrame:
    """メールデータDataFrameを受け取り、抽出結果と信頼度スコアを返す。"""
    
    # 1. 抽出に使用する全文を決定
    SOURCE_TEXT_COL = '本文(抽出元結合)'
    if SOURCE_TEXT_COL not in mail_data_df.columns:
        SOURCE_TEXT_COL = '本文(テキスト形式)'
        
    # 抽出に使用する全文をクリーンアップ (念のため)
    mail_data_df[SOURCE_TEXT_COL] = mail_data_df[SOURCE_TEXT_COL].astype(str).str.replace(r'[\r\n\t"]', ' ', regex=True)
    mail_data_df[SOURCE_TEXT_COL] = mail_data_df[SOURCE_TEXT_COL].astype(str).str.replace(r'\s+', ' ', regex=True)
    
    all_extracted_rows = []
    
    for index, row in mail_data_df.iterrows():
        mail_id = str(row.get('EntryID', f'Row_{index+1}'))
        # 📌 修正3: 正規表現の抽出に使う全文から、エラーメッセージを除去してノイズを減らす
        full_text_for_search = str(row.get(SOURCE_TEXT_COL, ''))
        full_text_for_search = re.sub(r'\[FILE (ERROR|WARN):.*?\]', '', full_text_for_search)


        # GUI表示用/最終出力用に元の本文を分離して取得
        mail_body_for_display = str(row.get('本文(テキスト形式)', 'N/A'))
        file_body_for_display = str(row.get('本文(ファイル含む)', 'N/A'))

        extracted_data = {'EntryID': mail_id, '件名': row.get('件名', 'N/A'), '宛先メール': row.get('宛先メール', 'N/A')} 
        reliability_scores = {} 
        
        # --- 2. 正規表現による抽出 ---
        for item_key, pattern_info in ITEM_PATTERNS.items():
            pattern = pattern_info['pattern'] 
            base_item_name = item_key.split('_')[0]
            
            flags = re.IGNORECASE
            if item_key == 'スキルor言語': flags |= re.DOTALL 

            match = re.search(pattern, full_text_for_search, flags)
            
            if match:
                # 抽出されたグループが複数ある場合も、最初のグループ(1)を使用
                extracted_value = match.group(1) if match.groups() else match.group(0)
                score = pattern_info['score']
                
                cleaned_val = clean_and_normalize(extracted_value, base_item_name)
                
                current_score = reliability_scores.get(base_item_name, 0)
                if score > current_score:
                    extracted_data[base_item_name] = cleaned_val
                    reliability_scores[base_item_name] = score
            
        # --- 3. 開発工程フラグの判定 ---
        for proc_name, keywords in PROCESS_KEYWORDS.items():
            flag_col = f'開発工程_{proc_name}'
            extracted_data[flag_col] = 'なし' 
            if re.search('|'.join(keywords), full_text_for_search, re.IGNORECASE):
                extracted_data[flag_col] = 'あり' 

        # --- 4. 最終データの構築と補完 ---
        final_row = {} 
        final_row.update(extracted_data) 

        # 出力データに分離した本文を格納
        final_row['本文(テキスト形式)'] = mail_body_for_display
        final_row['本文(ファイル含む)'] = file_body_for_display 
        
        # 📌 修正4: Attachments の型をチェックし、確実にカンマ区切り文字列に変換する
        attachments_list = row.get('Attachments', [])
        if isinstance(attachments_list, list):
            final_row['Attachments'] = ', '.join(attachments_list)
        else:
            final_row['Attachments'] = str(attachments_list)


        # 信頼度スコアの計算
        valid_scores = [s for s in reliability_scores.values() if s > 0]
        final_row['信頼度スコア'] = round(sum(valid_scores) / len(valid_scores) if valid_scores else 0, 1)

        all_extracted_rows.append(final_row)
            
    # 最終DataFrameの構築
    df_extracted = pd.DataFrame(all_extracted_rows)
    
    # 抽出に使用した結合カラムを DataFrame から削除 (最終出力に含めない)
    if '本文(抽出元結合)' in df_extracted.columns:
        df_extracted = df_extracted.drop(columns=['本文(抽出元結合)']) 
        
    # MASTER_COLUMNS + 制御カラムの順に列を並び替え、欠損をN/Aで埋める
    all_cols_in_order = [
        'EntryID', '件名', '宛先メール', '本文(テキスト形式)', '本文(ファイル含む)', 
        'Attachments', '信頼度スコア'
    ] + [col for col in MASTER_COLUMNS if col not in ['EntryID', '件名', '宛先メール', '本文(テキスト形式)', '本文(ファイル含む)', 'Attachments', '信頼度スコア', '本文(抽出元結合)']]
    
    df_extracted = df_extracted.reindex(columns=[c for c in all_cols_in_order if c in df_extracted.columns], fill_value='N/A')
    df_extracted = df_extracted.astype(str)
    
    return df_extracted

# ----------------------------------------------------
# GUIフィルタリングロジック (省略)
# ----------------------------------------------------

def apply_checkbox_filter(df, column_name, selected_items, keyword_list):
    """GUIの検索ウィンドウで使用されるデータフィルタリングロジック。"""
    if not selected_items and not keyword_list:
        return df
    if column_name not in df.columns:
        return df 
    
    is_match = pd.Series(True, index=df.index) 
    column_series = df[column_name].astype(str)
    
    if selected_items:
        delimiter_chars = r'[\s,、/・]'
        for item in selected_items:
            escaped_item = re.escape(item)
            pattern = r'(?:^|' + delimiter_chars + r')' + escaped_item + r'(?:' + delimiter_chars + r'|$)'
            current_item_match = column_series.str.contains(pattern, na=False, flags=re.IGNORECASE, regex=True)
            is_match = is_match & current_item_match
        
    if keyword_list:
        for keyword in keyword_list:
            escaped_keyword = re.escape(keyword)
            keyword_match = column_series.str.contains(escaped_keyword, na=False, flags=re.IGNORECASE, regex=True)
            is_match = is_match & keyword_match 
            
    return df[is_match]