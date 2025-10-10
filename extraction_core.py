# extraction_core.py

import pandas as pd
import re
from config import MASTER_COLUMNS, ITEM_PATTERNS, PROCESS_KEYWORDS


def clean_and_normalize(value: str, item_name: str) -> str:
    """抽出結果をクリーンアップし、正規化する関数。（ノイズ除去を含む）"""
    if not value or value.strip() == '': return 'N/A'
    cleaned = value.strip().replace('\xa0', ' ')
    cleaned = re.sub(r'[\s\u3000]+', ' ', cleaned).strip() 
    
    if item_name == '名前' or item_name == '氏名':
        # 括弧や記号で囲まれたノイズを、中身ごと一括で削除
        cleaned = re.sub(r'[\(（\[【].*?[\)）\]】]', '', cleaned) 
        
        # 名前の区切りに使われるノイズ文字（・、_）をスペースに変換し、末尾のハイフンを削除
        cleaned = re.sub(r'[・\_]', ' ', cleaned)
        cleaned = re.sub(r'[-]+$', '', cleaned).strip() 
        
        # 氏名内の全てのスペースを削除し、連結
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
        
        cleaned = re.sub(r'[・、/\\|,;]', ',', cleaned) 
        cleaned = re.sub(r'\s*,\s*', ',', cleaned).strip(',')
        
    return cleaned


def extract_skills_data(mail_data_df: pd.DataFrame) -> pd.DataFrame:
    """メールデータDataFrameを受け取り、抽出結果と信頼度スコアを返す。"""
    
    # 抽出前に本文をクリーンアップ (構造崩壊防止)
    mail_data_df['本文(テキスト形式)'] = mail_data_df['本文(テキスト形式)'].str.replace(r'[\r\n\t]', ' ', regex=True)
    mail_data_df['本文(テキスト形式)'] = mail_data_df['本文(テキスト形式)'].str.replace(r'\s+', ' ', regex=True) 

    all_extracted_rows = []
    
    for index, row in mail_data_df.iterrows():
        mail_id = str(row.get('EntryID', f'Row_{index+1}'))
        full_mail_text = str(row.get('本文(テキスト形式)', '')) 
        
        full_text_for_search = full_mail_text
        
        extracted_data = {'EntryID': mail_id, '件名': row.get('件名', 'N/A'), '宛先メール': row.get('宛先メール', 'N/A')} 
        reliability_scores = {} 
        
        for item_key, pattern_info in ITEM_PATTERNS.items():
            pattern = pattern_info['pattern'] 
            base_item_name = item_key.split('_')[0]
            
            flags = re.IGNORECASE
            if item_key == 'スキルor言語':
                flags |= re.DOTALL 

            match = re.search(pattern, full_text_for_search, flags)
            
            if match:
                extracted_value = match.group(1)
                score = pattern_info['score']
                
                cleaned_val = clean_and_normalize(extracted_value, base_item_name)
                
                current_score = reliability_scores.get(base_item_name, 0)
                if score > current_score:
                    extracted_data[base_item_name] = cleaned_val
                    reliability_scores[base_item_name] = score
            
        for proc_name, keywords in PROCESS_KEYWORDS.items():
            flag_col = f'開発工程_{proc_name}'
            extracted_data[flag_col] = 'なし' 
            if re.search('|'.join(keywords), full_mail_text, re.IGNORECASE):
                extracted_data[flag_col] = 'あり' 

        for col in MASTER_COLUMNS:
              if col not in extracted_data:
                  if col.startswith('開発工程_'):
                      extracted_data[col] = 'なし' 
                  else:
                      extracted_data[col] = 'N/A' 

        valid_scores = [s for s in reliability_scores.values() if s > 0]
        extracted_data['信頼度スコア'] = round(sum(valid_scores) / len(valid_scores) if valid_scores else 0, 1)
        extracted_data['本文(テキスト形式)'] = full_mail_text 

        all_extracted_rows.append(extracted_data)
            
    df_extracted = pd.DataFrame(all_extracted_rows)
    df_extracted = df_extracted.reindex(columns=MASTER_COLUMNS, fill_value='N/A')
    df_extracted = df_extracted.astype(str)
    
    return df_extracted