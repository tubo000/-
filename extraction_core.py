# extraction_core.py
# 責務: メール本文からの情報抽出（正規表現適用）と、抽出値のクリーンアップ・正規化を行う

import pandas as pd
import re
# configファイルから、定数（マスター列、正規表現パターン、工程キーワード）をインポート
from config import MASTER_COLUMNS, ITEM_PATTERNS, PROCESS_KEYWORDS


def clean_and_normalize(value: str, item_name: str) -> str:
    """
    正規表現で抽出された値を受け取り、項目名に応じてクリーンアップと正規化を行う。
    (例: 氏名のノイズ除去、単金の円換算)
    """
    # 値がNoneまたは空文字列の場合、'N/A'を返す
    if not value or value.strip() == '': return 'N/A'
    
    # 制御文字/全角スペースを統一し、前後のスペースを除去
    cleaned = value.strip().replace('\xa0', ' ')
    cleaned = re.sub(r'[\s\u3000]+', ' ', cleaned).strip() 
    
    # --- 項目ごとのクリーンアップロジック ---
    
    if item_name == '名前' or item_name == '氏名':
        # 1. 括弧（全角/半角/角括弧/波括弧）とその中の内容を削除 (例: (フリガナ) や【現職】を除去)
        cleaned = re.sub(r'[\(（\[【].*?[\)）\]】]', '', cleaned) 
        
        # 2. 名前の区切りに使われるノイズ文字（・、_）をスペースに変換し、末尾のハイフンを削除
        cleaned = re.sub(r'[・\_]', ' ', cleaned)
        cleaned = re.sub(r'[-]+$', '', cleaned).strip() 
        
        # 3. 氏名内の全てのスペースを削除し、連結 (例: 田中 太郎 -> 田中太郎)
        cleaned = re.sub(r'\s+', '', cleaned).strip()
    
    if item_name == '年齢':
        # 数字とカンマ、小数点、ハイフン以外の文字を除去
        numeric_val = re.sub(r'[\D\.,]+', '', cleaned)
        if numeric_val.isdigit():
            age = int(numeric_val)
            # 年齢が妥当な範囲 (18〜100歳) かチェック
            if 18 <= age <= 100:
                return str(age) 
            else:
                return 'N/A' 
        return 'N/A' 
    
    if item_name == '単金':
        cleaned = value.strip()
        # 数字とカンマ以外の文字を除去（万円表記の数字のみを残す）
        numeric_val = re.sub(r'[\D\.,]+', '', cleaned)
        
        if numeric_val.isdigit():
            num = int(numeric_val)
            # 抽出値（万円単位の数字）を円単位に変換して返す (例: 50 -> 500000)
            return str(num * 10000)
        return 'N/A' 

    if item_name in ['マネジメント経験人数', '人数']:
        # 数字のみを残す（評価時に円換算は不要）
        return re.sub(r'[\D\.,]+', '', cleaned) 
        
    if item_name in ['スキルor言語', 'OS', 'データベース', 'フレームワーク/ライブラリ', '開発ツール']:
        # 1. プレフィックス (例: 【言語】) を除去 (本文がゼロ化された後の残りを考慮)
        cleaned = re.sub(r'^【\s*言\s*語\s*】|^【\s*DB\s*】|^【\s*OS\s*】', '', cleaned, flags=re.IGNORECASE)
        cleaned = cleaned.strip() 
        
        # 2. 区切り文字（・、/など）をカンマに統一
        cleaned = re.sub(r'[・、/\\|,;]', ',', cleaned) 
        # 3. カンマ前後のスペースを除去し、不要なカンマを削除
        cleaned = re.sub(r'\s*,\s*', ',', cleaned).strip(',')
        
    return cleaned


def extract_skills_data(mail_data_df: pd.DataFrame) -> pd.DataFrame:
    """メールデータ DataFrame 全体を処理し、各項目を正規表現で抽出する。"""
    
    # --- 前処理 (構造崩壊防止のための改行・スペース除去) ---
    # セル内にある改行コード/タブ文字を全てスペースに置換 (DataFrame構造の維持)
    mail_data_df['本文(テキスト形式)'] = mail_data_df['本文(テキスト形式)'].str.replace(r'[\r\n\t]', ' ', regex=True)
    # 連続するスペースを一つに統一 (ゼロ化に近い処理)
    mail_data_df['本文(テキスト形式)'] = mail_data_df['本文(テキスト形式)'].str.replace(r'\s+', ' ', regex=True) 

    all_extracted_rows = []
    
    # データフレームの行ごとにループ
    for index, row in mail_data_df.iterrows():
        mail_id = str(row.get('EntryID', f'Row_{index+1}'))
        full_mail_text = str(row.get('本文(テキスト形式)', '')) 
        
        full_text_for_search = full_mail_text # 検索対象のクリーンな本文
        
        # 抽出結果を格納する辞書を初期化
        extracted_data = {'EntryID': mail_id, '件名': row.get('件名', 'N/A'), '宛先メール': row.get('宛先メール', 'N/A')} 
        reliability_scores = {} 
        
        # --- 正規表現による抽出 ---
        for item_key, pattern_info in ITEM_PATTERNS.items():
            pattern = pattern_info['pattern'] 
            base_item_name = item_key.split('_')[0]
            
            flags = re.IGNORECASE # 大文字小文字を無視
            # スキルor言語のみ、本文が複数行に渡る可能性を考慮し DOTALL フラグを付与
            if item_key == 'スキルor言語':
                flags |= re.DOTALL 

            # 正規表現検索を実行
            match = re.search(pattern, full_text_for_search, flags)
            
            if match:
                extracted_value = match.group(1) # キャプチャグループ1の内容を取得
                score = pattern_info['score']
                
                # クリーンアップ関数を適用し、値を正規化
                cleaned_val = clean_and_normalize(extracted_value, base_item_name)
                
                # 信頼度スコアを記録
                current_score = reliability_scores.get(base_item_name, 0)
                if score > current_score:
                    extracted_data[base_item_name] = cleaned_val
                    reliability_scores[base_item_name] = score
            
        # --- 開発工程フラグの判定 (本文全体を検索) ---
        for proc_name, keywords in PROCESS_KEYWORDS.items():
            flag_col = f'開発工程_{proc_name}'
            extracted_data[flag_col] = 'なし' 
            # 開発工程のキーワードが本文に含まれるかチェック
            if re.search('|'.join(keywords), full_mail_text, re.IGNORECASE):
                extracted_data[flag_col] = 'あり' 

        # --- N/Aの補完と最終データの構築 ---
        for col in MASTER_COLUMNS:
              if col not in extracted_data:
                  if col.startswith('開発工程_'):
                      extracted_data[col] = 'なし' 
                  else:
                      extracted_data[col] = 'N/A' 

        # 信頼度スコアの平均値を計算
        valid_scores = [s for s in reliability_scores.values() if s > 0]
        extracted_data['信頼度スコア'] = round(sum(valid_scores) / len(valid_scores) if valid_scores else 0, 1)
        extracted_data['本文(テキスト形式)'] = full_mail_text # 本文を最終出力に含める

        all_extracted_rows.append(extracted_data)
            
    # 最終的なDataFrameを構築し、列順を MASTER_COLUMNS に合わせる
    df_extracted = pd.DataFrame(all_extracted_rows)
    df_extracted = df_extracted.reindex(columns=MASTER_COLUMNS, fill_value='N/A')
    df_extracted = df_extracted.astype(str)
    
    return df_extracted