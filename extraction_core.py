# extraction_core.py
# 責務: メール本文からの情報抽出（正規表現適用）と、GUIで使用されるデータフィルタリングロジックの提供。

import pandas as pd
import re
import math # 単金処理のためにmathをインポート
import datetime             # 💡 追加: 日付処理用
from config import MASTER_COLUMNS, ITEM_PATTERNS, PROCESS_KEYWORDS
# 📌 修正1: configから高度な抽出ロジックの正規表現をインポート
from config import RE_AGE_PATTERNS, RE_TANAKA_KW_PATTERNS, RE_TANAKA_RAW_PATTERNS, KEYWORD_TANAKA

# 実働開始日の優先度ランキングを定義 (値が小さいほど高優先度)
DATE_RANKING = {
    'DATE_FULL': 1,      # yyyy年mm月形式 (例: 202511)
    'DURATION': 2,       # Nヶ月, 即日, asap
    'ADJUSTMENT': 3,     # 要調整, 要相談, 調整
    'OTHER': 4           # その他
}

def get_target_year(target_month):
    """4ヶ月フィルタリングを考慮して年を補完する"""
    now = datetime.datetime.now()
    current_year = now.year
    current_month = now.month
    
    target_year = current_year
    
    # ターゲット月が現在の月より過去の場合、来年と見なす 
    if target_month < current_month:
        target_year += 1
        
    return str(target_year)

def process_start_date(date_str):
    """
    実働開始日の後処理ロジック。
    - yyyy/mm/dd, yyyy-mm, mm/dd, yyyymm 形式を yyyymm 形式に正規化する。
    - 4ヶ月以上の未来日をフィルタリングする。
    - 最終的に 'YYYYMM', '即日', 'Nヶ月', '要調整', 'n/a' のいずれかのみを返す。
    """
    if pd.isna(date_str) or date_str is None: return 'nan', DATE_RANKING['OTHER']
    
    date_str_lower = date_str.lower().strip() 
    
    # --- 4ヶ月フィルタリング用のヘルパー関数を定義 ---
    def is_within_4_months(year_str, month_str):
        """ターゲット日(YYYYMM)が現在月を含め4ヶ月以内(0～4ヶ月)であるかをチェックする"""
        try:
            target_year = int(year_str)
            target_month = int(month_str)
        except ValueError:
            return False 

        now = datetime.datetime.now()
        current_year = now.year
        current_month = now.month

        current_total_months = current_year * 12 + current_month
        target_total_months = target_year * 12 + target_month
        
        month_difference = target_total_months - current_total_months
        
        # 0ヶ月（現在月）から4ヶ月先（+4）までを許可し、過去の日付（< 0）は除外
        if month_difference > 4 or month_difference < 0:
            return False 
        
        return True
        
    # --- 1. ○ヶ月の処理 ---
    month_match = re.search(r'([0-9]{1,3})[ヶか]月', date_str_lower)
    if month_match: return f"{month_match.group(1)}ヶ月", DATE_RANKING['DURATION'] 
    
    # --- 2. 即日/ASAPの処理 ---
    if '即日' in date_str_lower or 'asap' in date_str_lower or re.match(r'^即[～~]?$', date_str_lower): 
        return '即日', DATE_RANKING['DURATION']
    
    # --- 3. yyyy[区切り文字]mm... 形式の処理 (DATE_FULL) ---
    date_match = re.search(r'(\d{4})[\s\./\-年](\d{1,2})', date_str_lower)
    if date_match:
        year = date_match.group(1)
        month = date_match.group(2).zfill(2)
        
        if 1 <= int(month) <= 12:
            # 📌 修正1: 4ヶ月フィルタリングを有効化
            if is_within_4_months(year, month): 
                return f"{year}{month}", DATE_RANKING['DATE_FULL']
            return 'nan', DATE_RANKING['OTHER'] 
            
    # 3.3. YYYYMM形式の処理 (区切り文字なしの6桁数字)
    date_6digit_match = re.search(r'^(\d{6})$', date_str_lower)
    if date_6digit_match:
        year_str = date_6digit_match.group(1)[:4]
        month_str = date_6digit_match.group(1)[4:6]
        try:
             year = int(year_str)
             month = int(month_str)
             if 1 <= month <= 12 and 2000 <= year <= (datetime.datetime.now().year + 5):
                 # 📌 修正2: 4ヶ月フィルタリングを有効化
                 if is_within_4_months(year_str, month_str): 
                     return f"{year_str}{month_str}", DATE_RANKING['DATE_FULL']
                 return 'nan', DATE_RANKING['OTHER'] 
        except ValueError:
             pass

    # 3.4. Dayを含む年なしの日付形式の処理
    date_partial_match_day = re.search(r'(\d{1,2})[\s\./\-](\d{1,2})[\s～~-]?.*', date_str_lower)
    if date_partial_match_day:
        month = int(date_partial_match_day.group(1))
        
        if 1 <= month <= 12: 
            target_year_str = get_target_year(month) 
            target_month_str = str(month).zfill(2)
            
            # 📌 修正3: 4ヶ月フィルタリングを有効化
            if is_within_4_months(target_year_str, target_month_str): 
                return f"{target_year_str}{target_month_str}", DATE_RANKING['DATE_FULL']
            return 'nan', DATE_RANKING['OTHER']

    # --- 4. mm月（年なし）の処理 (例: 10月~) ---
    month_only_match = re.search(r'(\d{1,2})月', date_str_lower) 
    if month_only_match:
        target_month = int(month_only_match.group(1))
        
        if 1 <= target_month <= 12:
            target_year_str = get_target_year(target_month) 
            target_month_str = str(target_month).zfill(2)
                
            # 📌 修正4: 4ヶ月フィルタリングを有効化
            if is_within_4_months(target_year_str, target_month_str): 
                return f"{target_year_str}{target_month_str}", DATE_RANKING['DATE_FULL']
            return 'nan', DATE_RANKING['OTHER']

    # --- 5. 4桁の数字のみの処理（年のみの数字をn/aにする） ---
    if re.match(r'^\d{4}$', date_str_lower):
        return 'nan', DATE_RANKING['OTHER'] 

    # --- 6. その他の調整が必要なものなど ---
    if '調整' in date_str_lower or '相談' in date_str_lower or '要' in date_str_lower:
        return '要調整', DATE_RANKING['ADJUSTMENT']
        
    return 'nan', DATE_RANKING['OTHER']


def check_start_date_prefix(match, text_processed, col):
    """抽出された期間_開始候補に対して、前10文字にキーワードがあるかをチェックする。"""
    extracted_value = match.group(1).strip()
    start_index_of_match = match.start(0) 
    
    prefix_start_cut = max(0, start_index_of_match - 10) 
    extracted_prefix_10 = text_processed[prefix_start_cut:start_index_of_match]
    
    start_keywords = ['参画', '稼働', '稼動', '実働', '開始', '入場', '時期', 'start']
    
    is_keyword_present = any(kw in extracted_prefix_10.lower() for kw in start_keywords)
    
    if is_keyword_present:
        # NOTE: match.start(1) はキャプチャグループ1の開始インデックス
        return extracted_value, col, match.start(1)
    
    return None, None, None

def find_start_date_info(row):
    """期間_開始の情報を抽出するメイン関数"""
    text_cols = ['本文(テキスト形式)', '件名'] 
    
    # --- 期間_開始の正規表現 (コアロジック) ---
    START_KEYWORDS = (
        r'(?:参\s*画|稼\s*[働動]\s*日?|実\s*働|開\s*始|入\s*場|時\s*期|start)' 
    )
    DATE_PATTERN_FULL_AND_PARTIAL = (
        r'(?:\d{4}[\s\./\-年]\d{1,2}[\s\./\-月]\d{1,2}日?|\d{4}[\s\./\-年]\d{1,2}月)' 
        r'[\s～~-]*' 
    )
    ALL_START_DATE_OPTIONS = (
        DATE_PATTERN_FULL_AND_PARTIAL +           
        r'|'
        r'\d{6}'                                
        r'|'
        r'\d{1,2}[\s\./\-]\d{1,2}[\s～~-]*'      
        r'|'
        r'[1-9][0-9]{0,2}[ヶか]月'               
        r'|'
        r'即日[\s～~-]*'                          
        r'|'
        r'asap[\s～~-]*'                          
        r'|'
        r'即[～~]?'                              
        r'|'
        r'\d{1,2}月\b'                           
        r'|'
        r'\d{1,2}\b'                             
        r'|'
        r'(?:要調整|要相談|調整)'                 
    )
    RE_START_DATE_KEYWORDED = START_KEYWORDS + (
        r'[\s:：【\[（\(]?'                       
        r'(?:可\s*能\s*日?|\s*日|\s*時期|\s*予定)?' 
        r'[】\]）\)]?'                            
        r'[\s:：]*'                              
        r'('                                      
        r'(?:' + ALL_START_DATE_OPTIONS + r')'    
        r'(?:[\s/,、・]+'                         
        r'(?:' + ALL_START_DATE_OPTIONS + r'))*'  
        r')'                                      
    )

    RE_START_DATE_RAW = (
        r'(?:\D|^)' 
        r'('         
        + DATE_PATTERN_FULL_AND_PARTIAL +           
        r'|'
        r'\d{6}'                                
        r'|'
        r'\d{1,2}[\s\./\-]\d{1,2}[\s～~-]*'      
        r'|'
        r'[1-9][0-9]{0,2}[ヶか]月'       
        r'|'
        r'即日[\s～~-]*'                          
        r'|'
        r'asap[\s～~-]*'                          
        r'|'
        r'即[～~]?'                              
        r'|'
        r'\d{1,2}月\b'                           
        r'|'
        r'\d{1,2}\b'                             
        r'|'
        r'(?:要調整|要相談|調整)'         
        r')'
        r'(?:\D|$)' 
    )
    # ----------------------------------------------------------------------
    
    all_candidates = [] 
    keyword_patterns = [RE_START_DATE_KEYWORDED] 
    
    for col in text_cols: 
        text = row.get(col) 
        if pd.isna(text) or text is None: continue
        text_processed = str(text).replace('　', ' ')
        
        for regex in keyword_patterns: 
            matches = re.finditer(regex, text_processed, re.IGNORECASE)
            for match in matches:
                extracted_value = match.group(1).strip()
                index = match.start(1)
                # 複数の日付候補（例: 7月, 8月）が抽出された場合に分割
                sub_candidates = re.split(r'[\s/,、・]+', extracted_value) 
                
                for sub_value in sub_candidates:
                    if sub_value.strip():
                         all_candidates.append({
                            'value': sub_value.strip(), 
                            'col': col, 
                            'index': index, 
                            'source': 'KEYWORDED'
                        })
                    
        raw_patterns = [RE_START_DATE_RAW] 

        for regex in raw_patterns: 
            matches = re.finditer(regex, text_processed, re.IGNORECASE)
            
            for match in matches:
                # RAWパターンは前10文字のキーワードでフィルタリング
                value, src, idx = check_start_date_prefix(match, text_processed, col)
                
                if value:
                    all_candidates.append({
                        'value': value, 
                        'col': src, 
                        'index': idx, 
                        'source': 'RAW_FILTERED'
                    })

    if not all_candidates:
        return None, None, None

    # ランクとインデックスに基づいて最適な候補を決定
    best_match = None
    best_rank = DATE_RANKING['OTHER'] + 1 
    best_index = -1
    
    for candidate in all_candidates:
        # 正規化と4ヶ月フィルタリングを実行し、ランク付け
        processed_value, current_rank = process_start_date(candidate['value'])
        
        if processed_value == 'n/a' and current_rank == DATE_RANKING['OTHER']:
            continue # 無効な候補はスキップ
        
        current_index = candidate['index']
        
        if current_rank < best_rank:
            # より高優先度のランクが見つかった場合
            best_rank = current_rank
            best_match = candidate
            best_match['processed_value'] = processed_value
            best_index = current_index
        
        elif current_rank == best_rank:
            # 同じランクの場合、より後ろのインデックス（最新の情報）を優先
            if current_index > best_index: 
                best_match = candidate
                best_match['processed_value'] = processed_value
                best_index = current_index

    if best_match:
        return best_match['processed_value'], best_match['col'], best_match['index']
    
    return None, None, None

# --- 年齢/単金 フィルタリング・後処理ヘルパー関数を定義 ---

def check_age_and_prefix(match, text_processed, item_name):
    """【年齢フィルタリング】前20文字に「齢」または「名」があるかをチェック"""
    extracted_value = match.group(1).strip()
    start_index = match.start(1)
    
    age_start_index = match.start(1)
    prefix_start_cut = max(0, age_start_index - 20) 
    extracted_prefix_20 = text_processed[prefix_start_cut:age_start_index]
    
    if '齢' in extracted_prefix_20 or '名' in extracted_prefix_20:
        # スコア100を付けて返す（他の正規表現と区別するため）
        return extracted_value, 100, start_index
    
    return None, None, None

def check_tanaka_and_prefix(match, text_processed, item_name):
    """【単金フィルタリング】前10文字に「単」または「金額」があるか、IDを除外するかをチェック"""
    extracted_value = match.group(1).strip()
    start_index = match.start(1)
    
    # ID/URL除外ロジック (前15文字)
    check_end = match.start(0)
    check_start = max(0, check_end - 15) 
    prefix = text_processed[check_start:check_end]
    if re.search(r'p\?t=M|【ID】|\[ID\]|ID[\s:：]', prefix, re.IGNORECASE):
        return None, None, None 

    # キーワード近接チェック (前10文字)
    tanaka_start_index = match.start(1)
    prefix_start_cut = max(0, tanaka_start_index - 10) # 前10文字
    extracted_prefix_10 = text_processed[prefix_start_cut:tanaka_start_index]
    
    # フィルタリング条件: 「単」単独、または「金」と「額」のセット
    is_tanaka_or_kin_gaku = '単' in extracted_prefix_10 or ('金' in extracted_prefix_10 and '額' in extracted_prefix_10)
    
    if is_tanaka_or_kin_gaku:
        # スコア90を付けて返す（キーワード付き(100)より低く設定）
        return extracted_value, 90, start_index
    
    return None, None, None

def process_tanaka(tanaka_str: str) -> str:
    """単金の後処理ロジック: 範囲指定はそのまま、単一値は万単位に変換し、切り上げ"""
    if not tanaka_str or pd.isna(tanaka_str): return 'nan'
    tanaka_str = str(tanaka_str).lower().replace(' ', '').replace(',', '')
    
    # 1. 範囲指定の処理
    if ('~' in tanaka_str or '～' in tanaka_str or '-' in tanaka_str):
        range_str = tanaka_str.replace('万', '').replace('円', '').replace('~', '～').replace('-', '～')
        parts = re.split(r'～', range_str)
        if all(re.match(r'^\d+(\.\d+)?$', part) for part in parts) and len(parts) == 2:
             return range_str
        return 'nan'
    
    # 2. 単一値の処理
    man_value = None
    if '万' in tanaka_str:
        num_part = tanaka_str.replace('万', '')
        try: man_value = float(num_part)
        except ValueError: pass
    elif '円' in tanaka_str:
        num_part = tanaka_str.replace('円', '')
        if re.match(r'^\d+$', num_part): 
            num = int(num_part)
            man_value = num / 10000.0 if num >= 10000 else None
            if man_value is None: return str(num)
    elif re.match(r'^\d+$', tanaka_str): 
        num = int(tanaka_str)
        man_value = num / 10000.0 if num >= 10000 else None
        if man_value is None: return str(num)

    # 3. 繰り上げ処理 (万単位のみ)
    if man_value is not None:
        # 例: 70.5万 -> 71万 となるように切り上げ
        return str(int(math.ceil(man_value)))

    return 'nan'


def clean_and_normalize(value: str, item_name: str) -> str:
    """抽出結果をクリーンアップし、正規化する関数。（ノイズ除去を含む）"""
    if not value or not value.strip(): 
        return 'nan'
    
    cleaned = value.strip().replace('\xa0', ' ')
    cleaned = re.sub(r'[\s\u3000]+', ' ', cleaned).strip() 
    
    if item_name == '名前' or item_name == '氏名':
        cleaned = re.sub(r'[\(（\[【].*?[\)）\]】]', '', cleaned) 
        cleaned = re.sub(r'[・\_]', ' ', cleaned)
        cleaned = re.sub(r'[-]+$', '', cleaned).strip() 
        cleaned = re.sub(r'\s+', '', cleaned).strip()
    
    # 📌 修正2: 年齢の後処理は、カスタムフィルタリング関数 (check_age_and_prefix) での抽出後に行うため、
    #          ここではシンプルな整形のみにとどめる（主な処理は`process_tanaka`で行う）
    if item_name == '年齢':
        # 数字、ハイフン、チルダ以外を除去し、数字のみを返す（フィルタリングはextract_skills_data内で実施済み）
        cleaned = re.sub(r'[^\d\s\-\-～~]', '', cleaned).strip()
        # 範囲指定の数字を単一値とみなし、最初の数字だけを抽出する
        match = re.search(r'(\d+)', cleaned)
        return match.group(1) if match else 'N/A'
        
    if item_name == '単金':
       
        return value.strip() 

    if item_name in ['マネジメント経験人数']:
        return re.sub(r'[\D\.,]+', '', cleaned) 
        
    if item_name in ['スキルor言語', 'OS', 'データベース', 'フレームワーク/ライブラリ', '開発ツール']:
        cleaned = re.sub(r'^【\s*言\s*語\s*】|^【\s*DB\s*】|^【\s*OS\s*】', '', cleaned, flags=re.IGNORECASE)
        cleaned = cleaned.strip() 
        
        cleaned = re.sub(r'[・、/\\|,;]+', ',', cleaned) 
        cleaned = re.sub(r'\s*,\s*', ',', cleaned).strip(',') 
        
        if not cleaned:
            return 'N/A'
        
    return cleaned


def extract_skills_data(mail_data_df: pd.DataFrame) -> pd.DataFrame:
    """メールデータDataFrameを受け取り、抽出結果と信頼度スコアを返す。"""
    
    SOURCE_TEXT_COL = '本文(抽出元結合)'
    if SOURCE_TEXT_COL not in mail_data_df.columns:
        SOURCE_TEXT_COL = '本文(テキスト形式)'
        
    mail_data_df[SOURCE_TEXT_COL] = mail_data_df[SOURCE_TEXT_COL].astype(str).str.replace(r'[\r\n\t"]', ' ', regex=True)
    mail_data_df[SOURCE_TEXT_COL] = mail_data_df[SOURCE_TEXT_COL].astype(str).str.replace(r'\s+', ' ', regex=True)
    
    all_extracted_rows = []
    
    for index, row in mail_data_df.iterrows():
        mail_id = str(row.get('EntryID', f'Row_{index+1}'))
        full_text_for_search = str(row.get(SOURCE_TEXT_COL, ''))
        full_text_for_search = re.sub(r'\[FILE (ERROR|WARN):.*?\]', '', full_text_for_search)
        text_processed = full_text_for_search # フィルタリング用として保持
        
        mail_body_for_display = str(row.get('本文(テキスト形式)', 'N/A'))
        file_body_for_display = str(row.get('本文(ファイル含む)', 'N/A'))

        extracted_data = {'EntryID': mail_id, '件名': row.get('件名', 'N/A'), '宛先メール': row.get('宛先メール', 'N/A')} 
        reliability_scores = {} 
        
        # --- 2.1. 高度な年齢抽出 (最優先) ---
        age_extracted_value, age_score = 'nan', 0
        for regex in RE_AGE_PATTERNS:
            match = re.search(regex, text_processed, re.IGNORECASE)
            if match:
                value, score, _ = check_age_and_prefix(match, text_processed, '年齢')
                if value and score > age_score:
                    age_extracted_value = value
                    age_score = score
                    break # 最初に見つかった確度の高いもの採用
                    
        if age_extracted_value != 'nan':
            # clean_and_normalizeで単一値の数字のみを抽出
            extracted_data['年齢'] = clean_and_normalize(age_extracted_value, '年齢')
            reliability_scores['年齢'] = age_score
        
        # --- 2.2. 高度な単金抽出 (最優先) ---
        tanaka_extracted_value, tanaka_score = 'nan', 0
        
        # 1. キーワード付きパターン (RE_TANAKA_KW_PATTERNS) を優先順位通りにチェック
        for regex in RE_TANAKA_KW_PATTERNS:
            match = re.search(regex, text_processed, re.IGNORECASE)
            if match:
                tanaka_extracted_value = match.group(1).strip()
                tanaka_score = 100 # キーワード付きは高スコア
                break # 優先度の高いものが採用されたら終了
        
        # 2. キーワード付きが見つからなかった場合、RAWパターン (RE_TANAKA_RAW_PATTERNS) を優先順位通りにチェック (フィルタリングあり)
        if tanaka_extracted_value == 'nan':
            for regex in RE_TANAKA_RAW_PATTERNS:
                matches = re.finditer(regex, text_processed, re.IGNORECASE)
                for match in matches:
                    # フィルタリングチェックを適用
                    value, score, _ = check_tanaka_and_prefix(match, text_processed, '単金')
                    if value and score > tanaka_score:
                        tanaka_extracted_value = value
                        tanaka_score = score
                        break # 最初に見つかった確度の高いもの採用

        # 後処理の適用
        extracted_data['単金'] = process_tanaka(tanaka_extracted_value)
        if extracted_data['単金'] != 'nan':
            reliability_scores['単金'] = tanaka_score if tanaka_score > 0 else 100
            
        # --- 2.3. 期間_開始 の高度な抽出とフィルタリング ---
        start_date_value, _, _ = find_start_date_info(row)
        if start_date_value and start_date_value != 'nan':
             extracted_data['期間_開始'] = start_date_value
             reliability_scores['期間_開始'] = 100 
            
        # --- 2.4. その他の項目の抽出 ---
        for item_key, pattern_info in ITEM_PATTERNS.items():
            base_item_name = item_key.split('_')[0]
            
            # 年齢、単金、期間はカスタムロジックで処理済みのためスキップ
            if base_item_name in ['年齢', '単金', '期間']:
                continue
            
            pattern = pattern_info['pattern'] 
            flags = re.IGNORECASE
            if item_key == 'スキルor言語': flags |= re.DOTALL 

            match = re.search(pattern, full_text_for_search, flags)
            
            if match:
                extracted_value = match.group(1) if match.groups() else match.group(0)
                score = pattern_info['score']
                
                cleaned_val = clean_and_normalize(extracted_value, base_item_name)
                
                current_score = reliability_scores.get(base_item_name, 0)
                if score > current_score:
                    extracted_data[base_item_name] = cleaned_val
                    reliability_scores[base_item_name] = score
            
        # --- 3. 開発工程フラグの判定 (変更なし) ---
        for proc_name, keywords in PROCESS_KEYWORDS.items():
            flag_col = f'開発工程_{proc_name}'
            extracted_data[flag_col] = 'なし' 
            if re.search('|'.join(keywords), full_text_for_search, re.IGNORECASE):
                extracted_data[flag_col] = 'あり' 

        # --- 4. 最終データの構築と補完 (変更なし) ---
        final_row = {} 
        final_row.update(extracted_data) 

        final_row['本文(テキスト形式)'] = mail_body_for_display
        final_row['本文(ファイル含む)'] = file_body_for_display 
        
        attachments_list = row.get('Attachments', [])
        if isinstance(attachments_list, list):
            final_row['Attachments'] = ', '.join(attachments_list)
        else:
            final_row['Attachments'] = str(attachments_list)

        valid_scores = [s for s in reliability_scores.values() if s > 0]
        final_row['信頼度スコア'] = round(sum(valid_scores) / len(valid_scores) if valid_scores else 0, 1)

        all_extracted_rows.append(final_row)
            
    # 最終DataFrameの構築
    df_extracted = pd.DataFrame(all_extracted_rows)
    
    if '本文(抽出元結合)' in df_extracted.columns:
        df_extracted = df_extracted.drop(columns=['本文(抽出元結合)']) 
        
    all_cols_in_order = [
        'EntryID', '件名', '宛先メール', '本文(テキスト形式)', '本文(ファイル含む)', 
        'Attachments', '信頼度スコア'
    ] + [col for col in MASTER_COLUMNS if col not in ['EntryID', '件名', '宛先メール', '本文(テキスト形式)', '本文(ファイル含む)', 'Attachments', '信頼度スコア', '本文(抽出元結合)']]
    
    df_extracted = df_extracted.reindex(columns=[c for c in all_cols_in_order if c in df_extracted.columns], fill_value='N/A')
    df_extracted = df_extracted.astype(str)
    
    return df_extracted