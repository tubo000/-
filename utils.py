# utils.py
#名前の保存と処理関数と設定ボタンの設定
#ソートの手伝い
import pandas as pd 
import os
import re
import unicodedata 

from config import CONFIG_FILE_PATH, TARGET_FOLDER_PATH

def load_config_csv():
    """name.csvからOutlookのアカウント名を読み込む"""
    try:
        df = pd.read_csv(CONFIG_FILE_PATH, encoding='utf-8-sig')
        df.columns = [col.strip().replace('\xa0', '').replace('\u3000', '') for col in df.columns] 
        if not df.empty and 'AccountName' in df.columns and len(df) > 0:
            account = df['AccountName'].iloc[0]
            # アカウント名の不要なスペース/制御文字を除去して返す
            return str(account).strip().replace('\xa0', '').replace('\u3000', ''), TARGET_FOLDER_PATH 
    except (pd.errors.EmptyDataError, FileNotFoundError):
        pass # ファイルがない、または空の場合は処理を継続
    except Exception as e:
        print(f"DEBUG: CSV設定ファイルの読み込み中にエラーが発生しました: {e}")
    return "", TARGET_FOLDER_PATH 

#name.csvに名前を保存するコード。名前の保存全てに使用される
def save_config_csv(account_name):
    """Outlookアカウント名をCSVに保存する"""
    try:
        config_df = pd.DataFrame({'AccountName': [account_name]})
        os.makedirs(os.path.dirname(CONFIG_FILE_PATH) or '.', exist_ok=True)
        config_df.to_csv(CONFIG_FILE_PATH, index=False, encoding='utf-8') # UTF-8で保存
        return True, f"アカウント名を上書き保存しました。"
    except Exception as e:
        return False, f" 設定ファイルの保存に失敗しました: {e}"
#正規表現の評価できるの形にする
def clean_and_normalize(value: str, item_name: str) -> str:
    """抽出した正規表現マッチ結果の値をクリーンアップし正規化する。"""
    if not value or value.strip() == '': return 'N/A'
    
    # 全角/半角スペース、制御文字を統一
    cleaned = value.strip().replace('\xa0', ' ').replace('\u3000', ' ')
    cleaned = re.sub(r'[\s　]+', ' ', cleaned).strip()
    
    if item_name == '氏名': 
        cleaned = re.sub(r'\s*\([^)]*\)', '', cleaned).strip() # (フリガナ)などを除去
        cleaned = re.sub(r'様\s*$', '', cleaned).strip() # 末尾の '様' を除去
    
    elif item_name == '年齢' or item_name == '単金': 
        # 抽出文字から数字、小数点、ハイフン、カンマ以外をすべて除去（safe_to_intで処理するために整形）
        # ※ここでは範囲指定のハイフンは考慮しない
        return re.sub(r'[^\d\.\-,]', '', cleaned).strip()
        
    elif item_name.startswith(('スキル_', '業務_')):
        # 区切り文字をカンマに統一し、不要なスペースを除去
        cleaned = re.sub(r'[・、/\\|,]', ',', cleaned)
        cleaned = re.sub(r'\s*,\s*', ',', cleaned).strip(',')
    
    return cleaned
def treeview_sort_column(tv, col, reverse):
    """Treeviewのカラムソート処理。数値カラムのソートを強化し、小数点以下を排除。"""
    l = [(tv.set(k, col), k) for k in tv.get_children('')]
    def try_convert(val):
        if pd.isna(val) or val is None or val == 'N/A': return ''
        if col in ['単金', '年齢']:
            val_str = str(val).replace(',', '').replace('万円', '').replace('歳', '').strip()
            try: 
                val_str = unicodedata.normalize('NFKC', val_str)
            except: pass
            try:
                # 整数に変換（小数点以下を切り捨て）
                return int(float(val_str))
            except ValueError: return val_str
        if col == '信頼度スコア':
             try: return float(val)
             except ValueError: return str(val)
        return str(val)
    l.sort(key=lambda t: try_convert(t[0]), reverse=reverse)
    for index, (val, k) in enumerate(l):
        tv.move(k, '', index)
    tv.heading(col, command=lambda c=col: treeview_sort_column(tv, c, not reverse))