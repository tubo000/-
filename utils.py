# utils.py
# 構成ファイル処理とTreeviewソートのユーティリティ

import pandas as pd 
import os
import re
import unicodedata 
import tkinter as tk 

# 📌 修正1: config.py から必要な変数をインポート
from config import CONFIG_FILE_PATH, TARGET_FOLDER_PATH, SCRIPT_DIR
# 📌 修正2: extraction_core.py で定義した process_tanaka をインポート (ソート処理で使う可能性があるため)
# from extraction_core import process_tanaka # NOTE: 相互インポートの回避のため、ここではインポートを省略し、ソートロジックを調整

def load_config_csv():
    """name.csvからOutlookのアカウント名を読み込む"""
    try:
        df = pd.read_csv(CONFIG_FILE_PATH, encoding='utf-8-sig')
        df.columns = [col.strip().replace('\xa0', '').replace('\u3000', '') for col in df.columns] 
        if not df.empty and 'AccountName' in df.columns and len(df) > 0:
            account = df['AccountName'].iloc[0]
            return str(account).strip().replace('\xa0', '').replace('\u3000', ''), TARGET_FOLDER_PATH 
    except (pd.errors.EmptyDataError, FileNotFoundError):
        pass
    except Exception as e:
        #(f"DEBUG: CSV設定ファイルの読み込み中にエラーが発生しました: {e}")
        pass
    return "", TARGET_FOLDER_PATH 

def save_config_csv(account_name):
    """Outlookアカウント名をCSVに保存する"""
    try:
        config_df = pd.DataFrame({'AccountName': [account_name]})
        os.makedirs(os.path.dirname(CONFIG_FILE_PATH) or SCRIPT_DIR, exist_ok=True)
        config_df.to_csv(CONFIG_FILE_PATH, index=False, encoding='utf-8')
        return True, f"アカウント名を上書き保存しました。"
    except Exception as e:
        return False, f" 設定ファイルの保存に失敗しました: {e}"

# 📌 修正3: utils.py 内の clean_and_normalize は削除し、
#          すべての抽出値のクリーンアップ/正規化は extraction_core.py の clean_and_normalize で一元管理します。
#          これにより、この関数定義は削除されます。

def treeview_sort_column(tv, col, reverse):
    """Treeviewのカラムソート処理。数値カラムのソートを強化し、小数点以下を排除。"""
    l = [(tv.set(k, col), k) for k in tv.get_children('')]
    
    def try_convert(val):
        if pd.isna(val) or val is None or val == 'N/A' or not str(val).strip(): return ''
        
        # 単金と年齢のソートロジックを調整
        if col in ['年齢']:
            val_str = str(val).replace(',', '').replace('歳', '').strip()
            try: 
                return int(float(unicodedata.normalize('NFKC', val_str)))
            except ValueError: return val_str
            
        elif col in ['期間_開始']:
            val_str = str(val).lower().strip()
            
            # 優先度: YYYYMM (古い順) < 即日 < Nヶ月 (短い順) < 要調整 < n/a
            if re.match(r'^\d{6}$', val_str): # YYYYMM 形式
                # 0をプレフィックスとして最も古い順にソートされるようにする
                return f"0{val_str}" 
            elif '即日' in val_str or 'asap' in val_str:
                return "1即日"
            elif month_match := re.search(r'(\d+)[ヶか]月', val_str):
                # Nヶ月。短い期間を優先するため、2の後にゼロ埋めしたNを付与
                return f"2{month_match.group(1).zfill(3)}ヶ月"
            elif '調整' in val_str or '相談' in val_str or '要' in val_str:
                return "3要調整"
            else:
                return "9n/a"
            
        elif col in ['単金']:
            val_str = str(val).strip()
            val_str = unicodedata.normalize('NFKC', val_str).replace(',', '').replace('万', '')
            
            # 範囲指定（例: 40~50）の場合は、最初の数字をソートキーとする
            range_match = re.search(r'(\d+)', val_str)
            if range_match:
                 try:
                    return int(range_match.group(1))
                 except ValueError: pass
                 
            try:
                # 単一値の場合、そのまま整数に変換
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