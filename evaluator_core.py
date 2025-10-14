# evaluator_core.py
# 責務: 試験用CSVの読み書き、抽出結果とマスターデータの比較評価、およびGUIユーティリティ

import pandas as pd
import re
import os
import unicodedata  # ソート用の文字正規化に必要
from tkinter import ttk     # Treeviewソートに必要 (ttk)
import tkinter as tk        # Tkinterの基本機能

from config import EVALUATION_TARGETS, MASTER_ANSWERS_PATH, OUTPUT_EVAL_PATH 


# ----------------------------------------------------
# 評価用ユーティリティ関数 (GUIフィルタリングとソートで使用)
# ----------------------------------------------------

def safe_to_int(value):
    """単金や年齢の文字列を安全に整数に変換するヘルパー関数（GUIフィルタリング用）"""
    if pd.isna(value) or value is None: return None
    value_str = str(value).strip()
    if not value_str: return None 
    try:
        # 文字列のクリーンアップと正規化
        cleaned_str = re.sub(r'[\s　\xa0\u3000]+', '', value_str) 
        normalized_str = unicodedata.normalize('NFKC', cleaned_str)
        
        # 不要な文字を除去 (万円, 歳, カンマなどを除去)
        cleaned_str = normalized_str.replace(',', '').replace('万円', '').replace('歳', '').strip()
        cleaned_str = re.sub(r'[^\d\.]', '', cleaned_str) 
        
        if not cleaned_str: return None
        
        # 浮動小数点数として解釈し、整数に変換（小数点以下を切り捨て）
        return int(float(cleaned_str))
        
    except ValueError:
        return None
    except Exception:
        return None 

def treeview_sort_column(tv, col, reverse):
    """Treeviewのカラムソート処理。数値カラムのソートを強化する。"""
    # Treeviewからデータをリストとして取得 (タプル形式: [(値, item_id), ...])
    l = [(tv.set(k, col), k) for k in tv.get_children('')]
    
    def try_convert(val):
        """ソートキーとして使うために値を数値または文字列に変換"""
        if pd.isna(val) or val is None or val == 'N/A': return ''
        
        if col in ['単金', '年齢']:
            # 数値カラム: 文字列から数値のみを抽出し、整数としてソート
            val_str = str(val).replace(',', '').replace('万円', '').replace('歳', '').strip()
            try:
                val_str = unicodedata.normalize('NFKC', val_str)
            except: pass
            try:
                return int(float(val_str))
            except ValueError: return val_str
            
        if col == '信頼度スコア':
             # 信頼度スコアは浮動小数点数としてソート
             try: return float(val)
             except ValueError: return str(val)
             
        return str(val)
        
    # リストをソート (ソートキーに関数 try_convert を適用)
    l.sort(key=lambda t: try_convert(t[0]), reverse=reverse)
    
    # Treeviewの並び順を更新
    for index, (val, k) in enumerate(l):
        tv.move(k, '', index)
        
    # ヘッダーのコマンドを再設定し、ソート順を反転させる
    tv.heading(col, command=lambda c=col: treeview_sort_column(tv, c, not reverse))

# ----------------------------------------------------
# 評価コアロジック
# ----------------------------------------------------

def get_question_data_from_csv(file_path: str) -> pd.DataFrame:
    """外部CSVを読み込み、抽出対象のDataFrameとして返す。"""
    if not os.path.exists(file_path):
        print(f"❌ エラー: 問題CSVファイル '{file_path}' が見つかりません。")
        return pd.DataFrame()
    
    try:
        df = pd.read_csv(file_path, encoding='utf-8-sig', sep=None, engine='python', dtype={'EntryID': str})
        print(f"✅ 問題CSVから {len(df)} 件のデータを読み込みました。")
        return df
    except Exception as e:
        print(f"❌ 問題CSVの読み込み中にエラーが発生しました。ファイル形式を確認してください。エラー: {e}")
        return pd.DataFrame()


def clean_name_for_comparison(name_str):
    """評価比較用に氏名をクリーンアップし、連結する"""
    name_str = str(name_str).strip()
    name_str = re.sub(r'[\(（\[【].*?[\)）\]】]', '', name_str) 
    name_str = re.sub(r'[・\_]', ' ', name_str)
    name_str = re.sub(r'[-]+$', '', name_str).strip() 
    name_str = re.sub(r'\s+', '', name_str).strip().lower()
    return name_str

def run_triple_csv_validation(df_extracted: pd.DataFrame, master_path: str, output_path: str):
    """
    抽出結果とマスターデータ（正解）を比較し、項目の正誤判定（✅/❌）と総合精度を計算し、
    結果を新しいタブ区切りCSVとして出力する。
    """
    
    print("\n--- 3. 評価と検証（リザルトCSV生成）---")
    
    # --- 1. マスターデータの読み込み ---
    try:
        df_master = pd.read_csv(master_path, 
                                encoding='utf-8-sig', 
                                dtype={'EntryID': str},
                                sep='\t').set_index('EntryID')
        print(f"✅ 正解マスターから {len(df_master)} 件のデータを読み込みました。")
    except Exception as e:
        print(f"❌ 処理停止: 正解マスターの読み込みに失敗しました。ファイル形式を確認してください。エラー: {e}")
        return

    # --- 2. データ結合と前処理 ---
    EVAL_COLS = [c for c in EVALUATION_TARGETS if c in df_master.columns] 
    merged_df = pd.merge(df_extracted.reset_index(drop=True), df_master.reset_index(), on='EntryID', how='inner', suffixes=('_E', '_M'))
    
    if merged_df.empty:
        print("⚠️ 抽出結果とマスターファイルで一致するメールID（EntryID）がありません。評価できませんでした。")
        return

    # --- 3. 評価ロジックの実行 ---
    total_checks = 0
    total_correct = 0
    
    for index, row in merged_df.iterrows():
        for col in EVAL_COLS:
            col_E = f'{col}_E'
            col_M = f'{col}_M'
            
            if col_E not in merged_df.columns or col_M not in merged_df.columns: continue
            
            total_checks += 1
            
            if col == '名前':
                extracted_val = clean_name_for_comparison(row[col_E])
                master_val = clean_name_for_comparison(row[col_M])
            else:
                extracted_val = re.sub(r'[\s\t\r\n\u200b\u3000,\-歳万]+', '', str(row[col_E]).strip().lower())
                master_val = re.sub(r'[\s\t\r\n\u200b\u3000,\-歳万]+', '', str(row[col_M]).strip().lower())
            
            if master_val == 'n/a' or not master_val:
                total_checks -= 1 
                continue
            
            is_match = (extracted_val == master_val)
            
            # 判定結果を記録
            merged_df.loc[index, f'{col}_判定'] = '✅' if is_match else '❌'
            
            if is_match: total_correct += 1
    
    # --- 4. 精度計算と最終出力 ---
    accuracy = (total_correct / total_checks) * 100 if total_checks > 0 else 0
    print(f"\n🎉 評価完了: 総合精度 = {accuracy:.2f}% ({total_correct} / {total_checks} 項目)")
    
    # 万円表示の補助列を生成
    def convert_yen_to_man(yen_str):
        """円単位の文字列を万円単位の文字列に変換する"""
        try:
            return str(int(yen_str) // 10000)
        except:
            return yen_str
            
    if '単金_E' in merged_df.columns:
        merged_df['単金_E_万'] = merged_df['単金_E'].apply(convert_yen_to_man)
    if '単金_M' in merged_df.columns:
        merged_df['単金_M_万'] = merged_df['単金_M'].apply(convert_yen_to_man)

    # 最終出力列の順序を決定
    output_cols = ['EntryID'] 
    for c in EVALUATION_TARGETS:
        if f'{c}_E' in merged_df.columns: output_cols.append(f'{c}_E')
        if f'{c}_M' in merged_df.columns: output_cols.append(f'{c}_M')
        if c == '単金': # 単金の場合は万円表示の列を追加
            if '単金_E_万' in merged_df.columns: output_cols.append('単金_E_万')
            if '単金_M_万' in merged_df.columns: output_cols.append('単金_M_万')
        if f'{c}_判定' in merged_df.columns: output_cols.append(f'{c}_判定')
        
    # タブ区切りCSVとして出力
    merged_df[output_cols].to_csv(output_path, index=False, encoding='utf-8-sig', sep='\t')
    print(f"✨ 評価結果をタブ区切りCSV '{output_path}' に出力しました。")