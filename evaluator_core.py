# evaluator_core.py

import pandas as pd
import re
import os
from config import EVALUATION_TARGETS


# ----------------------------------------------------
# 評価用ユーティリティ関数
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
    
    print("\n--- 3. 評価と検証（リザルトCSV生成）---")
    
    try:
        df_master = pd.read_csv(master_path, 
                                encoding='utf-8-sig', 
                                dtype={'EntryID': str},
                                sep='\t').set_index('EntryID')
        print(f"✅ 正解マスターから {len(df_master)} 件のデータを読み込みました。")
    except Exception as e:
        print(f"❌ 処理停止: 正解マスターの読み込みに失敗しました。ファイル形式を確認してください。エラー: {e}")
        return

    EVAL_COLS = [c for c in EVALUATION_TARGETS if c in df_master.columns] 
    merged_df = pd.merge(df_extracted.reset_index(drop=True), df_master.reset_index(), on='EntryID', how='inner', suffixes=('_E', '_M'))
    
    if merged_df.empty:
        print("⚠️ 抽出結果とマスターファイルで一致するメールID（EntryID）がありません。評価できませんでした。")
        return

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
            
            merged_df.loc[index, f'{col}_判定'] = '✅' if is_match else '❌'
            
            if is_match: total_correct += 1
    
    accuracy = (total_correct / total_checks) * 100 if total_checks > 0 else 0
    print(f"\n🎉 評価完了: 総合精度 = {accuracy:.2f}% ({total_correct} / {total_checks} 項目)")
    
    def convert_yen_to_man(yen_str):
        try:
            return str(int(yen_str) // 10000)
        except:
            return yen_str
            
    if '単金_E' in merged_df.columns:
        merged_df['単金_E_万'] = merged_df['単金_E'].apply(convert_yen_to_man)
    if '単金_M' in merged_df.columns:
        merged_df['単金_M_万'] = merged_df['単金_M'].apply(convert_yen_to_man)

    output_cols = ['EntryID'] 
    for c in EVALUATION_TARGETS:
        if f'{c}_E' in merged_df.columns: output_cols.append(f'{c}_E')
        if f'{c}_M' in merged_df.columns: output_cols.append(f'{c}_M')
        if c == '単金': 
            if '単金_E_万' in merged_df.columns: output_cols.append('単金_E_万')
            if '単金_M_万' in merged_df.columns: output_cols.append('単金_M_万')
        if f'{c}_判定' in merged_df.columns: output_cols.append(f'{c}_判定')
        
    merged_df[output_cols].to_csv(output_path, index=False, encoding='utf-8-sig', sep='\t')
    print(f"✨ 評価結果をタブ区切りCSV '{output_path}' に出力しました。")