# evaluator_core.py
# 責務: 試験用CSVの読み書き、抽出結果とマスターデータの比較評価、および評価結果の出力を行う。

import pandas as pd
import re
import os
# configから評価対象の項目リスト（EVALUATION_TARGETS）をインポート
from config import EVALUATION_TARGETS 


# ----------------------------------------------------
# 評価用ユーティリティ関数
# ----------------------------------------------------

def get_question_data_from_csv(file_path: str) -> pd.DataFrame:
    """
    外部の試験用CSVファイル（問題データ）を読み込み、DataFrameとして返す。
    CSV/TSVどちらでも読み込めるよう、sep=None (自動判別) を使用。
    """
    if not os.path.exists(file_path):
        print(f"❌ エラー: 問題CSVファイル '{file_path}' が見つかりません。")
        return pd.DataFrame()
    
    try:
        # UTF-8 BOM付き ('utf-8-sig') で読み込み、EntryIDを文字列として保持
        df = pd.read_csv(file_path, encoding='utf-8-sig', sep=None, engine='python', dtype={'EntryID': str})
        print(f"✅ 問題CSVから {len(df)} 件のデータを読み込みました。")
        return df
    except Exception as e:
        print(f"❌ 問題CSVの読み込み中にエラーが発生しました。ファイル形式を確認してください。エラー: {e}")
        return pd.DataFrame()


def clean_name_for_comparison(name_str):
    """
    評価比較のために、氏名文字列からノイズ（括弧、記号、スペース）をすべて除去し、連結する。
    抽出結果とマスターデータの両方に適用される。
    """
    name_str = str(name_str).strip()
    # 1. 括弧（全角/半角/角括弧/波括弧）とその中の内容を削除
    name_str = re.sub(r'[\(（\[【].*?[\)）\]】]', '', name_str) 
    # 2. 名前の区切りに使われるノイズ文字（・、_）をスペースに変換
    name_str = re.sub(r'[・\_]', ' ', name_str)
    # 3. 末尾に連続するハイフンを削除
    name_str = re.sub(r'[-]+$', '', name_str).strip() 
    # 4. 氏名内の全てのスペースを削除し、小文字化して連結 (例: David Lee -> davidlee)
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
        # マスターはダミーデータ生成側でタブ区切りで出力されるため、sep='\t'を指定
        df_master = pd.read_csv(master_path, 
                                encoding='utf-8-sig', 
                                dtype={'EntryID': str},
                                sep='\t').set_index('EntryID')
        print(f"✅ 正解マスターから {len(df_master)} 件のデータを読み込みました。")
    except Exception as e:
        print(f"❌ 処理停止: 正解マスターの読み込みに失敗しました。ファイル形式を確認してください。エラー: {e}")
        return

    # --- 2. データ結合と前処理 ---
    # 評価対象の列をマスターデータから取得
    EVAL_COLS = [c for c in EVALUATION_TARGETS if c in df_master.columns] 
    
    # EntryIDをキーに、抽出結果（_E）とマスター（_M）を結合
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
                # 氏名: 専用のクリーンアップ関数を適用
                extracted_val = clean_name_for_comparison(row[col_E])
                master_val = clean_name_for_comparison(row[col_M])
            else:
                # その他: スペース、タブ、改行、単位（歳/万）などを除去して比較
                extracted_val = re.sub(r'[\s\t\r\n\u200b\u3000,\-歳万]+', '', str(row[col_E]).strip().lower())
                master_val = re.sub(r'[\s\t\r\n\u200b\u3000,\-歳万]+', '', str(row[col_M]).strip().lower())
            
            # マスターがN/Aまたは空の場合は比較対象から除外
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