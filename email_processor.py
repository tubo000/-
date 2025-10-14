# email_processor.py

import pandas as pd
import os
import re
import sys
import win32com.client as win32
import csv 
# 📌 OpenPyXLは使わず、win32comとpandasのみを使用
from config import MASTER_COLUMNS
from extraction_core import extract_skills_data, clean_and_normalize
from outlook_api import get_mail_data_from_outlook_in_memory # Outlook接続関数をインポート


# =================================================================
# 【設定項目】
# =================================================================
TARGET_FOLDER_PATH = "受信トレイ" 
OUTPUT_FILENAME = 'extracted_skills_result.xlsx' # XLSX出力
# =================================================================


# ----------------------------------------------------
# メイン実行関数（本番相当）
# ----------------------------------------------------

def run_email_extraction(target_email: str):
    """Outlookからデータを取得し、スキル抽出を行い、結果をXLSXに出力する。（Python完結の生URL方式）"""
    
    print("★★ Outlook メール抽出システム（本番環境模擬）実行 ★★")
    
    print("\n--- 1. Outlookからデータを読み込み ---")
    # outlook_api.py の関数を呼び出す
    df_mail_data = get_mail_data_from_outlook_in_memory(TARGET_FOLDER_PATH, target_email)
    
    if df_mail_data.empty:
        print("処理対象のメールがありませんでした。処理を終了します。")
        return

    print("\n--- 2. スキル抽出実行 ---")
    df_extracted = extract_skills_data(df_mail_data)
    
    # 3. 結果を単一のXLSXとして出力
    try:
        df_output = df_extracted.copy()
        output_file_abs_path = os.path.abspath(OUTPUT_FILENAME)

        # ★ リンク機能のための URL 列を生成 ★
        # EntryIDを Outlook URI スキーム形式の文字列として格納 (Python完結の生URL方式)
        df_output.insert(0, 'メールURL', df_output.apply(
            lambda row: f"outlook:{row['EntryID']}",
            axis=1
        ))

        # 📌 最終出力列の整理（EntryID, 宛先メール, 本文は削除）
        df_output = df_output.drop(columns=['EntryID', '宛先メール', '本文(テキスト形式)'], errors='ignore')

        # 列順序を調整し、メールURL、件名、名前を左側に固定
        fixed_leading_cols = ['メールURL', '件名', '名前']
        remaining_cols = [col for col in df_output.columns if col not in fixed_leading_cols]
        final_col_order = fixed_leading_cols + remaining_cols
        df_output = df_output[final_col_order]

        # pandasでベースデータ(.xlsx)を生成
        df_output.to_excel(output_file_abs_path, index=False)
        
        print(f"\n🎉 処理完了: 抽出結果を XLSX ファイル '{OUTPUT_FILENAME}' に出力しました。")
        print("💡 リンク機能はExcelに依存します。URL列をコピーし、Win+Rで貼り付けて開いてください。")
        
    
    except Exception as e:
        print(f"\n❌ XLSXファイル出力エラー: {e}")