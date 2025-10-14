# email_processor.py

import pandas as pd
import os
import re
import sys
import win32com.client as win32
import csv 
# 📌 OpenPyXLは不要なため削除
from config import MASTER_COLUMNS
from extraction_core import extract_skills_data, clean_and_normalize

# =================================================================
# 【設定項目】
# =================================================================
TARGET_FOLDER_PATH = "受信トレイ" 
OUTPUT_FILENAME = 'extracted_skills_result.xlsx' # XLSX出力
# =================================================================


# ----------------------------------------------------
# 1. Outlook連携モジュール
# ----------------------------------------------------

def get_outlook_folder(outlook_ns, target_email, folder_path):
    """指定されたアカウントとパスに基づいてOutlookフォルダを取得する。"""
    
    if outlook_ns.Stores.Count == 0:
        print("DEBUG: Outlookにアカウント（ストア）が登録されていません。")
        return None
    
    target_store = None
    
    if target_email:
         try:
            target_store = next(st for st in outlook_ns.Stores if target_email.lower() in st.DisplayName.lower())
         except StopIteration:
            print(f"❌ エラー: アカウント名/メールアドレス '{target_email}' がOutlookに見つかりませんでした。")
            return None
    
    if target_store is None:
        try:
            target_store = outlook_ns.Stores.Item(1)
            print("DEBUG: アカウント指定なし。デフォルトストアを使用します。")
        except:
             print("DEBUG: デフォルトストア（インデックス1）の取得に失敗しました。")
             return None
        
    try:
        root_folder = target_store.GetRootFolder()
        current_folder = root_folder
        
        folders = re.split(r'[\\/]', folder_path)
        
        for folder_name in folders:
            current_folder = next((f for f in current_folder.Folders if f.Name.lower() == folder_name.lower()), None)
            
            if current_folder is None:
                print(f"DEBUG: フォルダ '{folder_name}' が '{folder_path}' パス内で見つかりません。")
                return None
        
        print(f"DEBUG: フォルダ '{folder_path}' をアカウント '{target_store.DisplayName}' から取得しました。")
        return current_folder
    
    except Exception as e:
        print(f"DEBUG: フォルダ検索中にエラーが発生: {e}")
        return None


def get_mail_data_from_outlook_in_memory(target_folder_path: str, target_email: str) -> pd.DataFrame:
    """Outlookからメールデータを取得し、DataFrameとして返す。"""
    data_records = []
    
    try:
        outlook_app = win32.Dispatch("Outlook.Application")
        outlook_ns = outlook_app.GetNamespace("MAPI")
        
        target_folder = get_outlook_folder(outlook_ns, target_email, target_folder_path)

        if target_folder is None:
            print(f"❌ 診断結果: フォルダが見つからないか、アカウントの認証に失敗しました。")
            return pd.DataFrame()
        
        filtered_items = target_folder.Items
        
        total_items_in_folder = filtered_items.Count
        print(f"DEBUG (A): フォルダ内のアイテム総数: {total_items_in_folder} 件")
        
        if total_items_in_folder == 0:
            print("✅ 処理完了。このフォルダにメールアイテムはありませんでした。")
            return pd.DataFrame()

        for item in filtered_items:
            
            if item.Class == 43:
                mail_item = item
                subject = getattr(mail_item, 'Subject', '')
                body = getattr(mail_item, 'Body', '')
                
                entry_id = getattr(mail_item, 'EntryID', f'OL_{len(data_records):04d}')
                to_address = getattr(mail_item, 'To', 'N/A')
                
                data_records.append({
                    'EntryID': entry_id,
                    '件名': subject,
                    '本文(テキスト形式)': body, 
                    '宛先メール': to_address,
                })
        
        print(f"✅ 成功: Outlookフォルダから {len(data_records)} 件のメールを抽出しました。")
        df = pd.DataFrame(data_records)
        return df.fillna('N/A').astype(str)

    except Exception as e:
        print(f"\n❌ Outlookアクセスエラーが発生しました。Outlookが起動しているか、win32comが正常に動作しているか確認してください。")
        print(f"詳細: {e}")
        return pd.DataFrame()


def run_email_extraction(target_email: str):
    """Outlookからデータを取得し、スキル抽出を行い、結果をXLSXに出力する。（Python完結の生URL方式）"""
    
    print("★★ Outlook メール抽出システム（本番環境模擬）実行 ★★")
    
    print("\n--- 1. Outlookからデータを読み込み ---")
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

        # EntryIDをURLに変換する処理 (EntryIDを左側に持ってくる役割を担う)
        df_output.insert(0, 'メールURL', df_output.apply(
            lambda row: f"outlook:{row['EntryID']}",
            axis=1
        ))
        
        # 📌 修正1: 最終出力列の順序を定義
        # 名前、件名、EntryIDを左側に配置し、残りの列を後に続ける
        
        # 最終出力に含める列 (このリストに沿って並び替えられる)
        fixed_leading_cols = ['メールURL', '件名', '名前']
        
        # その他の列を現在の順序で取得し、重複を避けて追加
        remaining_cols = [col for col in df_output.columns if col not in fixed_leading_cols]
        final_col_order = fixed_leading_cols + remaining_cols
        
        # DataFrameの列順を調整
        df_output = df_output[final_col_order]

        # 📌 修正2: 最終出力から不要な列を削除
        df_output = df_output.drop(columns=['EntryID', '宛先メール', '本文(テキスト形式)'], errors='ignore')

        # 1. pandasでベースデータ(.xlsx)を生成
        df_output.to_excel(output_file_abs_path, index=False)
        
        print(f"\n🎉 処理完了: 抽出結果を XLSX ファイル '{OUTPUT_FILENAME}' に出力しました。")
        print("💡 リンク機能はExcelに依存します。URL列をコピーし、Win+Rで貼り付けて開いてください。")
    
    
    except Exception as e:
        print(f"\n❌ XLSXファイル出力エラー: {e}")
        print("→ ファイルがロックされていないか確認してください。")