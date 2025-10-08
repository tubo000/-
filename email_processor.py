# email_processor.py

import pandas as pd
import os
import re
import sys
import win32com.client as win32
import csv 
from config import MASTER_COLUMNS
from extraction_core import extract_skills_data, clean_and_normalize

# =================================================================
# 【設定項目】
# =================================================================
# 📌 ユーザー設定項目
# DEFAULT_ACCOUNT_NAMEは、実行時に引数として渡すため、ここでは削除
TARGET_FOLDER_PATH = "受信トレイ" 
OUTPUT_FILENAME = 'extracted_skills_result.csv' 
# =================================================================


# ----------------------------------------------------
# 1. Outlook連携モジュール
# ----------------------------------------------------

# 📌 修正1: target_email 引数を追加
def get_outlook_folder(outlook_ns, target_email, folder_path):
    """指定されたアカウントとパスに基づいてOutlookフォルダを取得する。"""
    
    if outlook_ns.Stores.Count == 0:
        print("DEBUG: Outlookにアカウント（ストア）が登録されていません。")
        return None
    
    target_store = None
    
    # ★★★ アカウント指定時の処理を強化 ★★★
    if target_email:
         try:
            # 📌 修正2: アカウントの DisplayName (通常はメールアドレス) を使って検索
            # target_emailを含むストア（アカウント）を検索
            target_store = next(st for st in outlook_ns.Stores if target_email.lower() in st.DisplayName.lower())
         except StopIteration:
            print(f"❌ エラー: アカウント名/メールアドレス '{target_email}' がOutlookに見つかりませんでした。")
            return None # 見つからない場合は処理を中止
    
    # アカウントが指定されていない場合は、Stores.Item(1) (デフォルト) を使用
    if target_store is None:
        try:
            target_store = outlook_ns.Stores.Item(1)
            print("DEBUG: アカウント指定なし。デフォルトストアを使用します。")
        except:
             print("DEBUG: デフォルトストア（インデックス1）の取得に失敗しました。")
             return None
        
    # フォルダの取得ロジック
    try:
        # GetRootFolder()は、指定されたストア（アカウント）の最上位のフォルダを返します
        root_folder = target_store.GetRootFolder()
        current_folder = root_folder
        
        folders = re.split(r'[\\/]', folder_path)
        
        for folder_name in folders:
            # フォルダ名検索ロジック
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
        
        # ターゲットフォルダの取得 (ターゲットメールを渡す)
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

        # ループ開始
        for item in filtered_items:
            
            if item.Class == 43: # 43 = olMailItem (MailItem のみ対象)
                mail_item = item
                subject = getattr(mail_item, 'Subject', '')
                body = getattr(mail_item, 'Body', '')
                
                entry_id = getattr(mail_item, 'EntryID', f'OL_{len(data_records):04d}')
                
                data_records.append({
                    'EntryID': entry_id,
                    '件名': subject,
                    '本文(テキスト形式)': body, 
                    '宛先メール': getattr(mail_item, 'To', 'N/A'),
                })
        
        print(f"✅ 成功: Outlookフォルダから {len(data_records)} 件のメールを抽出しました。")
        df = pd.DataFrame(data_records)
        return df.fillna('N/A').astype(str)

    except Exception as e:
        print(f"\n❌ Outlookアクセスエラーが発生しました。Outlookが起動しているか、win32comが正常に動作しているか確認してください。")
        print(f"詳細: {e}")
        return pd.DataFrame()


# ----------------------------------------------------
# 2. メイン実行関数（本番相当）
# ----------------------------------------------------

def run_email_extraction(target_email: str):
    """Outlookからデータを取得し、スキル抽出を行い、結果をCSVに出力する。"""
    
    print("★★ Outlook メール抽出システム（本番環境模擬）実行 ★★")
    
    # 1. Outlookからメールデータを取得 (ターゲットメールを渡す)
    print("\n--- 1. Outlookからデータを読み込み ---")
    df_mail_data = get_mail_data_from_outlook_in_memory(TARGET_FOLDER_PATH, target_email)
    
    if df_mail_data.empty:
        print("処理対象のメールがありませんでした。処理を終了します。")
        return

    # 2. 抽出コアロジックを実行
    print("\n--- 2. スキル抽出実行 ---")
    df_extracted = extract_skills_data(df_mail_data)
    
    # 3. 結果を単一のCSVとして出力
    try:
        df_output = df_extracted.copy()
        
        df_output.to_csv(
            OUTPUT_FILENAME, 
            index=False, 
            encoding='utf-8-sig', 
            sep='\t', 
            quoting=csv.QUOTE_ALL
        )
        print(f"\n🎉 処理完了: 抽出結果をタブ区切りCSV '{OUTPUT_FILENAME}' に出力しました。")
    
    except Exception as e:
        print(f"\n❌ 結果ファイル出力エラー: '{OUTPUT_FILENAME}' の書き込みに失敗しました。詳細: {e}")