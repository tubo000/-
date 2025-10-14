# email_processor.py
# 責務: Outlookデスクトップアプリに接続し、指定アカウント/フォルダのメールを取得。
#       その後、抽出コアロジックを実行し、結果をXLSX形式で出力する。

import pandas as pd
import os
import re
import sys
import win32com.client as win32 # 必須: Windows COMインターフェース (Outlook接続) 用
import csv 
from config import MASTER_COLUMNS # 抽出結果の最終列定義に使用
from extraction_core import extract_skills_data, clean_and_normalize # 抽出ロジックをインポート

# =================================================================
# 【設定項目】
# =================================================================
TARGET_FOLDER_PATH = "受信トレイ"           # Outlookからメールを取得するデフォルトフォルダ
OUTPUT_FILENAME = 'extracted_skills_result.xlsx' # 抽出結果の出力ファイル名 (XLSX形式)
# =================================================================


# ----------------------------------------------------
# 1. Outlook連携モジュール (アカウント/フォルダの検索)
# ----------------------------------------------------

def get_outlook_folder(outlook_ns, target_email, folder_path):
    """
    指定されたメールアドレス（アカウント）とフォルダパスに基づいて、
    Outlookフォルダオブジェクトを MAPI Namespace から取得する。
    """
    
    # Outlookにアカウントが登録されているかチェック
    if outlook_ns.Stores.Count == 0:
        print("DEBUG: Outlookにアカウント（ストア）が登録されていません。")
        return None
    
    target_store = None
    
    # --- アカウント指定がある場合の検索 ---
    if target_email:
         try:
             # 指定されたメールアドレスを含む DisplayName のストア（アカウント）を検索
            target_store = next(st for st in outlook_ns.Stores if target_email.lower() in st.DisplayName.lower())
         except StopIteration:
            print(f"❌ エラー: アカウント名/メールアドレス '{target_email}' がOutlookに見つかりませんでした。")
            return None
    
    # --- アカウント指定がない、または見つからなかった場合のデフォルト処理 ---
    if target_store is None:
        try:
            # 最初のストア（通常はデフォルトアカウント）を使用
            target_store = outlook_ns.Stores.Item(1)
            print("DEBUG: アカウント指定なし。デフォルトストアを使用します。")
        except:
             print("DEBUG: デフォルトストア（インデックス1）の取得に失敗しました。")
             return None
        
    # --- フォルダパスの検索（ルートから階層を辿る） ---
    try:
        root_folder = target_store.GetRootFolder() # 選択されたアカウントの最上位フォルダを取得
        current_folder = root_folder
        
        # パスを区切り文字 ('\' or '/') で分割
        folders = re.split(r'[\\/]', folder_path)
        
        for folder_name in folders:
             # 現在のフォルダの子フォルダを名前で検索（大文字小文字を無視）
            current_folder = next((f for f in current_folder.Folders if f.Name.lower() == folder_name.lower()), None)
            
            if current_folder is None:
                print(f"DEBUG: フォルダ '{folder_name}' が '{folder_path}' パス内で見つかりません。")
                return None # 途中のフォルダが見つからなければ終了
        
        print(f"DEBUG: フォルダ '{folder_path}' をアカウント '{target_store.DisplayName}' から取得しました。")
        return current_folder
    
    except Exception as e:
        print(f"DEBUG: フォルダ検索中にエラーが発生: {e}")
        return None


def get_mail_data_from_outlook_in_memory(target_folder_path: str, target_email: str) -> pd.DataFrame:
    """Outlookからメールデータを取得し、DataFrameとして返す。"""
    data_records = []
    
    try:
        # Outlookアプリケーションへの接続
        outlook_app = win32.Dispatch("Outlook.Application")
        outlook_ns = outlook_app.GetNamespace("MAPI")
        
        # ターゲットフォルダのオブジェクトを取得
        target_folder = get_outlook_folder(outlook_ns, target_email, target_folder_path)

        if target_folder is None:
            print(f"❌ 診断結果: フォルダが見つからないか、アカウントの認証に失敗しました。")
            return pd.DataFrame()
        
        filtered_items = target_folder.Items # フォルダ内の全アイテム
        
        total_items_in_folder = filtered_items.Count
        print(f"DEBUG (A): フォルダ内のアイテム総数: {total_items_in_folder} 件")
        
        if total_items_in_folder == 0:
            print("✅ 処理完了。このフォルダにメールアイテムはありませんでした。")
            return pd.DataFrame()

        # フォルダ内のアイテムをループ
        for item in filtered_items:
            
            if item.Class == 43: # 43は olMailItem（メールアイテム）を示す
                mail_item = item
                subject = getattr(mail_item, 'Subject', '')
                body = getattr(mail_item, 'Body', '')
                
                # EntryID（メール固有のID）を取得
                entry_id = getattr(mail_item, 'EntryID', f'OL_{len(data_records):04d}')
                to_address = getattr(mail_item, 'To', 'N/A')
                
                # 抽出の元となるデータをレコードに追加
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
    """
    Outlookからデータを取得し、スキル抽出を行い、結果をXLSXに出力する。
    （Python完結の生URL方式: リンク機能はOSのWin+Rに依存）
    """
    
    print("★★ Outlook メール抽出システム（本番環境模擬）実行 ★★")
    
    print("\n--- 1. Outlookからデータを読み込み ---")
    df_mail_data = get_mail_data_from_outlook_in_memory(TARGET_FOLDER_PATH, target_email)
    
    if df_mail_data.empty:
        print("処理対象のメールがありませんでした。処理を終了します。")
        return

    print("\n--- 2. スキル抽出実行 ---")
    # 抽出コアモジュール（extraction_core.py）を呼び出し、抽出を実行
    df_extracted = extract_skills_data(df_mail_data)
    
    # 3. 結果を単一のXLSXとして出力 (Entry IDをURL文字列として出力)
    try:
        df_output = df_extracted.copy()
        output_file_abs_path = os.path.abspath(OUTPUT_FILENAME)

        # ★ リンク機能のための URL 列を生成 ★
        # EntryIDを Outlook URI スキーム形式の文字列として格納 (例: outlook:000...)
        df_output.insert(0, 'メールURL', df_output.apply(
            lambda row: f"outlook:{row['EntryID']}",
            axis=1
        ))

        # 📌 最終出力列から不要な列を削除 (EntryID, 宛先メール, 本文は整理)
        df_output = df_output.drop(columns=['EntryID', '宛先メール', '本文(テキスト形式)'], errors='ignore')

        # 最終的な列順を決定し、DataFrameを調整
        fixed_leading_cols = ['メールURL', '件名', '名前']
        remaining_cols = [col for col in df_output.columns if col not in fixed_leading_cols]
        final_col_order = fixed_leading_cols + remaining_cols
        df_output = df_output[final_col_order]


        # pandasでベースデータ(.xlsx)を生成
        df_output.to_excel(output_file_abs_path, index=False)
        
        print(f"\n🎉 処理完了: 抽出結果を XLSX ファイル '{OUTPUT_FILENAME}' に出力しました。")
        print("💡 リンク機能はExcelに依存します。URL列をコピーし、Win+Rで貼り付けて開いてください。")
        
        # 4. ファイルを自動で開く
        os.startfile(output_file_abs_path)
    
    except Exception as e:
        print(f"\n❌ XLSXファイル出力エラー: {e}")
        print("→ ファイルがロックされていないか確認してください。")