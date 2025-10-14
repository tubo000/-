# outlook_api.py
#outlookに接続し、メールを取得

import pandas as pd 
import re
import win32com.client as win32 
from gui_config import MUST_INCLUDE_KEYWORDS,EXCLUDE_KEYWORDS

def get_outlook_folder(outlook_ns, account_name, folder_path):
    """Outlookフォルダオブジェクトを取得する。（アカウント名と階層パスを辿る）"""
    if outlook_ns.Stores.Count == 0: return None
    target_store = None
    clean_account_name = account_name.lower().strip()
    
    # 1. 指定アカウント名に一致するストア（アカウント）を検索
    if clean_account_name:
        try: target_store = next(st for st in outlook_ns.Stores if clean_account_name in st.DisplayName.lower())
        except StopIteration: return None 
    
    # 2. アカウント名指定がない場合や見つからない場合は、デフォルトのストア（通常最初のストア）を使用
    if target_store is None and outlook_ns.Stores.Count > 0: 
        target_store = outlook_ns.Stores.Item(1)
    if target_store is None: return None
    
    # 3. ルートフォルダからフォルダパスを辿る
    try:
        root_folder = target_store.GetRootFolder()
        current_folder = root_folder
        folders = folder_path.replace('/', '\\').split("\\") # パス区切り文字を統一
        for folder_name in folders:
            # フォルダ名を一つずつ辿っていく（大文字小文字を無視）
            current_folder = next((f for f in current_folder.Folders if f.Name.lower() == folder_name.lower()), None)
            if current_folder is None: return None # 途中で見つからなかった場合は失敗
        return current_folder
    except Exception as e:
        print(f"DEBUG: フォルダ検索中にエラーが発生: {e}")
        return None
#Outlookからメールを読み込み、メールの選択をして、フィルタリングされたデータを取得する。
def get_mail_data_from_outlook_in_memory(target_folder_path: str, account_name: str) -> pd.DataFrame:
    """Outlookからメールデータ（件名、本文、添付ファイル名など）を抽出する。"""
    data_records = []
    try:
        # Outlookアプリケーションを起動（または既存のインスタンスに接続）
        outlook_app = win32.Dispatch("Outlook.Application")
        outlook_ns = outlook_app.GetNamespace("MAPI")
        target_folder = get_outlook_folder(outlook_ns, account_name, target_folder_path)
        
        if target_folder is None: 
            error_msg = f"Outlookアカウント名またはフォルダパスが不正です。\n\n「Outlookアカウント (メールアドレス/表示名)」が正しいか、もう一度確かめてください。\n\n【アカウント】'{account_name if account_name else 'デフォルト'}'\n【フォルダ】'{target_folder_path}'"
            raise RuntimeError(error_msg)
            
        filtered_items = target_folder.Items
        total_items_in_folder = filtered_items.Count
        if total_items_in_folder == 0: return pd.DataFrame() # フォルダが空の場合は空のDataFrameを返す
        
        # フォルダ内の全アイテムをループし、フィルタリング
        for item in filtered_items:
            if item.Class == 43: # 43はメールアイテム (olMail) のクラスコード
                mail_item = item
                subject = getattr(mail_item, 'Subject', '')
                body = getattr(mail_item, 'Body', '')
                
                # キーワードフィルタリング（大文字小文字を無視）
                must_include = any(re.search(kw, subject, re.IGNORECASE) or re.search(kw, body, re.IGNORECASE) for kw in MUST_INCLUDE_KEYWORDS)
                is_excluded = any(re.search(kw, subject, re.IGNORECASE) or re.search(kw, body, re.IGNORECASE) for kw in EXCLUDE_KEYWORDS)
                
                # 除外キーワードが含まれる、または必須キーワードが含まれない場合はスキップ
                if is_excluded or not must_include: continue
                
                # 抽出対象のレコードとして追加
                record = {
                    'EntryID': str(getattr(mail_item, 'EntryID', 'UNKNOWN')),
                    '件名': subject, '本文(テキスト形式)': body, '__Source_Mail__': subject, 
                    'Attachments': [att.FileName for att in mail_item.Attachments] # 添付ファイル名も記録
                }
                data_records.append(record)
    except RuntimeError as re_e: raise re_e
    except Exception as e:
        # COM関連のエラーを検出して、より分かりやすいメッセージに変換
        if "ルールのパス" in str(e) or "クラスが登録されていません" in str(e):
            raise RuntimeError(f"Outlook操作エラー: Outlookが起動しているか、またはCOMアクセスが許可されているか確認してください。\n【詳細】{e}")
        raise RuntimeError(f"Outlook操作エラー: {e}")
    finally:
        # Outlookアプリケーションを安全に終了（接続解除）
        if 'outlook_app' in locals():
            try: outlook_app.Quit()
            except Exception: pass
    return pd.DataFrame(data_records)