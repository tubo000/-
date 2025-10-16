# outlook_api.py
# 責務: Outlookデスクトップアプリに接続し、メールを取得。

import pandas as pd 
import re
import os 
import win32com.client as win32 
import unicodedata # テキストの正規化（NFKC）に必須
import uuid # ファイル名の安全な生成に使用
import time # 📌 追加: ファイル削除前の待機に使用

# config.py から必須/除外キーワードをインポート
from config import MUST_INCLUDE_KEYWORDS, EXCLUDE_KEYWORDS, SCRIPT_DIR
from file_processor import get_attachment_text # 添付ファイルテキスト抽出関数をインポート


# ... (get_outlook_folder 関数は変更なし)
def get_outlook_folder(outlook_ns, account_name, folder_path):
    """Outlookフォルダオブジェクトを取得する。（アカウント名と階層パスを辿る）"""
    # ... (コードは変更なし)
    if outlook_ns.Stores.Count == 0: return None
    target_store = None
    clean_account_name = account_name.lower().strip()
    
    if clean_account_name:
        try: target_store = next(st for st in outlook_ns.Stores if clean_account_name in st.DisplayName.lower())
        except StopIteration: return None 
    
    if target_store is None and outlook_ns.Stores.Count > 0: 
        target_store = outlook_ns.Stores.Item(1)
    if target_store is None: return None
    
    try:
        root_folder = target_store.GetRootFolder()
        current_folder = root_folder
        # パスの区切り文字を正規化
        folders = folder_path.replace('/', '\\').split('\\')
        
        for folder_name in folders:
            if not folder_name: continue
            
            current_folder = next(
                (f for f in current_folder.Folders if f.Name.lower() == folder_name.lower()),
                None
            )
            if current_folder is None: return None
        
        return current_folder
    except Exception as e:
        raise RuntimeError(f"指定されたフォルダパスの検索に失敗しました: {folder_path}。詳細: {e}")


def get_mail_data_from_outlook_in_memory(target_folder_path: str, account_name: str) -> pd.DataFrame:
    """Outlookからメールデータ（件名、本文、添付ファイル名、添付ファイルテキスト）を抽出する。"""
    data_records = []
    total_attachments = 0
    non_supported_count = 0
    
    temp_dir = os.path.join(SCRIPT_DIR, "temp_attachments_safe")
    os.makedirs(temp_dir, exist_ok=True)
    
    try:
        outlook_app = win32.Dispatch("Outlook.Application")
        outlook_ns = outlook_app.GetNamespace("MAPI")
        target_folder = get_outlook_folder(outlook_ns, account_name, target_folder_path)
        
        if target_folder is None:
            raise RuntimeError(f"指定されたアカウント名 '{account_name}' またはフォルダパス '{target_folder_path}' が見つかりませんでした。")

        items = target_folder.Items
        
        for item in items:
            if item.Class == 43: # olMailItem
                mail_item = item
                subject = getattr(mail_item, 'Subject', '')
                body = getattr(mail_item, 'Body', '')
                
                # HTML本文の取得ロジック (HTMLタグを削除し、テキストとして使用)
                if not body or not body.strip():
                    html_body = getattr(mail_item, 'HTMLBody', '')
                    if html_body:
                        body = re.sub('<[^>]*>', ' ', html_body) 
                        body = re.sub(r'\s+', ' ', body).strip()
                if not body: body = 'N/A' 
                
                attachments_text = ""
                attachment_names = []
                
                # 添付ファイルの処理ブロック
                if mail_item.Attachments.Count > 0:
                    for attachment in mail_item.Attachments:
                        total_attachments += 1 
                        
                        # 安全なファイルパスの生成
                        safe_filename = re.sub(r'[\\/:*?"<>|]', '_', attachment.FileName)
                        # 拡張子がない場合の暫定対応 (必須)
                        if not os.path.splitext(safe_filename)[1]:
                            safe_filename = f"file_{uuid.uuid4().hex}_{safe_filename}.dat"
                            
                        temp_file_path = os.path.join(temp_dir, safe_filename)
                        
                        try:
                            attachment.SaveAsFile(temp_file_path)
                            
                            attachment_text = get_attachment_text(temp_file_path, attachment.FileName)
                            
                            # file_processor側でクリーンアップされているため、ここでは状態の確認のみ
                            if attachment_text.startswith("[WARN:") or attachment_text.startswith("[ERROR:"):
                                non_supported_count += 1
                                attachments_text += f"\n{attachment_text}"
                            else:
                                # 抽出テキストを改行区切りで結合
                                attachments_text += "\n" + attachment_text
                                
                            attachment_names.append(attachment.FileName)
                        except Exception as file_e:
                            attachments_text += f"\n[FILE ERROR: {attachment.FileName}: {file_e}]"
                        finally:
                            # 📌 修正3: 一時ファイルのクリーンアップの堅牢性を高める
                            if os.path.exists(temp_file_path):
                                for i in range(3): # 3回まで再試行
                                    try:
                                        os.remove(temp_file_path)
                                        break
                                    except OSError as e:
                                        # WinError 32 (ファイルがロックされている) の場合、少し待って再試行
                                        if e.errno == 32:
                                            print(f"⚠️ ファイルロック解除待機中... {temp_file_path}")
                                            time.sleep(0.1 * (i + 1)) # 0.1, 0.2, 0.3秒待機
                                        else:
                                            # それ以外のOSErrorは無視せず、ログに記録
                                            print(f"❌ ファイル削除失敗: {temp_file_path}: {e}")
                                            break
                                
                # 添付ファイルテキストの最終確認
                attachments_text = attachments_text.strip()
                if not attachments_text: attachments_text = 'N/A' 
                
                # 抽出に使用する結合全文をローカル変数として作成 (フィルタリング用)
                full_search_text = body + " " + attachments_text + " " + subject
                
                # キーワードフィルタリング
                must_include = any(re.search(kw, full_search_text, re.IGNORECASE) for kw in MUST_INCLUDE_KEYWORDS)
                is_excluded = any(re.search(kw, full_search_text, re.IGNORECASE) for kw in EXCLUDE_KEYWORDS)
                
                if is_excluded or not must_include: continue
                
                # 抽出対象のレコードとして追加
                record = {
                    'EntryID': str(getattr(mail_item, 'EntryID', 'UNKNOWN')),
                    '件名': subject,
                    '本文(テキスト形式)': body, # メール本文のみ
                    '本文(ファイル含む)': attachments_text, # ファイルから抜き取ったテキストのみ
                    '本文(抽出元結合)': full_search_text, # 抽出に使用する全文 (一時的に格納)
                    'Attachments': attachment_names 
                }
                data_records.append(record)
    except RuntimeError as re_e: raise re_e
    except Exception as e:
        # 致命的なエラーログ
        if "ルールのパス" in str(e) or "クラスが登録されていません" in str(e):
            raise RuntimeError(f"Outlook操作エラー: Outlookが起動しているか、またはCOMアクセスが許可されているか確認してください。\n【詳細】{e}")
        # 📌 WinError 32 がここで再送出される可能性があるため、エラーコードを明示的にチェック
        if isinstance(e, win32.client.pywintypes.com_error) and e.hresult == -2147352567: # COM Error
             raise RuntimeError(f"Outlook操作エラー: {e}")
        elif isinstance(e, OSError) and e.errno == 32:
             raise RuntimeError(f"Outlook操作エラー: [WinError 32] ファイルロックエラー: {e}")
        raise RuntimeError(f"Outlook操作エラー: {e}")
    finally:
        # 一時ディレクトリのクリーンアップ
        if os.path.exists(temp_dir) and not os.listdir(temp_dir):
            try: os.rmdir(temp_dir)
            except OSError: pass
            
    # 最終的な処理結果の表示
    print("\n--- 添付ファイル処理結果 ---")
    print(f"✅ 処理された添付ファイルの総数: {total_attachments} 個")
    print(f"⚠️ 非対応または処理中にエラーが発生したファイルの数: {non_supported_count} 個")
    print("---------------------------\n")
            
    df = pd.DataFrame(data_records)
    return df.fillna('N/A').astype(str)