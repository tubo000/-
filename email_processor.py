# email_processor.py (本文内容の取得を修正)

import pandas as pd
import win32com.client as win32
import pythoncom
import os
import datetime
import re
from datetime import timedelta
import sys
import uuid # ファイル保存用
import traceback

# 外部定数と関数の依存関係を想定
try:
    from config import MUST_INCLUDE_KEYWORDS, EXCLUDE_KEYWORDS, SCRIPT_DIR
    # 実際には outlook_api.py などに実装が必要な関数を仮定義
    def get_outlook_folder(outlook_ns, account_name, folder_path):
        try:
            return outlook_ns.Folders[account_name].Folders[folder_path]
        except Exception:
            return None
    
    # 📌 修正1: 添付ファイルの中身（テキスト）を取得する関数 (file_processor.py に実装されているはず)
    # ここでは、外部関数 get_attachment_text の存在と動作を想定します。
    # 実際の処理では、添付ファイルを一時ファイルとして保存し、そこからテキストを抽出する必要があります。
    try:
        from file_processor import get_attachment_text
    except ImportError:
        def get_attachment_text(*args, **kwargs): return "ATTACHMENT_CONTENT_ERROR_SKIP" # 外部モジュールがない場合のダミー
    
except ImportError:
    MUST_INCLUDE_KEYWORDS = [r'スキルシート']
    EXCLUDE_KEYWORDS = []
    SCRIPT_DIR = os.getcwd()
    def get_outlook_folder(*args, **kwargs): return None
    def get_attachment_text(*args, **kwargs): return "ATTACHMENT_CONTENT_ERROR_SKIP" 
    
OUTPUT_FILENAME = 'extracted_skills_result.xlsx' 
PROCESSED_CATEGORY_NAME = "スキルシート処理済" 

# ----------------------------------------------------------------------
# 💡 共通機能: メールアイテムの処理済みマーク (維持)
# ----------------------------------------------------------------------

def mark_email_as_processed(mail_item):
    """
    指定されたメールアイテムに「処理済み」カテゴリを設定する。
    """
    if mail_item.Class == 43: # olMailItem
        try:
            current_categories = getattr(mail_item, 'Categories', '')
            if PROCESSED_CATEGORY_NAME not in current_categories:
                if current_categories:
                    mail_item.Categories = f"{current_categories},{PROCESSED_CATEGORY_NAME}"
                else:
                    mail_item.Categories = PROCESSED_CATEGORY_NAME
                mail_item.Save()
        except Exception as e:
            pass 
        return True
    return False

# ----------------------------------------------------------------------
# 💡 メイン抽出関数: Outlookからメールを取得
# ----------------------------------------------------------------------

def get_mail_data_from_outlook_in_memory(target_folder_path: str, account_name: str, read_mode: str = "all", days_ago: int = None) -> pd.DataFrame:
    """
    Outlookからメールデータを抽出する。read_modeに基づいてフィルタリングを行う。
    """
    data_records = []
    temp_dir = os.path.join(SCRIPT_DIR, "temp_attachments_safe")
    os.makedirs(temp_dir, exist_ok=True)
    
    try:
        pythoncom.CoInitialize()
        outlook_app = win32.Dispatch("Outlook.Application")
        outlook_ns = outlook_app.GetNamespace("MAPI")
        target_folder = get_outlook_folder(outlook_ns, account_name, target_folder_path)
        
        if target_folder is None:
            raise RuntimeError(f"指定されたフォルダパス '{target_folder_path}' が見つかりませんでした。")

        items = target_folder.Items
        
        # 日付フィルタリングクエリの構築 (維持)
        filter_query = []
        if read_mode == "days" and days_ago is not None:
            start_date = (datetime.datetime.now() - timedelta(days=days_ago)).strftime('%m/%d/%Y %H:%M %p')
            filter_query.append(f"[ReceivedTime] >= '{start_date}'")

        if filter_query:
            query_string = " AND ".join(filter_query)
            items = items.Restrict(query_string)
            
        # --- メインループ ---
        for item in items:
            if item.Class == 43: # olMailItem (メールアイテムのみを処理)
                
                try: 
                    # 未処理チェック
                    if read_mode == "unprocessed":
                        current_categories = getattr(item, 'Categories', '')
                        if PROCESSED_CATEGORY_NAME in current_categories:
                            continue 
                    
                    mail_item = item
                    subject = getattr(mail_item, 'Subject', '')
                    body = getattr(mail_item, 'Body', '')
                    received_time = getattr(mail_item, 'ReceivedTime', datetime.datetime.now())
                    
                    if received_time is not None and received_time.tzinfo is not None:
                        received_time = received_time.replace(tzinfo=None)
                    elif received_time is None:
                        received_time = datetime.datetime.now().replace(tzinfo=None)
                    
                    attachments_text = ""
                    attachment_names = []
                    
                    # 📌 修正2: 添付ファイルの内容を実際に取得するロジック
                    if hasattr(mail_item, 'Attachments') and mail_item.Attachments.Count > 0:
                        for attachment in mail_item.Attachments:
                            attachment_names.append(attachment.FileName)
                            
                            # 安全なファイルパスの生成
                            safe_filename = re.sub(r'[\\/:*?"<>|]', '_', attachment.FileName)
                            temp_file_path = os.path.join(temp_dir, f"{uuid.uuid4().hex}_{safe_filename}")
                            
                            try:
                                # 添付ファイルを一時保存
                                attachment.SaveAsFile(temp_file_path)
                                # テキストを抽出
                                extracted_content = get_attachment_text(temp_file_path, attachment.FileName)
                                attachments_text += f"\n--- FILE: {attachment.FileName} ---\n{extracted_content}\n"
                            except Exception as file_ex:
                                attachments_text += f"\n--- ERROR reading {attachment.FileName}: {file_ex} ---\n"
                            finally:
                                # 一時ファイルを削除
                                if os.path.exists(temp_file_path):
                                    os.remove(temp_file_path)
                        
                        attachments_text = attachments_text.strip()
                    
                    full_search_text = subject + " " + body + " " + attachments_text
                    
                    # キーワードフィルタリング (MUST/EXCLUDE)
                    must_include = any(re.search(kw, full_search_text, re.IGNORECASE) for kw in MUST_INCLUDE_KEYWORDS)
                    is_excluded = any(re.search(kw, full_search_text, re.IGNORECASE) for kw in EXCLUDE_KEYWORDS)
                    
                    if is_excluded or not must_include: 
                        continue 

                    # レコードの準備
                    record = {
                        'EntryID': getattr(mail_item, 'EntryID', 'UNKNOWN'),
                        '件名': subject,
                        '受信日時': received_time, 
                        '本文(テキスト形式)': body, 
                        # 📌 修正3: 抽出したファイル本文を格納
                        '本文(ファイル含む)': attachments_text, 
                        'Attachments': ", ".join(attachment_names),
                    }
                    data_records.append(record)
                    
                    # 抽出が成功したら、メールを「処理済み」としてマーク
                    mark_email_as_processed(mail_item) 

                except Exception as item_ex:
                    print(f"警告: メールアイテムの処理中にエラーが発生しました (EntryID: {getattr(item, 'EntryID', '不明')}). スキップします。エラー: {item_ex}")
                    continue 

    except Exception as e:
        # 📌 修正1: traceback モジュールがインポートされているため、エラー処理を修正
        raise RuntimeError(f"Outlook操作エラー: {e}\n詳細: {traceback.format_exc()}")
    finally:
        # 一時ディレクトリのクリーンアップ
        if os.path.exists(temp_dir) and not os.listdir(temp_dir):
            try: os.rmdir(temp_dir)
            except OSError: pass
        pythoncom.CoUninitialize() 
            
    df = pd.DataFrame(data_records)
    str_cols = [col for col in df.columns if col != '受信日時']
    df[str_cols] = df[str_cols].fillna('N/A').astype(str)
    return df

# ----------------------------------------------------------------------
# 💡 外部公開関数
# ----------------------------------------------------------------------

def run_email_extraction(target_email: str, read_mode: str = "all", days_ago: int = None):
    pass

def delete_old_emails_core(target_email: str, folder_path: str, days_ago: int) -> int:
    pass