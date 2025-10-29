# email_processor.py (最終安定版 - Restrictエラー対策適用版)

import pandas as pd
import win32com.client as win32
import pythoncom
import os
import datetime
import re
from datetime import timedelta
import sys
import uuid
import traceback
from typing import Dict, Any, List
import sqlite3
# ----------------------------------------------------------------------
# イニシャルを検出する正規表現を追加
# ----------------------------------------------------------------------
# 大文字2連続 (IR, KK) または 大文字,大文字 (K, K) または 名前(IR)
INITIALS_REGEX = r'(\b[A-Z]{2}\b|\b[A-Z]\s*.\s*[A-Z]\b|名前\([A-Z]{2}\))'
# 外部定数と関数の依存関係を想定 (維持)
try:
    from config import MUST_INCLUDE_KEYWORDS, EXCLUDE_KEYWORDS, SCRIPT_DIR, OUTPUT_CSV_FILE as OUTPUT_FILENAME
    def get_outlook_folder(outlook_ns, account_name, folder_path):
        """Outlookフォルダオブジェクトを取得する（実装は outlook_api.py にあるものと仮定）"""
        try:
            return outlook_ns.Folders[account_name].Folders[folder_path]
        except Exception:
            return None
    try:
        from file_processor import get_attachment_text
    except ImportError:
        def get_attachment_text(*args, **kwargs): return "ATTACHMENT_CONTENT_FILE_IO_FAILED"

except ImportError:
    MUST_INCLUDE_KEYWORDS = [r'スキルシート']
    EXCLUDE_KEYWORDS = [r'案\s*件\s*名',r'案\s*件\s*番\s*号',r'案\s*件:',r'案\s*件：',r'【案\s*件】',r'概\s*要',r'必\s*須']
    SCRIPT_DIR = os.getcwd()
    def get_outlook_folder(*args, **kwargs): return None
    def get_attachment_text(*args, **kwargs): return "ATTACHMENT_CONTENT_FILE_IO_FAILED"

# 保存先を .db ファイルに変更
DATABASE_NAME = 'extraction_cache.db'
PROCESSED_CATEGORY_NAME = "スキルシート処理済"

# ----------------------------------------------------------------------
# 💡 ヘルパー関数: 過去の本文データ復元 (sqlite3版)
# ----------------------------------------------------------------------

def _load_previous_attachment_content() -> Dict[str, str]:
    """
    (高速化) sqlite3 データベースから EntryID と 本文(ファイル含む) を読み込み、
    本文復元用の辞書を返す。
    """
    db_path = os.path.join(os.path.abspath(SCRIPT_DIR), DATABASE_NAME)

    if os.path.exists(db_path):
        try:
            conn = sqlite3.connect(db_path)
            df_prev = pd.read_sql_query("SELECT \"EntryID\", \"本文(ファイル含む)\" FROM emails", conn)
            conn.close()
            df_prev.set_index('EntryID', inplace=True)
            return df_prev['本文(ファイル含む)'].dropna().to_dict()

        except Exception as e:
            print(f"警告: データベースからの本文復元に失敗しました。エラー: {e}")
            return {}
    return {}

# ----------------------------------------------------------------------
# 💡 共通機能: メールアイテムの処理済みマーク (維持)
# ----------------------------------------------------------------------
def mark_email_as_processed(mail_item):
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

# email_processor.py の remove_processed_category 関数のみを差し替え
# (L125 付近)

# ----------------------------------------------------------------------
# 💡 処理済みカテゴリの解除 (Restrictエラー対策 + 降順ソート対応)
# ----------------------------------------------------------------------
def remove_processed_category(target_email: str, folder_path: str, days_ago: int = None) -> int:
    reset_count = 0
    start_date_dt = None
    if days_ago is not None:
        start_date_dt = (datetime.datetime.now() - timedelta(days=days_ago))

    try:
        pythoncom.CoInitialize()
        try:
            outlook = win32.GetActiveObject("Outlook.Application")
        except:
            outlook = win32.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        folder = get_outlook_folder(namespace, target_email, folder_path)
        if folder is None:
            raise RuntimeError(f"指定されたフォルダパス '{folder_path}' が見つかりませんでした。")

        items = folder.Items
        filter_query_list = []
        
        # ----------------------------------------------------
        # 📌 修正: Restrict は日付のみで行う
        # (カテゴリ LIKE 検索は Restrict から除外)
        # ----------------------------------------------------
        # category_filter_query = f"[Categories] LIKE '%{PROCESSED_CATEGORY_NAME}%'"
        # filter_query_list.append(category_filter_query)

        if days_ago is not None:
            start_date_str = start_date_dt.strftime('%Y/%m/%d')
            filter_query_list.append(f"[ReceivedTime] < '{start_date_str}'")

        query_string = " AND ".join(filter_query_list)
        
        try:
            if query_string: # 日付指定がある場合のみ Restrict
                items_to_reset = items.Restrict(query_string)
            else: # 日付指定がない場合は全件
                items_to_reset = items
        except Exception as restrict_error:
            # 💥 日付絞り込みも失敗する環境の場合
            print(f"警告: カテゴリ解除のRestrict(日付)に失敗しました: {restrict_error}")
            items_to_reset = items

        # ----------------------------------------------------
        # 📌 修正: 取得したアイテムを受信日時の降順 (新しい順) に並び替え
        # ----------------------------------------------------
        items_to_reset.Sort("[ReceivedTime]", True)

        item = items_to_reset.GetFirst()
        while item:
            if item.Class == 43:
                try:
                    # ----------------------------------------------------
                    # 📌 修正: Python側でカテゴリをチェック (必須)
                    # ----------------------------------------------------
                    current_categories = getattr(item, 'Categories', '')
                    if PROCESSED_CATEGORY_NAME in current_categories:
                        
                        # (日付 Restrict 失敗時のフォールバックチェック)
                        is_target = True
                        if days_ago is not None:
                            received_time = getattr(item, 'ReceivedTime', datetime.datetime.now())
                            if received_time.tzinfo is not None:
                                received_time = received_time.replace(tzinfo=None)
                            if received_time >= start_date_dt:
                                is_target = False

                        if is_target:
                            categories_list = [c.strip() for c in current_categories.split(',') if c.strip() != PROCESSED_CATEGORY_NAME]
                            item.Categories = ", ".join(categories_list)
                            item.Save()
                            reset_count += 1
                except Exception as e:
                    print(f"警告: カテゴリ解除中にアイテムエラー: {e}")
            item = items_to_reset.GetNext()
        return reset_count
    except Exception as e:
        raise RuntimeError(f"カテゴリマーク解除中にエラーが発生しました。詳細: {e}")
    finally:
        pythoncom.CoUninitialize()

# email_processor.py の has_unprocessed_mail 関数のみを差し替え

def has_unprocessed_mail(folder_path: str, target_email: str, days_to_check: int =  14) -> int:
    """
    指定されたフォルダに、処理済みカテゴリがなく、
    (1) 添付ファイルがある、または
    (2) 本文/件名にキーワードやイニシャルが含まれる
    メールの件数をカウントする。
    
    📌 修正 (ハイブリッドアプローチ):
    カテゴリでの Restrict が失敗するため、
    1. サーバー側で日付絞り込み (例: 過去90日) を行い (高速)
    2. Python側でカテゴリをチェックする (安定的)
    """
    unprocessed_count = 0
    if not folder_path or not target_email: return 0

    try:
        pythoncom.CoInitialize()

        try:
            outlook = win32.GetActiveObject("Outlook.Application")
        except:
            outlook = win32.Dispatch("Outlook.Application")

        namespace = outlook.GetNamespace("MAPI")
        folder = get_outlook_folder(namespace, target_email, folder_path)

        if folder:
            items = folder.Items

            # ----------------------------------------------------
            # 🚀 ハイブリッドアプローチ: 1. サーバー側で日付絞り込み (高速)
            # ----------------------------------------------------
            try:
                # 起動時のチェックは、直近90日間に限定する
                cutoff_date_dt = (datetime.datetime.now() - timedelta(days=days_to_check))
                cutoff_date_str = cutoff_date_dt.strftime('%Y/%m/%d')
                date_filter = f"[ReceivedTime] >= '{cutoff_date_str}'"
                
                items = items.Restrict(date_filter)
                print(f"DEBUG: has_unprocessed_mail: 過去{days_to_check}日間に絞り込み成功。")
                
            except Exception as restrict_error:
                # 💥 日付絞り込みも失敗する環境の場合
                print(f"警告: has_unprocessed_mailの日付絞り込みに失敗。全件スキャン (低速): {restrict_error}")
                items = folder.Items # 失敗時は全件スキャン (安定だが遅い)
            
            # ----------------------------------------------------
            # 2. Python側でカテゴリ絞り込み (安定的)
            # ----------------------------------------------------
            item = items.GetFirst()
            while item:
                try:
                    if item.Class == 43:
                        # Python側でのフィルタリング (必須)
                        categories = str(getattr(item, 'Categories', ''))
                        if PROCESSED_CATEGORY_NAME not in categories:

                            has_attachments = hasattr(item, 'Attachments') and item.Attachments.Count > 0

                            if has_attachments:
                                unprocessed_count += 1
                                item = items.GetNext() 
                                continue

                            subject = str(getattr(item, 'Subject', ''))
                            body = str(getattr(item, 'Body', ''))
                            full_search_text = subject + " " + body

                            must_include = any(re.search(kw, full_search_text, re.IGNORECASE) for kw in MUST_INCLUDE_KEYWORDS)
                            has_initials = re.search(INITIALS_REGEX, full_search_text)

                            if must_include or has_initials:
                                unprocessed_count += 1

                except Exception as e:
                    print(f"警告: アイテムスキャン中にCOMエラー: {e}")

                item = items.GetNext() 

    except Exception as e:
        print(f"警告: 未処理メールチェック中にCOMエラー発生: {e}")
        unprocessed_count = 0

    finally:
        pythoncom.CoUninitialize()

    return unprocessed_count


# ----------------------------------------------------------------------
# 💡 メイン抽出関数: Outlookからメールを取得 (高精度ロジック + 高速Restrict)
# ----------------------------------------------------------------------

def get_mail_data_from_outlook_in_memory(target_folder_path: str, account_name: str, read_mode: str = "all", days_ago: int = None) -> pd.DataFrame:
    """
    Outlookからメールデータを抽出する。read_modeに基づいてフィルタリングを行う。
    """
    data_records = []
    temp_dir = os.path.join(SCRIPT_DIR, "temp_attachments_safe")
    os.makedirs(temp_dir, exist_ok=True)

    previous_attachment_content = _load_previous_attachment_content()
    
    start_date_dt = None # 
    if days_ago is not None:
        start_date_dt = (datetime.datetime.now() - timedelta(days=days_ago))

    try:
        pythoncom.CoInitialize()

        try:
            outlook_app = win32.GetActiveObject("Outlook.Application")
        except:
            outlook_app = win32.Dispatch("Outlook.Application")

        outlook_ns = outlook_app.GetNamespace("MAPI")
        target_folder = get_outlook_folder(outlook_ns, account_name, target_folder_path)

        if target_folder is None:
            raise RuntimeError(f"指定されたフォルダパス '{target_folder_path}' が見つかりませんでした。")

        items = target_folder.Items

        filter_query_list = []

        if days_ago is not None:
            # ----------------------------------------------------
            # 📌 修正 (エラー対策): 日付形式を 'YYYY/MM/DD' に変更 (時刻を削除)
            # ----------------------------------------------------
            start_date_str = start_date_dt.strftime('%Y/%m/%d')
            filter_query_list.append(f"[ReceivedTime] >= '{start_date_str}'")


        if filter_query_list:
            query_string = " AND ".join(filter_query_list)
            try:
                items = items.Restrict(query_string)
            except Exception as restrict_error:
                print(f"警告: Outlookの絞り込み(Restrict)に失敗しました: {restrict_error}")
                items = target_folder.Items

        item = items.GetFirst()
        while item:

            is_processed = False
            mail_entry_id = 'UNKNOWN'
            mail_item = None

            if item.Class == 43: # olMailItem (メールアイテムのみを処理)

                extraction_succeeded = False

                try:
                    mail_item = item

                    is_processed = False
                    mail_entry_id = str(getattr(mail_item, 'EntryID', 'UNKNOWN'))

                    if hasattr(item, 'Categories'):
                        current_categories = str(getattr(item, 'Categories', ''))
                        if PROCESSED_CATEGORY_NAME in current_categories:
                            is_processed = True

                    if read_mode == "unprocessed" and is_processed:
                        item = items.GetNext()
                        continue
                    
                    # (Restrictが失敗した場合のフォールバック)
                    if days_ago is not None:
                         received_time_check = getattr(mail_item, 'ReceivedTime', datetime.datetime.now())
                         if received_time_check.tzinfo is not None:
                             received_time_check = received_time_check.replace(tzinfo=None)
                         if received_time_check < start_date_dt:
                             item = items.GetNext()
                             continue

                    subject = str(getattr(mail_item, 'Subject', ''))
                    body = str(getattr(mail_item, 'Body', ''))
                    received_time = getattr(mail_item, 'ReceivedTime', datetime.datetime.now())

                    if received_time is not None and received_time.tzinfo is not None:
                        received_time = received_time.replace(tzinfo=None)
                    elif received_time is None:
                        received_time = datetime.datetime.now().replace(tzinfo=None)

                    attachments_text = ""
                    attachment_names = []

                    has_files = hasattr(mail_item, 'Attachments') and mail_item.Attachments.Count > 0

                    if has_files:
                        attachment_names = [att.FileName for att in mail_item.Attachments]

                        if is_processed and mail_entry_id in previous_attachment_content:
                            attachments_text = str(previous_attachment_content.get(mail_entry_id, ""))
                        else:
                            for attachment in mail_item.Attachments:
                                safe_filename = re.sub(r'[\\/:*?"<>|]', '_', attachment.FileName)
                                temp_file_path = os.path.join(temp_dir, f"{uuid.uuid4().hex}_{safe_filename}")
                                try:
                                    attachment.SaveAsFile(temp_file_path)
                                    extracted_content = get_attachment_text(temp_file_path, attachment.FileName)
                                    attachments_text += f"\n--- FILE: {attachment.FileName} ---\n{str(extracted_content)}\n"
                                except Exception as file_ex:
                                    attachments_text += f"\n--- ERROR reading {attachment.FileName}: {file_ex} ---\n"
                                finally:
                                    if os.path.exists(temp_file_path):
                                        os.remove(temp_file_path)
                            attachments_text = attachments_text.strip()

                    body_subject_search_text = str(subject) + " " + str(body)
                    must_include_body = any(re.search(kw, body_subject_search_text, re.IGNORECASE) for kw in MUST_INCLUDE_KEYWORDS)
                    has_initials = re.search(INITIALS_REGEX, body_subject_search_text)

                    must_include_attach_text = any(re.search(kw, attachments_text, re.IGNORECASE) for kw in MUST_INCLUDE_KEYWORDS)

                    must_include = must_include_body or must_include_attach_text or has_initials

                    full_search_text = body_subject_search_text + " " + attachments_text
                    is_excluded = any(re.search(kw, full_search_text, re.IGNORECASE) for kw in EXCLUDE_KEYWORDS)


                    if is_excluded:
                         if not is_processed:
                             mark_email_as_processed(mail_item) # ノイズとしてマーク
                         item = items.GetNext()
                         continue

                    is_target = has_files or must_include

                    if is_target and not is_processed:
                        pass
                    elif is_target and is_processed:
                        pass
                    elif not is_target and not is_processed:
                        mark_email_as_processed(mail_item) # ノイズとしてマーク
                        item = items.GetNext()
                        continue
                    elif not is_target and is_processed:
                        item = items.GetNext()
                        continue

                    # レコードの準備
                    record = {
                        'EntryID': mail_entry_id,
                        '件名': subject,
                        '受信日時': received_time,
                        '本文(テキスト形式)': body,
                        '本文(ファイル含む)': attachments_text,
                        'Attachments': ", ".join(attachment_names),
                    }
                    data_records.append(record)

                    extraction_succeeded = True

                except Exception as item_ex:
                    print(f"警告: メールアイテムの処理中にエラーが発生しました (EntryID: {mail_entry_id}). スキップします。エラー: {item_ex}")
                    if mail_item and not is_processed:
                        try:
                            mark_email_as_processed(mail_item)
                        except Exception as mark_e:
                            print(f"  警告: エラー発生後のマーク付けにも失敗しました: {mark_e}")

                    item = items.GetNext()
                    continue

                if extraction_succeeded and not is_processed:
                    mark_email_as_processed(mail_item)

            # ループの最後に次のアイテムを取得
            item = items.GetNext()

    except Exception as e:
        raise RuntimeError(f"Outlook操作エラー: {e}\n詳細: {traceback.format_exc()}")
    finally:
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