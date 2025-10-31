# email_processor.py (ログ出力削除・COM初期化削除版)

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
from typing import Dict, Any, List, Iterator # ← ★ Iterator を追加 ★
import threading # スレッドID取得に必要（ただしログ削除したので不要かも）
import time # 📌 5秒待機のために time モジュールをインポート

# ----------------------------------------------------------------------
# イニシャルを検出する正規表現を追加
# ----------------------------------------------------------------------
INITIALS_REGEX = r'(\b[A-Z]{2}\b|\b[A-Z]\s*.\s*[A-Z]\b|名前\([A-Z]{2}\))'

# --- インポート処理 ---

# 1. get_attachment_text のデフォルト（代替）定義
def get_attachment_text(*args, **kwargs):
    # print("警告: file_processor.py から get_attachment_text を読み込めませんでした。")
    return "ATTACHMENT_CONTENT_IMPORT_FAILED"

# 2. get_outlook_folder のデフォルト（代替）定義
def get_outlook_folder(outlook_ns, account_name, folder_path):
     # print(f"警告: config.py から get_outlook_folder を読み込めませんでした。デフォルト処理を使用します。")
     try:
          return outlook_ns.Folders[account_name].Folders[folder_path]
     except Exception:
          # print(f"エラー: デフォルトのフォルダ取得も失敗しました: {account_name}/{folder_path}")
          return None

# 3. config.py から設定値と関数を読み込む
try:
# ▼▼▼ DATABASE_NAME をインポート対象に追加 ▼▼▼
    from config import MUST_INCLUDE_KEYWORDS, EXCLUDE_KEYWORDS, SCRIPT_DIR, OUTPUT_CSV_FILE as OUTPUT_FILENAME, DATABASE_NAME
    # ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    
    # ▼▼▼ 修正点 ▼▼▼
    try:
        from config import get_outlook_folder as real_get_outlook_folder
        get_outlook_folder = real_get_outlook_folder
        # print("INFO: config.py から get_outlook_folder を読み込みました。")
    except ImportError:
        # print("警告: config.py に get_outlook_folder が定義されていません。デフォルト処理を使用します。")
        pass
    # print("INFO: config.py から設定値を読み込みました。")
except ImportError:
    # print("警告: config.py が見つからないかインポートできませんでした。デフォルト設定を使用します。")
    MUST_INCLUDE_KEYWORDS = [r'スキルシート']
    EXCLUDE_KEYWORDS = [r'案\s*件\s*名',r'案\s*件\s*番\s*号',r'案\s*件:',r'案\s*件：',r'【案\s*件】',r'必\s*須']
    SCRIPT_DIR = os.getcwd()
    OUTPUT_FILENAME = 'output_extraction.xlsx'

# 4. file_processor.py から関数を読み込む
try:
    from file_processor import get_attachment_text as real_get_attachment_text
    get_attachment_text = real_get_attachment_text
    # print("INFO: file_processor.py から get_attachment_text を読み込みました。")
except ImportError:
    # print("警告: file_processor.py が見つからないか 'get_attachment_text' が含まれていません。")
    pass
except Exception as e:
    # print(f"エラー: file_processor.py のインポート中にエラー: {e}")
    pass

# --- 修正ここまで ---
#DATABASE_NAME = 'extraction_cache.db'
PROCESSED_CATEGORY_NAME = "スキルシート処理済"

# ----------------------------------------------------------------------
# 💡 ヘルパー関数: 過去の本文データ復元 (sqlite3版)
# ----------------------------------------------------------------------
def _load_previous_attachment_content() -> Dict[str, str]:
    db_path = os.path.join(os.path.abspath(SCRIPT_DIR), DATABASE_NAME)
    if os.path.exists(db_path):
        try:
            conn = sqlite3.connect(db_path)
            df_prev = pd.read_sql_query("SELECT \"EntryID\", \"本文(ファイル含む)\" FROM emails", conn)
            conn.close()
            df_prev.set_index('EntryID', inplace=True)
            return df_prev['本文(ファイル含む)'].dropna().to_dict()
        except Exception as e:
            # print(f"警告: データベースからの本文復元に失敗しました。エラー: {e}")
            return {}
    return {}

# ----------------------------------------------------------------------
# 💡 共通機能: メールアイテムの処理済みマーク (維持)
# ----------------------------------------------------------------------
def mark_email_as_processed(mail_item):
    if mail_item.Class == 43:
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
# 💡 処理済みカテゴリの解除 (COM初期化削除 + ログ削除)
# ----------------------------------------------------------------------
def remove_processed_category(target_email: str, folder_path: str, days_ago: int = None) -> int:
    reset_count = 0
    start_date_dt = None

    if days_ago is not None:
        try:
             days_ago = int(days_ago)
             if days_ago <= 0:
                  days_ago = None
             else:
                  start_date_dt = (datetime.datetime.now() - timedelta(days=days_ago))
        except (ValueError, TypeError):
             days_ago = None 
             start_date_dt = None

    try:
        outlook = None
        try:
            outlook = win32.GetActiveObject("Outlook.Application")
        except:
            try:
                 outlook = win32.Dispatch("Outlook.Application")
            except Exception as dispatch_err:
                 raise RuntimeError(f"Outlook の起動/接続に失敗しました: {dispatch_err}")

        namespace = outlook.GetNamespace("MAPI")
        folder = get_outlook_folder(namespace, target_email, folder_path)
        if folder is None:
            raise RuntimeError(f"指定フォルダ '{folder_path}' が見つかりません。")

        items = folder.Items
        
        filter_query_list = []
        if start_date_dt is not None:
            start_date_str = start_date_dt.strftime('%Y/%m/%d %H:%M')
            filter_query_list.append(f"[ReceivedTime] < '{start_date_str}'")

        query_string = " AND ".join(filter_query_list)
        items_to_reset = items

        if query_string:
            try:
                items_to_reset = items.Restrict(query_string)
            except Exception as restrict_error:
                print(f"警告: カテゴリ解除のRestrict(日付)に失敗: {restrict_error}。全件チェックにフォールバックします。") # この警告は残す

        try:
            items_to_reset.Sort("[ReceivedTime]", True)
        except Exception as sort_err:
             print(f"警告(remove_category): ソート失敗: {sort_err}") # この警告は残す

        item = items_to_reset.GetFirst()
        
        while item:
            if item.Class == 43:
                try:
                    current_categories = getattr(item, 'Categories', '')
                    if PROCESSED_CATEGORY_NAME in current_categories:
                        is_target_date = True
                        if start_date_dt is not None:
                            received_time = getattr(item, 'ReceivedTime', datetime.datetime.now())
                            if received_time.tzinfo is not None:
                                received_time = received_time.replace(tzinfo=None)
                            if received_time >= start_date_dt:
                                is_target_date = False

                        if is_target_date:
                            try:
                                categories_list = [c.strip() for c in current_categories.split(',') if c.strip() != PROCESSED_CATEGORY_NAME]
                                new_categories = ", ".join(categories_list)
                                item.Categories = new_categories
                                item.Save()
                                reset_count += 1
                            except Exception as save_err:
                                 print(f"エラー(remove_category): カテゴリ保存/Save失敗: {save_err}") # このエラーは残す
                        
                except pythoncom.com_error as com_err:
                     print(f"警告(remove_category Loop): アイテム処理中 COMエラー: {com_err.hresult if hasattr(com_err, 'hresult') else 'N/A'}") # この警告は残す
                except Exception as e:
                    print(f"警告(remove_category Loop): アイテム処理中エラー: {e}") # この警告は残す
            
            try:
                item = items_to_reset.GetNext()
            except:
                break
                
    except Exception as e:
        print(f"エラー(remove_category Main): カテゴリマーク解除中に予期せぬエラー: {e}\n{traceback.format_exc(limit=2)}")
        reset_count = -1
    finally:
        pass 

    return reset_count


# ----------------------------------------------------------------------
# 💡 未処理メールの件数をカウント (COM初期化削除 + ログ削除)
# ----------------------------------------------------------------------
def has_unprocessed_mail(folder_path: str, target_email: str, days_to_check: int = None) -> int:
    unprocessed_count = 0
    if not folder_path or not target_email: return 0

    valid_days_to_check = None
    cutoff_date_dt = None 

    if days_to_check is not None:
        try:
            days_to_check_int = int(days_to_check)
            if days_to_check_int >= 0:
                valid_days_to_check = days_to_check_int
                if valid_days_to_check == 0:
                    cutoff_date_dt = datetime.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
                else:
                    cutoff_date_dt = datetime.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0) - timedelta(days=valid_days_to_check)
        except (ValueError, TypeError):
             pass # ログ削除

    try:
        try:
            outlook = win32.GetActiveObject("Outlook.Application")
        except:
            outlook = win32.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        folder = get_outlook_folder(namespace, target_email, folder_path)

        if folder:
            items = folder.Items

            if cutoff_date_dt is not None:
                try:
                    cutoff_date_str = cutoff_date_dt.strftime('%Y/%m/%d %H:%M')
                    date_filter = f"[ReceivedTime] >= '{cutoff_date_str}'"
                    items = items.Restrict(date_filter)
                except Exception as restrict_error:
                    print(f"警告: has_unprocessed_mailの日付絞り込み失敗。全件スキャン: {restrict_error}") # この警告は残す
                    items = folder.Items
            
            try: items.Sort("[ReceivedTime]", True)
            except Exception as sort_error: print(f"警告(has_unprocessed): Sort失敗: {sort_error}") # この警告は残す

            item = items.GetFirst()
            
            while item:
                 mail_entry_id_debug = 'UNKNOWN_ID'
                 try:
                    mail_entry_id_debug = getattr(item, 'EntryID', 'UNKNOWN_ID')
                    if item.Class == 43:
                         categories = str(getattr(item, 'Categories', ''))
                         if PROCESSED_CATEGORY_NAME not in categories:
                            has_files = False
                            has_initials_in_filename = False
                            try:
                                if item and hasattr(item, 'Attachments'):
                                     attachments_collection = item.Attachments
                                     attachment_count = attachments_collection.Count
                                     if attachment_count > 0:
                                         has_files = True
                                         attachment_names = [att.FileName for att in attachments_collection if hasattr(att, 'FileName')]
                                         all_filenames_text = " ".join(attachment_names)
                                         if re.search(INITIALS_REGEX, all_filenames_text):
                                             has_initials_in_filename = True
                            except (pythoncom.com_error, AttributeError, Exception) as attach_err:
                                 print(f"警告(has_unprocessed): 添付情報/名前チェックエラー (ID: {mail_entry_id_debug}): {attach_err}") # この警告は残す

                            subject = str(getattr(item, 'Subject', ''))
                            body = str(getattr(item, 'Body', ''))
                            full_search_text = subject + " " + body
                            must_include = any(re.search(kw, full_search_text, re.IGNORECASE) for kw in MUST_INCLUDE_KEYWORDS)
                            has_initials_in_text = re.search(INITIALS_REGEX, full_search_text) 
                            is_target_for_count = must_include or has_initials_in_text or (has_files and has_initials_in_filename)

                            if is_target_for_count:
                                unprocessed_count += 1
                                
                 except pythoncom.com_error as com_err:
                      print(f"警告(has_unprocessed Loop): COMエラー (ID: {mail_entry_id_debug}): {com_err.hresult if hasattr(com_err, 'hresult') else 'N/A'}") # この警告は残す
                 except Exception as e:
                     print(f"警告(has_unprocessed Loop): アイテム処理エラー (ID: {mail_entry_id_debug}): {e}") # この警告は残す

                 try:
                     item = items.GetNext()
                 except:
                     break

    except Exception as e:
        print(f"警告(has_unprocessed Main): チェック処理エラー: {e}") # この警告は残す
        unprocessed_count = 0
    finally:
        pass 

    return unprocessed_count

# ----------------------------------------------------------------------
# 💡 メイン抽出関数: Outlookからメールを取得 (バッチ処理・待機機能付き)
# ----------------------------------------------------------------------
def get_mail_data_from_outlook_in_memory(target_folder_path: str, account_name: str, read_mode: str = "all", days_ago: int = None, main_elements: dict = None) -> Iterator[pd.DataFrame]:
    """
    Outlookからメールデータを抽出する (ジェネレータ)。
    300件スキャンするごとにバッチ (DataFrame) を yield (返送) し、5秒待機する。
    """
    # data_records はバッチごとにリセットされる
    data_records_batch = [] 
    temp_dir = os.path.join(SCRIPT_DIR, "temp_attachments_safe")
    os.makedirs(temp_dir, exist_ok=True)

    previous_attachment_content = _load_previous_attachment_content()

    start_date_dt = None
    log_period_message = "全期間" 

    if days_ago is not None:
        try:
             days_ago = int(days_ago)
             if days_ago < 0: raise ValueError("日数は0以上")
             if days_ago == 0:
                 today_date = datetime.date.today()
                 start_date_dt = datetime.datetime.combine(today_date, datetime.time.min)
                 log_period_message = "今日のみ"
             else:
                 start_date_dt = datetime.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0) - timedelta(days=days_ago)
                 log_period_message = f"過去{days_ago}日間"
        except ValueError as e:
             # print(f"警告: 不正日数 '{days_ago}', 全期間対象. Error: {e}") # ログ削除
             days_ago = None
             start_date_dt = None
             log_period_message = "全期間 (入力不正)"

    # print(f"INFO: Outlookメール読み込み開始 (対象期間: {log_period_message})") # ログ削除

    existing_ids_set = set()
    db_path = os.path.join(os.path.abspath(SCRIPT_DIR), DATABASE_NAME)
    if os.path.exists(db_path):
        try:
            conn_check = sqlite3.connect(db_path)
            existing_ids_set = set(pd.read_sql_query("SELECT EntryID FROM emails", conn_check)['EntryID'].tolist())
            conn_check.close()
            # print(f"INFO: 既存DBから {len(existing_ids_set)} 件のEntryIDを読み込みました。") # ログ削除
        except Exception as e:
            print(f"警告: 既存DBのEntryID読み込み失敗: {e}。全件新規として扱います。")
            existing_ids_set = set()
    
    # --- ▼▼▼ バッチ処理用の設定 ▼▼▼ ---
    processed_item_count = 0 # スキャンした総数
    batch_size = 300         # 300件ごとに処理
    pause_duration = 5       # 5秒間停止
    gui_queue = main_elements.get("gui_queue") if main_elements else None
    # --- ▲▲▲ バッチ処理用の設定 ▲▲▲ ---

    try:
        # --- 📌 CoInitialize() 削除 (スレッド側で実行) ---
        outlook_app = None
        try:
            outlook_app = win32.GetActiveObject("Outlook.Application")
        except:
            outlook_app = win32.Dispatch("Outlook.Application")
        outlook_ns = outlook_app.GetNamespace("MAPI")
        target_folder = get_outlook_folder(outlook_ns, account_name, target_folder_path)
        if target_folder is None: raise RuntimeError(f"指定フォルダ '{target_folder_path}' が見つかりません。")

        items = target_folder.Items

        # --- 日付絞り込み (Restrict) ---
        filter_query_list = []
        if start_date_dt is not None:
            start_date_str = start_date_dt.strftime('%Y/%m/%d %H:%M')
            filter_query_list.append(f"[ReceivedTime] >= '{start_date_str}'")

        if filter_query_list:
            query_string = " AND ".join(filter_query_list)
            try:
                items = items.Restrict(query_string)
            except Exception as restrict_error:
                print(f"警告: Outlook Restrict失敗: {restrict_error}")
                items = target_folder.Items
                
        try:
            items.Sort("[ReceivedTime]", True)
        except Exception as sort_error:
             print(f"警告: Outlook Sort失敗: {sort_error}")

        item = items.GetFirst()

        while item:
            
            # --- ▼▼▼ バッチ処理（一時停止 & yield） ▼▼▼ ---
            if processed_item_count > 0 and processed_item_count % batch_size == 0:
                status_message = f"状態: {processed_item_count}件スキャン完了。DB保存中..."
                print(f"INFO: {status_message}") # ログは残す
                if gui_queue: gui_queue.put(status_message)
                
                # ★★★ 現在のバッチ(data_records_batch)をDataFrameにして返す (yield) ★★★
                df_batch = pd.DataFrame(data_records_batch)
                yield df_batch # <-- ★ これがジェネレータの「返す」動作
                
                # バッチリストをクリア
                data_records_batch.clear() 
                
                # 5秒待機
                status_message_wait = f"状態: {processed_item_count}件スキャン。{pause_duration}秒待機中..."
                if gui_queue: gui_queue.put(status_message_wait)
                print(f"INFO: {status_message_wait}") # ログは残す
                time.sleep(pause_duration)
                
                if gui_queue: gui_queue.put(f"状態: {processed_item_count}件スキャン。処理再開...")
            # --- ▲▲▲ バッチ処理ここまで ▲▲▲ ---

            processed_item_count += 1
            is_processed = False
            mail_entry_id = 'UNKNOWN'
            mail_item = None
            subject = "[件名取得エラー]"
            body = "[本文取得エラー]"
            received_time = datetime.datetime.now().replace(tzinfo=None)
            attachments_text = ""
            attachment_names = []
            has_files = False
            attachments_collection = None
            extraction_succeeded = False
            is_target = False

            if item.Class == 43:
                skip_reason = None
                try:
                    mail_item = item
                    try:
                        mail_entry_id = str(getattr(mail_item, 'EntryID', 'UNKNOWN_ID'))
                    except Exception as id_err:
                         mail_entry_id = f"ERROR_ID_{uuid.uuid4().hex}"
                         is_already_in_db = False
                    else:
                         is_already_in_db = mail_entry_id in existing_ids_set
                    
                    try:
                        subject = str(getattr(mail_item, 'Subject', ''))
                    except Exception as subj_err:
                        subject = "[件名取得エラー]"

                    # print(f"\n[{processed_item_count}] 処理中: {subject[:50]}...") # ログ削除

                    try:
                        current_categories = getattr(mail_item, 'Categories', '')
                        if PROCESSED_CATEGORY_NAME in current_categories:
                            is_processed = True
                    except Exception as cat_err:
                         pass 

                    try:
                        received_time_check = getattr(mail_item, 'ReceivedTime', datetime.datetime.now())
                        if received_time_check.tzinfo is not None:
                            received_time_check = received_time_check.replace(tzinfo=None)
                        received_time = received_time_check
                    except Exception as rt_err:
                         received_time = datetime.datetime.now().replace(tzinfo=None)

                    try:
                        body = str(getattr(mail_item, 'Body', ''))
                    except Exception as body_err:
                        body = "[本文取得エラー]"

                    # --- スキップ判定 ---
                    if read_mode == "unprocessed" and is_processed:
                        skip_reason = "Outlook処理済み"
                    elif read_mode == "unprocessed" and is_already_in_db:
                         skip_reason = "DB登録済み"
                    elif start_date_dt is not None and received_time < start_date_dt:
                         skip_reason = f"期間外"
                    
                    if skip_reason:
                        # print(f"  -> スキップ: {skip_reason}") # ログ削除
                        pass 
                    
                    else:
                        # --- スキップ理由がない場合のみ、詳細な処理に進む ---
                        try:
                            if mail_item and hasattr(mail_item, 'Attachments'):
                                 attachments_collection = mail_item.Attachments
                                 attachment_count = attachments_collection.Count
                                 if attachment_count > 0:
                                     has_files = True
                                     attachment_names = [att.FileName for att in attachments_collection if hasattr(att, 'FileName')]
                        except Exception as attach_err:
                             print(f"  -> 警告(has_files): 添付情報取得エラー: {attach_err}")

                        if has_files and attachments_collection:
                            if not is_already_in_db:
                                 try:
                                    for attachment in attachments_collection:
                                        if not hasattr(attachment, 'FileName'): continue
                                        safe_filename = re.sub(r'[\\/:*?"<>|]', '_', attachment.FileName)
                                        if len(safe_filename) > 150:
                                             name, ext = os.path.splitext(safe_filename)
                                             safe_filename = name[:150-len(ext)] + ext
                                        temp_file_path = os.path.join(temp_dir, f"{uuid.uuid4().hex}_{safe_filename}")
                                        try:
                                            attachment.SaveAsFile(temp_file_path)
                                            extracted_content = get_attachment_text(temp_file_path, attachment.FileName)
                                            attachments_text += f"\n--- FILE: {attachment.FileName} ---\n{str(extracted_content)}\n"
                                        except pythoncom.com_error as com_err:
                                             print(f"エラー(Attach Save/Read): COMエラー (File: {attachment.FileName}, ID: {mail_entry_id}): {com_err}")
                                             attachments_text += f"\n--- ERROR reading {attachment.FileName}: COM Error ---\n"
                                        except Exception as file_ex:
                                            print(f"エラー(Attach Save/Read): 例外 (File: {attachment.FileName}, ID: {mail_entry_id}): {file_ex}")
                                            attachments_text += f"\n--- ERROR reading {attachment.FileName}: {file_ex} ---\n"
                                        finally:
                                            if os.path.exists(temp_file_path):
                                                try: os.remove(temp_file_path)
                                                except OSError as oe: print(f"警告: 一時ファイル削除失敗: {oe}")
                                 except Exception as loop_err:
                                      print(f"警告: 添付ファイルループ処理エラー (ID: {mail_entry_id}): {loop_err}")
                                      attachments_text += "\n--- ERROR during attachment loop ---\n"
                                 attachments_text = attachments_text.strip()

                        body_subject_search_text = subject + " " + body
                        search_text_for_keywords = body_subject_search_text + " " + attachments_text
                        has_must_include_keyword = any(re.search(kw, search_text_for_keywords, re.IGNORECASE) for kw in MUST_INCLUDE_KEYWORDS)
                        has_initials_in_filename = False
                        if has_files:
                            all_filenames_text = " ".join(attachment_names)
                            if re.search(INITIALS_REGEX, all_filenames_text): has_initials_in_filename = True
                        
                        full_search_text = body_subject_search_text + " " + attachments_text
                        is_excluded = False
                        matched_exclude_kw = None
                        for kw in EXCLUDE_KEYWORDS:
                            if re.search(kw, full_search_text, re.IGNORECASE):
                                is_excluded = True
                                matched_exclude_kw = kw
                                break
                                
                        if is_excluded:
                             # print(f"  -> スキップ: 除外キーワード '{matched_exclude_kw}' にマッチ") # ログ削除
                             if not is_processed: mark_email_as_processed(mail_item)
                        
                        else:
                            is_target = has_must_include_keyword or (has_files and has_initials_in_filename)
                            # print(f"  -> 判定: is_target={is_target} (...)") # ログ削除

                            if is_target:
                                if not is_already_in_db:
                                    # print(f"  -> ★★★ 新規抽出対象としてレコード追加 ★★★") # ログ削除
                                    record = {
                                        'EntryID': mail_entry_id, '件名': subject, '受信日時': received_time,
                                        '本文(テキスト形式)': body, '本文(ファイル含む)': attachments_text,
                                        'Attachments': ", ".join(attachment_names),
                                    }
                                    data_records_batch.append(record) # ★ バッチリストに追加
                                    extraction_succeeded = True
                                    # new_record_counter は削除
                            elif not is_target:
                                # print(f"  -> スキップ: 抽出対象外") # ログ削除
                                if not is_processed: mark_email_as_processed(mail_item)

                except (pythoncom.com_error, AttributeError, Exception) as item_ex:
                    current_id = mail_entry_id if mail_entry_id != 'UNKNOWN' else getattr(item, 'EntryID', 'ID取得失敗')
                    print(f"警告(Item Loop): 処理中にエラー (ID: {current_id}): {item_ex}\n{traceback.format_exc(limit=1)}")
                    if mail_item and not is_processed:
                        try: mark_email_as_processed(mail_item)
                        except Exception as mark_e: print(f"  -> 警告: エラー後のマーク付け失敗: {mark_e}")
                finally:
                      if extraction_succeeded and not is_processed:
                          if mail_item:
                              try:
                                  mark_email_as_processed(mail_item)
                                  # print(f"  -> INFO: 処理済みマークを付与") # ログ削除
                              except Exception as mark_e:
                                   print(f"  -> 警告: 抽出成功後のマーク付け失敗: {mark_e}")
            
            else:
                 pass 

            try:
                item = items.GetNext() 
            except (pythoncom.com_error, Exception) as next_err:
                 print(f"警告: GetNext() でエラー。ループ中断。エラー: {next_err}")
                 break 

    except pythoncom.com_error as com_outer_err:
         raise RuntimeError(f"Outlook操作エラー (COM): {com_outer_err}\n{traceback.format_exc()}")
    except Exception as e:
        raise RuntimeError(f"Outlook操作エラー: {e}\n{traceback.format_exc()}")
    finally:
        # --- ▼▼▼【修正】ループ終了後、残りのバッチを yield する ---
        if data_records_batch:
            print(f"INFO: 最後のバッチ {len(data_records_batch)} 件を返します。")
            df_batch = pd.DataFrame(data_records_batch)
            yield df_batch
            data_records_batch.clear()
        # --- ▲▲▲ 修正ここまで ▲▲▲ ---

        if os.path.exists(temp_dir):
             try:
                 if not os.listdir(temp_dir): os.rmdir(temp_dir)
             except OSError as oe: print(f"警告: 一時フォルダクリーンアップ失敗: {oe}")
        pass 

    # --- ▼▼▼【修正】最終的な return は削除 (generator のため) ---
    # print(f"INFO: Outlookメール読み込みループ終了。...")
    # df = pd.DataFrame(data_records)
    # ...
    # return df
    # --- ▲▲▲ 修正ここまで ▲▲▲ ---

# ----------------------------------------------------------------------
# 💡 外部公開関数
# ----------------------------------------------------------------------
def run_email_extraction(target_email: str, read_mode: str = "all", days_ago: int = None):
    pass

def delete_old_emails_core(target_email: str, folder_path: str, days_ago: int) -> int:
    pass