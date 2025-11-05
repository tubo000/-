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

# 📌 修正: PROCESSED_CATEGORY_NAME はもう使わない
# PROCESSED_CATEGORY_NAME = "スキルシート処理済"

# ----------------------------------------------------------------------
# 💡 ヘルパー関数: 過去の本文データ復元 (sqlite3版)
# ----------------------------------------------------------------------
def _load_previous_attachment_content() -> Dict[str, str]:
    db_path = os.path.join(os.path.abspath(SCRIPT_DIR), DATABASE_NAME)
    if os.path.exists(db_path):
        try:
            conn = sqlite3.connect(db_path)
            # 📌 修正: テーブルが存在しないエラーを考慮
            try:
                df_prev = pd.read_sql_query("SELECT \"EntryID\", \"本文(ファイル含む)\" FROM emails", conn)
            except pd.io.sql.DatabaseError:
                return {} # テーブルがなければ空
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
    """
    📌 修正: この関数はDB台帳方式では使用されないが、
    main_application.py 側で呼び出しが残っている場合に備えて定義だけ残す。
    (ただし、何もしない)
    """
    pass # 何もしない

# ----------------------------------------------------------------------
# 💡 処理済みカテゴリの解除 (削除)
# ----------------------------------------------------------------------
def remove_processed_category(target_email: str, folder_path: str, days_ago: int = None) -> int:
    """
    📌 修正: この関数はDB台帳方式では使用されないため、0を返すだけにする。
    """
    print("INFO: remove_processed_category は DB台帳方式では使用されません。")
    return 0


# ----------------------------------------------------------------------
# 💡 未処理メールの件数をカウント (COM初期化削除 + ログ削除)
# ----------------------------------------------------------------------
# email_processor.py (L220 付近)
def has_unprocessed_mail(folder_path: str, target_email: str, days_to_check: int = None) -> int:
    """
    【軽量版】
    指定されたフォルダの「最新300件」だけをチェックし、
    その中にDB未登録のメールが何件あるかをカウントする。
    """
    unprocessed_count = 0
    if not folder_path or not target_email: return 0

    outlook_ids_latest_300 = set()
    db_ids = set()
    
    # 1. データベースから「抽出済みID」と「スキップ済みID」をすべて取得
    # (OutlookよりDB接続の方がはるかに速いので、先に読み込む)
    db_path = os.path.join(os.path.abspath(SCRIPT_DIR), DATABASE_NAME)
    if os.path.exists(db_path):
        conn_check = None
        try:
            conn_check = sqlite3.connect(db_path)
            try:
                 db_ids.update(pd.read_sql_query("SELECT EntryID FROM emails", conn_check)['EntryID'].tolist())
            except Exception:
                 pass
            try:
                 db_ids.update(pd.read_sql_query("SELECT EntryID FROM skipped_ids", conn_check)['EntryID'].tolist())
            except Exception:
                 pass
        except Exception as e:
            print(f"警告: 既存DB(has_unprocessed)のEntryID読み込み失敗: {e}。")
        finally:
            if conn_check: conn_check.close()

    # 2. Outlook から「最新300件」の EntryID だけを取得
    try:
        try:
            outlook = win32.GetActiveObject("Outlook.Application")
        except:
            outlook = win32.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        folder = get_outlook_folder(namespace, target_email, folder_path)

        if folder:
            items = folder.Items
            
            # ▼▼▼【高速化】ここで日付絞り込み(Restrict)は行わない ▼▼▼
            # (Restrict自体が重いため、ソートして先頭から取るほうが速い)
            
            try: 
                items.Sort("[ReceivedTime]", True) # ★ 新しい順にソート
            except Exception as sort_error: 
                print(f"警告(has_unprocessed): Sort失敗: {sort_error}")

            item = items.GetFirst()
            
            # ▼▼▼【高速化】最新300件だけをチェック ▼▼▼
            count = 0
            while item and count < 300:
                 try:
                    if item.Class == 43:
                         outlook_ids_latest_300.add(str(getattr(item, 'EntryID', '')))
                 except Exception:
                     pass
                 try:
                     item = items.GetNext()
                     count += 1
                 except:
                     break
        
        # print(f"INFO (has_unprocessed): Outlookから最新 {len(outlook_ids_latest_300)} 件のIDを取得しました。") # ログ削除

    except Exception as e:
        print(f"警告(has_unprocessed Main): Outlook読み込みエラー: {e}")
        return 0

    # 3. 差分（最新300件にあってDBにないもの）をカウント
    unprocessed_count = len(outlook_ids_latest_300 - db_ids)
    
    # print(f"INFO (has_unprocessed): DBに {len(db_ids)} 件のIDあり。未処理件数: {unprocessed_count} 件") # ログ削除

    return unprocessed_count
# ----------------------------------------------------------------------
# 💡 メイン抽出関数: Outlookからメールを取得 (バッチ処理・待機機能付き)
# ----------------------------------------------------------------------
# email_processor.py (L330 付近)
# (import time, Iterator などはファイル先頭にある想定)

def get_mail_data_from_outlook_in_memory(target_folder_path: str, account_name: str, read_mode: str = "all", days_ago: int = None, main_elements: dict = None) -> Iterator[dict]:
    """
    Outlookからメールデータを抽出する (ジェネレータ)。
    軽いプロパティで先にスキップ判定を行い、重い処理を後回しにすることで高速化。
    抽出対象データとスキップ対象IDをバッチで yield する。
    """
    data_records_batch = [] 
    skip_ids_batch = []     
    
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
             days_ago = None
             start_date_dt = None
             log_period_message = "全期間 (入力不正)"

    # print(f"INFO: Outlookメール読み込み開始 (対象期間: {log_period_message})") # ログ削除

    existing_ids_set = set()
    db_path = os.path.join(os.path.abspath(SCRIPT_DIR), DATABASE_NAME)
    if os.path.exists(db_path):
        conn_check = None
        try:
            conn_check = sqlite3.connect(db_path)
            try:
                 existing_ids_set.update(pd.read_sql_query("SELECT EntryID FROM emails", conn_check)['EntryID'].tolist())
            except pd.io.sql.DatabaseError:
                 pass 
            except Exception as e_emails:
                 print(f"警告: 既存DB(emails)のEntryID読み込み失敗: {e_emails}。")

            try:
                 existing_ids_set.update(pd.read_sql_query("SELECT EntryID FROM skipped_ids", conn_check)['EntryID'].tolist())
            except pd.io.sql.DatabaseError:
                 pass
            except Exception as e_skipped:
                 print(f"警告: 既存DB(skipped_ids)のEntryID読み込み失敗: {e_skipped}。")
                 
        except Exception as e:
            print(f"警告: 既存DBのEntryID読み込み失敗: {e}。全件新規として扱います。")
            existing_ids_set = set()
        finally:
            if conn_check: conn_check.close()
            
    ids_processed_this_session = set()
    
    processed_item_count = 0 
    batch_size = 300         
    pause_duration = 3       # 3秒待機
    gui_queue = main_elements.get("gui_queue") if main_elements else None

    try:
        outlook_app = None
        try:
            outlook_app = win32.GetActiveObject("Outlook.Application")
        except:
            outlook_app = win32.Dispatch("Outlook.Application")
        outlook_ns = outlook_app.GetNamespace("MAPI")
        target_folder = get_outlook_folder(outlook_ns, account_name, target_folder_path)
        if target_folder is None: raise RuntimeError(f"指定フォルダ '{target_folder_path}' が見つかりません。")

        items = target_folder.Items

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
            if processed_item_count > 0 and processed_item_count % batch_size == 0:
                status_message = f"状態: {processed_item_count}件スキャン完了。DB保存中..."
                print(f"INFO: {status_message}")
                if gui_queue: gui_queue.put(status_message)
                
                yield {
                    "data_batch": pd.DataFrame(data_records_batch),
                    "skip_ids": skip_ids_batch
                }
                
                data_records_batch.clear() 
                skip_ids_batch.clear()
                
                status_message_wait = f"状態: {processed_item_count}件スキャン。{pause_duration}秒待機中..."
                if gui_queue: gui_queue.put(status_message_wait)
                print(f"INFO: {status_message_wait}")
                time.sleep(pause_duration)
                
                if gui_queue: gui_queue.put(f"状態: {processed_item_count}件スキャン。処理再開...")

            processed_item_count += 1
            mail_entry_id = 'UNKNOWN'
            mail_item = None
            received_time = datetime.datetime.now().replace(tzinfo=None)
            
            # --- ▼▼▼【修正】変数をループの先頭で初期化 ▼▼▼ ---
            subject = "[件名取得エラー]"
            body = "[本文取得エラー]"
            attachments_text = ""
            attachment_names = []
            has_files = False
            attachments_collection = None
            is_target = False
            # --- ▲▲▲ 修正ここまで ▲▲▲ ---

            if item.Class == 43:
                skip_reason = None
                try:
                    mail_item = item
                    # 1. 軽い処理を先に実行
                    try:
                        mail_entry_id = str(getattr(mail_item, 'EntryID', 'UNKNOWN_ID'))
                    except Exception as id_err:
                         mail_entry_id = f"ERROR_ID_{uuid.uuid4().hex}"
                         is_already_in_db = False
                    else:
                        is_in_db = mail_entry_id in existing_ids_set
                        is_in_session = mail_entry_id in ids_processed_this_session
                        is_already_in_db = is_in_db or is_in_session 
                    
                    try:
                        received_time_check = getattr(mail_item, 'ReceivedTime', datetime.datetime.now())
                        if received_time_check.tzinfo is not None:
                            received_time_check = received_time_check.replace(tzinfo=None)
                        received_time = received_time_check
                    except Exception as rt_err:
                         received_time = datetime.datetime.now().replace(tzinfo=None)
                         
                    # 2. 最速スキップ判定 (DB登録済みか？期間外か？)
                    # 📌 修正: Outlookカテゴリ (is_processed) でのスキップは削除
                    if read_mode == "unprocessed" and is_already_in_db:
                         skip_reason = "DB or Session processed"
                    elif start_date_dt is not None and received_time < start_date_dt:
                         skip_reason = f"期間外"
                    
                    if skip_reason:
                        pass 
                    
                    else:
                        # --- 3. スキップしないメールのみ、重い処理を実行 ---
                        try:
                            subject = str(getattr(mail_item, 'Subject', ''))
                        except Exception as subj_err:
                            subject = "[件名取得エラー]"

                        try:
                            body = str(getattr(mail_item, 'Body', ''))
                        except Exception as body_err:
                            body = "[本文取得エラー]"

                        try:
                            if mail_item and hasattr(mail_item, 'Attachments'):
                                 attachments_collection = mail_item.Attachments
                                 attachment_count = attachments_collection.Count
                                 if attachment_count > 0:
                                     has_files = True
                                     attachment_names = [att.FileName for att in attachments_collection if hasattr(att, 'FileName')]
                        except Exception as attach_err:
                             print(f"  -> 警告(has_files): 添付情報取得エラー: {attach_err}")

                        # 📌 修正: 'has_files' はここで必ず True/False が設定されている
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
                             if not is_already_in_db: 
                                 skip_ids_batch.append(mail_entry_id)
                                 ids_processed_this_session.add(mail_entry_id)
                        
                        else:
                            is_target = has_must_include_keyword or (has_files and has_initials_in_filename)

                            if is_target:
                                if not is_already_in_db:
                                    record = {
                                        'EntryID': mail_entry_id, '件名': subject, '受信日時': received_time,
                                        '本文(テキスト形式)': body, '本文(ファイル含む)': attachments_text,
                                        'Attachments': ", ".join(attachment_names),
                                    }
                                    data_records_batch.append(record)
                                    ids_processed_this_session.add(mail_entry_id)
                            elif not is_target:
                                if not is_already_in_db: 
                                    skip_ids_batch.append(mail_entry_id)
                                    ids_processed_this_session.add(mail_entry_id)

                except (pythoncom.com_error, AttributeError, Exception) as item_ex:
                    current_id = mail_entry_id if mail_entry_id != 'UNKNOWN' else getattr(item, 'EntryID', 'ID取得失敗')
                    print(f"警告(Item Loop): 処理中にエラー (ID: {current_id}): {item_ex}\n{traceback.format_exc(limit=1)}")
                finally:
                      pass
            
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
        if data_records_batch or skip_ids_batch: 
            print(f"INFO: 最後のバッチ (抽出:{len(data_records_batch)}件, スキップ:{len(skip_ids_batch)}件) を返します。")
            yield {
                "data_batch": pd.DataFrame(data_records_batch),
                "skip_ids": skip_ids_batch
            }
            data_records_batch.clear()
            skip_ids_batch.clear()

        if os.path.exists(temp_dir):
             try:
                 if not os.listdir(temp_dir): os.rmdir(temp_dir)
             except OSError as oe: print(f"警告: 一時フォルダクリーンアップ失敗: {oe}")
        pass

    # --- 最終的な return は削除 (generator のため) ---

# ----------------------------------------------------------------------
# 💡 外部公開関数
# ----------------------------------------------------------------------
def run_email_extraction(target_email: str, read_mode: str = "all", days_ago: int = None):
    pass

def delete_old_emails_core(target_email: str, folder_path: str, days_ago: int) -> int:
    pass