# email_processor.py (★ v3o エラー対策・ハイブリッド・デバッグ版 ★)
#11/13

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
from typing import Dict, Any, List, Iterator
import sqlite3
import threading
import time
import base64 # 👈 ★ 1. base64 をインポート (v3mで必要)
import pywintypes # 👈 ★ 2. エラー処理のためにインポート

# ----------------------------------------------------------------------
# イニシャルを検出する正規表現を追加
# ----------------------------------------------------------------------
INITIALS_REGEX = r'(\b[A-Z]{2}\b|\b[A-Z]\s*.\s*[A-Z]\b|名前\([A-Z]{2}\))'

# --- インポート処理 ---

# 1. get_attachment_text のデフォルト（代替）定義
def get_attachment_text(*args, **kwargs):
    return "ATTACHMENT_CONTENT_IMPORT_FAILED"

# 2. get_outlook_folder のデフォルト（代替）定義
def get_outlook_folder(outlook_ns, account_name, folder_path):
    try:
        return outlook_ns.Folders[account_name].Folders[folder_path]
    except Exception:
        return None

# 3. config.py から設定値と関数を読み込む
try:
    from config import MUST_INCLUDE_KEYWORDS, EXCLUDE_KEYWORDS, SCRIPT_DIR, OUTPUT_CSV_FILE as OUTPUT_FILENAME, DATABASE_NAME
    try:
        from config import get_outlook_folder as real_get_outlook_folder
        get_outlook_folder = real_get_outlook_folder
    except ImportError:
        pass
except ImportError:
    MUST_INCLUDE_KEYWORDS = [r'スキルシート']
    EXCLUDE_KEYWORDS = [r'案\s*件\s*名',r'案\s*件\s*番\s*号',r'案\s*件:',r'案\s*件：',r'【案\s*件】',r'必\s*須']
    SCRIPT_DIR = os.getcwd()
    OUTPUT_FILENAME = 'output_extraction.xlsx'

# 4. file_processor.py から関数を読み込む
try:
    # ★ 修正: file_processor (v1p/v3m) が両方の引数を受け取る前提
    from file_processor import get_attachment_text as real_get_attachment_text
    get_attachment_text = real_get_attachment_text
except ImportError:
    pass
except Exception as e:
    pass

PROCESSED_CATEGORY_NAME = "スキルシート処理済"

# ( ... _load_previous_attachment_content 関数 (変更なし) ... )
def _load_previous_attachment_content() -> Dict[str, str]:
    db_path = os.path.join(os.path.abspath(SCRIPT_DIR), DATABASE_NAME)
    if os.path.exists(db_path):
        try:
            conn = sqlite3.connect(db_path)
            try:
                df_prev = pd.read_sql_query("SELECT \"EntryID\", \"本文(ファイル含む)\" FROM emails", conn)
            except pd.io.sql.DatabaseError:
                return {} 
            conn.close()
            df_prev.set_index('EntryID', inplace=True)
            return df_prev['本文(ファイル含む)'].dropna().to_dict()
        except Exception as e:
            print(f"警告: データベースからの本文復元に失敗しました。エラー: {e}")
            return {}
    return {}

# ( ... mark_email_as_processed, remove_processed_category 関数 (変更なし) ... )
def mark_email_as_processed(mail_item):
    pass 

def remove_processed_category(target_email: str, folder_path: str, days_ago: int = None) -> int:
    print("INFO: remove_processed_category は DB台帳方式では使用されません。")
    return 0

# ( ... has_unprocessed_mail 関数 (変更なし) ... )
def has_unprocessed_mail(folder_path: str, target_email: str, days_to_check: int = None) -> int:
    """ (v34 のコードと変更なし) """
    unprocessed_count = 0
    if not folder_path or not target_email: return 0
    outlook_ids_latest_300 = set()
    db_ids = set()
    db_path = os.path.join(os.path.abspath(SCRIPT_DIR), DATABASE_NAME)
    if os.path.exists(db_path):
        conn_check = None
        try:
            conn_check = sqlite3.connect(db_path)
            try:
                db_ids.update(pd.read_sql_query("SELECT EntryID FROM emails", conn_check)['EntryID'].tolist())
            except Exception: pass
            try:
                db_ids.update(pd.read_sql_query("SELECT EntryID FROM skipped_ids", conn_check)['EntryID'].tolist())
            except Exception: pass
        except Exception as e:
            print(f"警告: 既存DB(has_unprocessed)のEntryID読み込み失敗: {e}。")
        finally:
            if conn_check: conn_check.close()
    try:
        try:
            outlook = win32.GetActiveObject("Outlook.Application")
        except:
            outlook = win32.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        folder = get_outlook_folder(namespace, target_email, folder_path)
        if folder:
            items = folder.Items
            try: 
                items.Sort("[ReceivedTime]", True) 
            except Exception as sort_error: 
                print(f"警告(has_unprocessed): Sort失敗: {sort_error}")
            item = items.GetFirst()
            count = 0
            while item and count < 300:
                try:
                    if item.Class == 43:
                        outlook_ids_latest_300.add(str(getattr(item, 'EntryID', '')))
                except Exception: pass
                try:
                    item = items.GetNext()
                    count += 1
                except: break
    except Exception as e:
        print(f"警告(has_unprocessed Main): Outlook読み込みエラー: {e}")
        return 0
    unprocessed_count = len(outlook_ids_latest_300 - db_ids)
    return unprocessed_count

# ----------------------------------------------------------------------
# 💡 メイン抽出関数: Outlookからメールを取得 (★ハイブリッド・デバッグ版★)
# ----------------------------------------------------------------------
def get_mail_data_from_outlook_in_memory(target_folder_path: str, account_name: str, read_mode: str = "all", days_ago: int = None, main_elements: dict = None) -> Iterator[dict]:
    """
    (★ v3o エラー対策・ハイブリッド・デバッグ版 ★)
    PropertyAccessor (メモリ) を試し、失敗したら SaveAsFile (ディスク) にフォールバックする
    """
    data_records_batch = [] 
    skip_ids_batch = []     
    
    # (★ 追加) ----------------------------------------------------
    # ▼▼▼ カウンタ変数を初期化 ▼▼▼
    count_cond1_attach = 0
    count_cond2_must = 0
    count_cond3_excluded = 0
    count_extracted_total = 0
    # ▲▲▲ カウンタ変数を初期化 ▲▲▲
    # -------------------------------------------------------------

    temp_dir = os.path.join(SCRIPT_DIR, "temp_attachments_safe")
    os.makedirs(temp_dir, exist_ok=True)
    previous_attachment_content = _load_previous_attachment_content()
    start_date_dt = None
    log_period_message = "全期間" 

    # ( ... 日付フィルタリング処理 (変更なし) ... )
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

    # ( ... 既存ID読み込み処理 (変更なし) ... )
    existing_ids_set = set()
    db_path = os.path.join(os.path.abspath(SCRIPT_DIR), DATABASE_NAME)
    if os.path.exists(db_path):
        conn_check = None
        try:
            conn_check = sqlite3.connect(db_path)
            try:
                existing_ids_set.update(pd.read_sql_query("SELECT EntryID FROM emails", conn_check)['EntryID'].tolist())
            except Exception: pass
            try:
                existing_ids_set.update(pd.read_sql_query("SELECT EntryID FROM skipped_ids", conn_check)['EntryID'].tolist())
            except Exception: pass
        except Exception as e:
            print(f"警告: 既存DBのEntryID読み込み失敗: {e}。全件新規として扱います。")
            existing_ids_set = set()
        finally:
            if conn_check: conn_check.close()
            
    ids_processed_this_session = set()
    
    processed_item_count = 0 
    batch_size = 300        
    pause_duration = 3      
    gui_queue = main_elements.get("gui_queue") if main_elements else None
    stop_flag = main_elements.get("stop_extraction_flag") if main_elements else threading.Event() 

    # ( ... Outlook接続・フィルタリング処理 (変更なし) ... )
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

        # --- ▼▼▼ ★★★ ハイブリッド版 添付ファイルループ ★★★ ▼▼▼
        # MAPIプロパティタグ (バイナリデータ)
        PR_ATTACH_DATA_BIN = "http://schemas.microsoft.com/mapi/proptag/0x37010102"
        # --- ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲ ---

        while item:
            if stop_flag.is_set(): 
                print("INFO: ユーザーにより処理が中断されました。")
                break 
            if processed_item_count > 0 and processed_item_count % batch_size == 0:
                # ( ... バッチ処理 ... )
                status_message = f"状態: {processed_item_count}件スキャン完了。DB保存中..."
                if gui_queue: gui_queue.put(status_message)

                # (★ 追加) --------------------------------------------
                # ▼▼▼ バッチ処理時のコンソール出力 ▼▼▼
                if (data_records_batch or skip_ids_batch):
                    print(f"--- バッチ処理結果 (スキャン {processed_item_count} 件時点) ---")
                    print(f"  (1) 添付ファイルあり (除外前): {count_cond1_attach} 件")
                    print(f"  (2) Mustキーワードあり(添付なし): {count_cond2_must} 件") # (★デバッグ表示名を変更)
                    print(f"  (3) 除外キーワードにより除外: {count_cond3_excluded} 件")
                    print(f"  [+] バッチ抽出件数: {count_extracted_total} 件")
                    print(f"----------------------------------------")
                    
                    # カウンタをリセット
                    count_cond1_attach = 0
                    count_cond2_must = 0
                    count_cond3_excluded = 0
                    count_extracted_total = 0
                # ▲▲▲ バッチ処理時のコンソール出力 ▲▲▲
                # -----------------------------------------------------

                yield { "data_batch": pd.DataFrame(data_records_batch), "skip_ids": skip_ids_batch }
                data_records_batch.clear(); skip_ids_batch.clear()
                status_message_wait = f"状態: {processed_item_count}件スキャン。{pause_duration}秒待機中..."
                if gui_queue: gui_queue.put(status_message_wait)
                time.sleep(pause_duration)
                if gui_queue: gui_queue.put(f"状態: {processed_item_count}件スキャン。処理再開...")

            processed_item_count += 1
            # ( ... メール基本情報取得 (変更なし) ... )
            mail_entry_id = 'UNKNOWN'
            mail_item = None
            received_time = datetime.datetime.now().replace(tzinfo=None)
            
            subject = "[件名取得エラー]"
            body = "[本文取得エラー]"
            attachments_text = ""
            attachment_names = []
            has_files = False
            attachments_collection = None
            is_target = False

            if item.Class == 43:
                skip_reason = None
                try:
                    mail_item = item
                    # 1. 軽い処理 (変更なし)
                    try:
                        mail_entry_id = str(getattr(mail_item, 'EntryID', 'UNKNOWN_ID'))
                    except Exception:
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
                    except Exception:
                        received_time = datetime.datetime.now().replace(tzinfo=None)
                        
                    # 2. スキップ判定 (変更なし)
                    # (★注意: 重複の原因はここのロジックです。read_mode="all" だと is_already_in_db を無視します)
                    if read_mode == "unprocessed" and is_already_in_db:
                        skip_reason = "DB or Session processed"
                    elif start_date_dt is not None and received_time < start_date_dt:
                        skip_reason = f"期間外"
                    
                    if skip_reason:
                        pass 
                    
                    else:
                        # 3. 重い処理 (変更なし)
                        try: subject = str(getattr(mail_item, 'Subject', ''))
                        except Exception: subject = "[件名取得エラー]"
                        try: body = str(getattr(mail_item, 'Body', ''))
                        except Exception: body = "[本文取得エラー]"

                        try:
                            if mail_item and hasattr(mail_item, 'Attachments'):
                                attachments_collection = mail_item.Attachments
                            if attachments_collection.Count > 0:
                                has_files = True
                                attachment_names = [att.FileName for att in attachments_collection if hasattr(att, 'FileName')]
                        except Exception as attach_err:
                            print(f"  -> 警告(has_files): 添付情報取得エラー: {attach_err}")

                        # --- ▼▼▼ ★★★ ハイブリッド版 添付ファイルループ (変更なし) ★★★ ▼▼▼
                        if has_files and attachments_collection:
                            if not is_already_in_db: 
                                print(f"DEBUG: Mail ID {mail_entry_id[:15]}... に {len(attachment_names)} 個の添付ファイルを発見。")
                                try:
                                    for attachment in attachments_collection:
                                        if not hasattr(attachment, 'FileName'): continue
                                        
                                        filename = attachment.FileName
                                        extracted_content = ""
                                        
                                        # 1. まず「高速な方法 (メモリ)」を試す
                                        try:
                                            print(f"DEBUG:  -> '{filename}' を [高速(メモリ)] モードで試行...")
                                            content_bytes = attachment.PropertyAccessor.GetProperty(PR_ATTACH_DATA_BIN)
                                            content_bytes_base64 = base64.b64encode(content_bytes).decode('utf-8')
                                            
                                            extracted_content = get_attachment_text(
                                                filename=filename,
                                                temp_file_path=None, 
                                                content_bytes_base64=content_bytes_base64
                                            )
                                            print(f"DEBUG:  -> '{filename}' [高速(メモリ)] 成功。 (抽出文字数: {len(extracted_content)})")
                                        
                                        # 2. 高速な方法が失敗 (埋め込み画像など) したら
                                        except (pythoncom.com_error, pywintypes.com_error) as com_err: # 👈 ★ pywintypes もキャッチ
                                            # 'パラメーターが間違っています' エラーをキャッチ
                                            if com_err.hresult == -2147024809:
                                                print(f"DEBUG:  -> '{filename}' [高速(メモリ)] 失敗。 [安全(ディスク)] モードにフォールバックします。")
                                                
                                                # 3. 「安全な方法 (一時ファイル)」にフォールバック
                                                temp_file_path = os.path.join(temp_dir, f"{uuid.uuid4().hex}_{filename}")
                                                try:
                                                    attachment.SaveAsFile(temp_file_path)
                                                    extracted_content = get_attachment_text(
                                                        filename=filename,
                                                        temp_file_path=temp_file_path, # 👈 ★ 一時ファイルパスを渡す
                                                        content_bytes_base64=None
                                                    )
                                                    print(f"DEBUG:  -> '{filename}' [安全(ディスク)] 成功。 (抽出文字数: {len(extracted_content)})")
                                                except Exception as save_err:
                                                    print(f"ERROR: (Attach Save/Slow): フォールバック失敗 (File: {filename}): {save_err}")
                                                finally:
                                                    if os.path.exists(temp_file_path):
                                                        try: os.remove(temp_file_path)
                                                        except OSError: pass
                                            else:
                                                # その他のCOMエラー
                                                print(f"ERROR: (Attach Read/Fast): 予期せぬCOMエラー (File: {filename}): {com_err}")
                                                raise 
                                        
                                        # 4. 抽出結果を結合
                                        attachments_text += f"\n--- FILE: {filename} ---\n{str(extracted_content)}\n"
                                
                                except Exception as loop_err:
                                    print(f"警告: 添付ファイルループ処理エラー (ID: {mail_entry_id}): {loop_err}")
                                    attachments_text += "\n--- ERROR during attachment loop ---\n"
                                attachments_text = attachments_text.strip()
                        # --- ▲▲▲ ★★★ ハイブリッド処理ここまで ★★★ ▲▲▲

                            # (★ 変更) --------------------------------------------
                            # --- ▼▼▼ ユーザー要望のロジックに変更 ▼▼▼ ---
                            
                        body_subject_search_text = subject + " " + body
                        condition_1_has_attachment = has_files 
                        condition_2_has_must_keyword = any(re.search(kw, body_subject_search_text, re.IGNORECASE) for kw in MUST_INCLUDE_KEYWORDS)
                        condition_3_is_excluded = any(re.search(kw, body_subject_search_text, re.IGNORECASE) for kw in EXCLUDE_KEYWORDS)
                        is_target = (condition_1_has_attachment or condition_2_has_must_keyword) and not condition_3_is_excluded
                            
                            # --- 判定ロジック終了 ---


                        if not is_already_in_db: 
                            # カウント処理 (要望4のため)
                            
                            # (1) 添付ファイルあり
                            if condition_1_has_attachment:
                                count_cond1_attach += 1
                                
                            # (2) 添付ファイルがなく、MUSTキーワードあり (★デバッグ要件変更箇所)
                            elif (not condition_1_has_attachment) and condition_2_has_must_keyword:
                                count_cond2_must += 1


                            if is_target:
                                # 抽出対象
                                record = {
                                    'EntryID': mail_entry_id, '件名': subject, '受信日時': received_time,
                                    '本文(テキスト形式)': body, '本文(ファイル含む)': attachments_text,
                                    'Attachments': ", ".join(attachment_names),
                                }
                                data_records_batch.append(record)
                                ids_processed_this_session.add(mail_entry_id)
                                count_extracted_total += 1

                            elif (condition_1_has_attachment or condition_2_has_must_keyword) and condition_3_is_excluded:
                                # 抽出対象だったが、除外された
                                skip_ids_batch.append(mail_entry_id)
                                ids_processed_this_session.add(mail_entry_id)
                                count_cond3_excluded += 1
                                
                            else:
                                # そもそも抽出対象外 (1にも2にも該当しない)
                                skip_ids_batch.append(mail_entry_id)
                                ids_processed_this_session.add(mail_entry_id)
                            # --- ▲▲▲ ユーザー要望のロジックに変更 ▲▲▲ ---
                            # -----------------------------------------------------

                except (pythoncom.com_error, pywintypes.com_error, AttributeError, Exception) as item_ex: # 👈 ★ pywintypes もキャッチ
                    current_id = mail_entry_id if mail_entry_id != 'UNKNOWN' else getattr(item, 'EntryID', 'ID取得失敗')
                    print(f"警告(Item Loop): 処理中にエラー (ID: {current_id}): {item_ex}\n{traceback.format_exc(limit=1)}")
                finally:
                    pass
            
            else:
                pass 

            try:
                item = items.GetNext() 
            except (pythoncom.com_error, pywintypes.com_error, Exception) as next_err: # 👈 ★ pywintypes もキャッチ
                print(f"警告: GetNext() でエラー。ループ中断。エラー: {next_err}")
                break 

    except (pythoncom.com_error, pywintypes.com_error) as com_outer_err: # 👈 ★ pywintypes もキャッチ
        raise RuntimeError(f"Outlook操作エラー (COM): {com_outer_err}\n{traceback.format_exc()}")
    except Exception as e:
        raise RuntimeError(f"Outlook操作エラー: {e}\n{traceback.format_exc()}")
    finally:
        # --- ループ終了後、残りのバッチを yield する ---
        if (data_records_batch or skip_ids_batch) and (not stop_flag or not stop_flag.is_set()):
        
            # (★ 追加) --------------------------------------------
            # ▼▼▼ 最後のコンソール出力 ▼▼▼
            print(f"--- 最終バッチ処理結果 ---")
            print(f"  (1) 添付ファイルあり (除外前): {count_cond1_attach} 件")
            print(f"  (2) Mustキーワードあり(添付なし): {count_cond2_must} 件") # (★デバッグ表示名を変更)
            print(f"  (3) 除外キーワードにより除外: {count_cond3_excluded} 件")
            print(f"  [+] 最終バッチ抽出件数: {count_extracted_total} 件")
            print(f"--------------------------")
            # ▲▲▲ 最後のコンソール出力 ▲▲▲
            # -----------------------------------------------------

            print(f"INFO: 最後のバッチ (抽出:{len(data_records_batch)}件, スキップ:{len(skip_ids_batch)}件) を返します。")
            yield {
                "data_batch": pd.DataFrame(data_records_batch),
                "skip_ids": skip_ids_batch
            }
            data_records_batch.clear()
            skip_ids_batch.clear()

        # ( ... 一時フォルダクリーンアップ (変更なし) ... )
        if os.path.exists(temp_dir):
            try:
                for f in os.listdir(temp_dir):
                    try:
                        os.remove(os.path.join(temp_dir, f))
                    except PermissionError:
                        print(f"WARN: 一時ファイル {f} の削除に失敗しました (使用中)。")
                if not os.listdir(temp_dir): 
                    os.rmdir(temp_dir)
            except OSError as oe: 
                print(f"警告: 一時フォルダクリーンアップ失敗: {oe}")
        pass