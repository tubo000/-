# email_processor.py (全ての修正を統合した最終版)

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
INITIALS_REGEX = r'(\b[A-Z]{2}\b|\b[A-Z]\s*.\s*[A-Z]\b|名前\([A-Z]{2}\))'

# --- インポート処理 ---

# 1. get_attachment_text のデフォルト（代替）定義
def get_attachment_text(*args, **kwargs):
    print("警告: file_processor.py から get_attachment_text を読み込めませんでした。")
    return "ATTACHMENT_CONTENT_IMPORT_FAILED"

# 2. get_outlook_folder のデフォルト（代替）定義
def get_outlook_folder(outlook_ns, account_name, folder_path):
     print(f"警告: config.py から get_outlook_folder を読み込めませんでした。デフォルト処理を使用します。")
     try:
          return outlook_ns.Folders[account_name].Folders[folder_path]
     except Exception:
          print(f"エラー: デフォルトのフォルダ取得も失敗しました: {account_name}/{folder_path}")
          return None

# 3. config.py から設定値と関数を読み込む
try:
    from config import MUST_INCLUDE_KEYWORDS, EXCLUDE_KEYWORDS, SCRIPT_DIR, OUTPUT_CSV_FILE as OUTPUT_FILENAME
    try:
        from config import get_outlook_folder as real_get_outlook_folder
        get_outlook_folder = real_get_outlook_folder # インポート成功、デフォルト関数を上書き
        print("INFO: config.py から get_outlook_folder を読み込みました。")
    except ImportError:
        print("警告: config.py に get_outlook_folder が定義されていません。デフォルト処理を使用します。")
    print("INFO: config.py から設定値を読み込みました。")
except ImportError:
    print("警告: config.py が見つからないかインポートできませんでした。デフォルト設定を使用します。")
    MUST_INCLUDE_KEYWORDS = [r'スキルシート']
    EXCLUDE_KEYWORDS = [r'案\s*件\s*名',r'案\s*件\s*番\s*号',r'案\s*件:',r'案\s*件：',r'【案\s*件】',r'概\s*要',r'必\s*須']
    SCRIPT_DIR = os.getcwd()
    OUTPUT_FILENAME = 'output_extraction.xlsx'

# 4. file_processor.py から関数を読み込む
try:
    from file_processor import get_attachment_text as real_get_attachment_text
    get_attachment_text = real_get_attachment_text
    print("INFO: file_processor.py から get_attachment_text を読み込みました。")
except ImportError:
    print("警告: file_processor.py が見つからないか 'get_attachment_text' が含まれていません。")
except Exception as e:
    print(f"エラー: file_processor.py のインポート中にエラー: {e}")

# --- 修正ここまで ---
DATABASE_NAME = 'extraction_cache.db'
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
            pass # マーク付け失敗は許容
        return True
    return False

# ----------------------------------------------------------------------
# 💡 処理済みカテゴリの解除 (COM初期化削除 + デバッグログ)
# ----------------------------------------------------------------------
def remove_processed_category(target_email: str, folder_path: str, days_ago: int = None) -> int:
    reset_count = 0
    start_date_dt = None

    print("\n--- DEBUG: remove_processed_category 開始 ---")
    print(f"DEBUG: target_email='{target_email}', folder_path='{folder_path}', days_ago={days_ago}")

    if days_ago is not None:
        try:
             days_ago = int(days_ago)
             if days_ago <= 0: # 0以下は全期間 (None) 扱い
                  print(f"INFO(remove_category): days_ago が {days_ago} のため、全期間 (None) として扱います。")
                  days_ago = None
             else:
                  start_date_dt = (datetime.datetime.now() - timedelta(days=days_ago))
                  print(f"DEBUG: 計算された cutoff datetime (これより古いメールのカテゴリを解除): {start_date_dt}")
        except (ValueError, TypeError):
             print(f"警告(remove_category): days_ago '{days_ago}' が不正なため、None として扱います。")
             days_ago = None 
             start_date_dt = None # days_ago が None なので start_date_dt も None

    try:
        # --- 📌 CoInitialize() 削除 ---
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
        try:
             initial_item_count = items.Count
             print(f"DEBUG: Restrict前のアイテム数: {initial_item_count}")
        except Exception as count_err:
             print(f"警告(remove_category): Restrict前のアイテム数取得失敗: {count_err}")

        filter_query_list = []
        # 日付フィルタ (days_ago が 1 以上の場合のみ)
        if start_date_dt is not None:
            start_date_str = start_date_dt.strftime('%Y/%m/%d %H:%M') # 時刻も比較
            filter_query_list.append(f"[ReceivedTime] < '{start_date_str}'")
            print(f"DEBUG: Restrict 条件 (日付): [ReceivedTime] < '{start_date_str}'")
        else:
             print("DEBUG: 日付指定なし (全期間のカテゴリ付きメールを対象)")

        query_string = " AND ".join(filter_query_list)
        items_to_reset = items

        if query_string:
            try:
                items_to_reset = items.Restrict(query_string)
                print(f"DEBUG: Restrict 実行成功。")
                try:
                     restricted_count = items_to_reset.Count
                     print(f"DEBUG: Restrict後のアイテム数: {restricted_count}")
                except Exception as count_err:
                     print(f"警告(remove_category): Restrict後のアイテム数取得失敗: {count_err}")
            except Exception as restrict_error:
                print(f"警告: カテゴリ解除のRestrict(日付)に失敗: {restrict_error}。全件チェックにフォールバックします。")

        try:
            items_to_reset.Sort("[ReceivedTime]", True)
            print(f"DEBUG: アイテムのソート成功 (降順)。")
        except Exception as sort_err:
             print(f"警告(remove_category): ソート失敗: {sort_err}")

        print(f"DEBUG: カテゴリ解除ループ開始...")
        item_counter = 0
        item = items_to_reset.GetFirst()
        
        # ▼▼▼【修正】無限ループバグ修正 (GetNextをループ最後に移動) ▼▼▼
        while item:
            item_counter += 1
            mail_entry_id_debug = getattr(item, 'EntryID', 'UNKNOWN_ID')

            if item.Class == 43:
                try:
                    current_categories = getattr(item, 'Categories', '')
                    if PROCESSED_CATEGORY_NAME in current_categories:
                        print(f"DEBUG: ★ カテゴリ発見！解除試行 (ID: {mail_entry_id_debug})")
                        
                        # (日付Restrict失敗時のフォールバックチェック)
                        is_target_date = True # デフォルトは解除対象
                        if start_date_dt is not None: # 日付指定がある場合のみチェック
                            received_time = getattr(item, 'ReceivedTime', datetime.datetime.now())
                            if received_time.tzinfo is not None:
                                received_time = received_time.replace(tzinfo=None)
                            if received_time >= start_date_dt: # 基準日時より新しいものは対象外
                                is_target_date = False
                                print(f"DEBUG:   -> 日付フォールバック: 対象外 (受信日時 {received_time})")

                        if is_target_date: # 日付条件も満たす場合のみ解除
                            try:
                                categories_list = [c.strip() for c in current_categories.split(',') if c.strip() != PROCESSED_CATEGORY_NAME]
                                new_categories = ", ".join(categories_list)
                                item.Categories = new_categories
                                item.Save()
                                reset_count += 1
                                print(f"DEBUG:   -> カテゴリ解除成功！ (累計: {reset_count})")
                            except Exception as save_err:
                                 print(f"エラー(remove_category): カテゴリ保存/Save失敗 (ID: {mail_entry_id_debug}): {save_err}")
                        
                except pythoncom.com_error as com_err:
                     print(f"警告(remove_category Loop): アイテム処理中 COMエラー (ID: {mail_entry_id_debug}): {com_err.hresult if hasattr(com_err, 'hresult') else 'N/A'}")
                except Exception as e:
                    print(f"警告(remove_category Loop): アイテム処理中エラー (ID: {mail_entry_id_debug}): {e}")
            
            # --- 次のアイテムへ (ループの最後) ---
            try:
                item = items_to_reset.GetNext()
            except:
                print(f"DEBUG: GetNext() でエラーまたは終端。ループ終了。")
                break
        # ▲▲▲【修正】ここまで ▲▲▲
                
        print(f"DEBUG: カテゴリ解除ループ終了。総ループ回数: {item_counter}")

    except Exception as e:
        print(f"エラー(remove_category Main): カテゴリマーク解除中に予期せぬエラー: {e}\n{traceback.format_exc(limit=2)}")
        reset_count = -1
    finally:
        pass # --- 📌 CoUninitialize() 削除 ---

    print(f"DEBUG: remove_processed_category 終了。解除された件数: {reset_count}")
    print("---------------------------------------")
    return reset_count


# ----------------------------------------------------------------------
# 💡 未処理メールの件数をカウント (COM初期化削除 + デバッグログ)
# ----------------------------------------------------------------------
def has_unprocessed_mail(folder_path: str, target_email: str, days_to_check: int = None) -> int:
    unprocessed_count = 0
    if not folder_path or not target_email: return 0

    valid_days_to_check = None
    cutoff_date_dt = None 

    print(f"\n--- DEBUG (has_unprocessed_mail 開始) ---")
    print(f"DEBUG: 受け取った days_to_check: {days_to_check} (型: {type(days_to_check)})")

    if days_to_check is not None:
        try:
            days_to_check_int = int(days_to_check)
            if days_to_check_int >= 0:
                valid_days_to_check = days_to_check_int
                if valid_days_to_check == 0:
                    cutoff_date_dt = datetime.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
                    print(f"DEBUG: 計算された cutoff_date_dt (今日0時): {cutoff_date_dt}")
                else:
                    cutoff_date_dt = datetime.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0) - timedelta(days=valid_days_to_check)
                    print(f"DEBUG: 計算された cutoff_date_dt ({valid_days_to_check}日前): {cutoff_date_dt}")
            else:
                 print(f"警告(has_unprocessed): 不正日数 {days_to_check}, 全期間チェック")
        except (ValueError, TypeError):
             print(f"警告(has_unprocessed): 日数 {days_to_check} が不正, 全期間チェック")

    try:
        # --- 📌 CoInitialize() 削除 ---
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
                    print(f"DEBUG: Outlook Restrict 実行。絞り込み条件: >= '{cutoff_date_str}'")
                except Exception as restrict_error:
                    print(f"警告: has_unprocessed_mailの日付絞り込み失敗。全件スキャン: {restrict_error}")
                    items = folder.Items
            else:
                 print("DEBUG: has_unprocessed_mail: 日付指定なし。全期間をチェックします。")

            try: items.Sort("[ReceivedTime]", True)
            except Exception as sort_error: print(f"警告(has_unprocessed): Sort失敗: {sort_error}")

            item = items.GetFirst()
            
            # ▼▼▼【修正】無限ループバグ修正 (GetNextをループ最後に移動) ▼▼▼
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
                                 print(f"警告(has_unprocessed): 添付情報/名前チェックエラー (ID: {mail_entry_id_debug}): {attach_err}")

                            subject = str(getattr(item, 'Subject', ''))
                            body = str(getattr(item, 'Body', ''))
                            full_search_text = subject + " " + body
                            must_include = any(re.search(kw, full_search_text, re.IGNORECASE) for kw in MUST_INCLUDE_KEYWORDS)
                            
                            # ▼▼▼【注意】昨日のコードでは本文イニシャルもカウント対象 ▼▼▼
                            has_initials_in_text = re.search(INITIALS_REGEX, full_search_text) 

                            is_target_for_count = must_include or has_initials_in_text or (has_files and has_initials_in_filename)
                            # ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

                            if is_target_for_count:
                                unprocessed_count += 1
                                
                 except pythoncom.com_error as com_err:
                      print(f"警告(has_unprocessed Loop): COMエラー (ID: {mail_entry_id_debug}): {com_err.hresult if hasattr(com_err, 'hresult') else 'N/A'}")
                 except Exception as e:
                     print(f"警告(has_unprocessed Loop): アイテム処理エラー (ID: {mail_entry_id_debug}): {e}")

                 try:
                     item = items.GetNext()
                 except:
                     break
            # ▲▲▲【修正】ここまで ▲▲▲

    except Exception as e:
        print(f"警告(has_unprocessed Main): チェック処理エラー: {e}")
        unprocessed_count = 0
    finally:
        pass # --- 📌 CoUninitialize() 削除 ---

    print(f"DEBUG: has_unprocessed_mail 最終カウント: {unprocessed_count}")
    print(f"--- DEBUG (has_unprocessed_mail 終了) ---")
    return unprocessed_count

# ----------------------------------------------------------------------
# 💡 メイン抽出関数: Outlookからメールを取得 (全機能統合)
# ----------------------------------------------------------------------
def get_mail_data_from_outlook_in_memory(target_folder_path: str, account_name: str, read_mode: str = "all", days_ago: int = None) -> pd.DataFrame:
    """
    Outlookからメールデータを抽出する。read_modeに基づいてフィルタリングを行う。
    days_ago=0 の場合は今日受信したメールのみを対象とする。
    """
    data_records = []
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
             print(f"警告: 不正日数 '{days_ago}', 全期間対象. Error: {e}")
             days_ago = None
             start_date_dt = None
             log_period_message = "全期間 (入力不正)"

    print(f"INFO: Outlookメール読み込み開始 (対象期間: {log_period_message})")

    existing_ids_set = set()
    db_path = os.path.join(os.path.abspath(SCRIPT_DIR), DATABASE_NAME)
    if os.path.exists(db_path):
        try:
            conn_check = sqlite3.connect(db_path)
            existing_ids_set = set(pd.read_sql_query("SELECT EntryID FROM emails", conn_check)['EntryID'].tolist())
            conn_check.close()
            print(f"INFO: 既存DBから {len(existing_ids_set)} 件のEntryIDを読み込みました。")
        except Exception as e:
            print(f"警告: 既存DBのEntryID読み込み失敗: {e}。全件新規として扱います。")
            existing_ids_set = set()
    
    new_record_counter = 0
    max_new_records = 100

    try:
        # --- 📌 CoInitialize() 削除 ---
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

        processed_item_count = 0
        item = items.GetFirst()

        # ▼▼▼【修正】無限ループバグ修正 (if...else 構造) ▼▼▼
        while item:
            if new_record_counter >= max_new_records:
                print(f"INFO: 新規レコードが {max_new_records} 件に達したため、処理を中断します。")
                break
            
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
                skip_reason = None # スキップ理由
                try:
                    mail_item = item
                    try:
                        mail_entry_id = str(getattr(mail_item, 'EntryID', 'UNKNOWN_ID'))
                    except Exception as id_err:
                         print(f"  -> 警告: EntryID取得エラー: {id_err}")
                         mail_entry_id = f"ERROR_ID_{uuid.uuid4().hex}"
                         is_already_in_db = False
                    else:
                         is_already_in_db = mail_entry_id in existing_ids_set
                    
                    try:
                        subject = str(getattr(mail_item, 'Subject', ''))
                    except Exception as subj_err:
                        print(f"  -> 警告: 件名取得エラー (ID: {mail_entry_id}): {subj_err}")
                        subject = "[件名取得エラー]"

                    print(f"\n[{processed_item_count}] 処理中: {subject[:50]}... (ID: ...{mail_entry_id[-20:]})")

                    try:
                        current_categories = getattr(mail_item, 'Categories', '')
                        if PROCESSED_CATEGORY_NAME in current_categories:
                            is_processed = True
                    except Exception as cat_err:
                         print(f"  -> 警告: カテゴリ取得エラー (ID: {mail_entry_id}): {cat_err}")

                    try:
                        received_time_check = getattr(mail_item, 'ReceivedTime', datetime.datetime.now())
                        if received_time_check.tzinfo is not None:
                            received_time_check = received_time_check.replace(tzinfo=None)
                        received_time = received_time_check
                    except Exception as rt_err:
                         print(f"  -> 警告: 受信日時取得エラー (ID: {mail_entry_id}): {rt_err}")
                         received_time = datetime.datetime.now().replace(tzinfo=None)

                    try:
                        body = str(getattr(mail_item, 'Body', ''))
                    except Exception as body_err:
                        print(f"  -> 警告: 本文取得エラー (ID: {mail_entry_id}): {body_err}")
                        body = "[本文取得エラー]"

                    # --- スキップ判定 ---
                    if read_mode == "unprocessed" and is_processed:
                        skip_reason = "Outlook処理済み"
                    elif read_mode == "unprocessed" and is_already_in_db:
                         skip_reason = "DB登録済み"
                    elif start_date_dt is not None and received_time < start_date_dt:
                         skip_reason = f"期間外 ({received_time.strftime('%Y-%m-%d %H:%M')})"
                    
                    if skip_reason:
                        print(f"  -> スキップ: {skip_reason}")
                        # ★ continue は使わない
                    
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
                             print(f"  -> スキップ: 除外キーワード '{matched_exclude_kw}' にマッチ")
                             if not is_processed: mark_email_as_processed(mail_item)
                        
                        else:
                            is_target = has_must_include_keyword or (has_files and has_initials_in_filename)
                            # ... (ログ表示) ...
                            print(f"  -> 判定: is_target={is_target} (...)") 

                            if is_target:
                                if not is_already_in_db:
                                    print(f"  -> ★★★ 新規抽出対象としてレコード追加 ★★★")
                                    record = {
                                        'EntryID': mail_entry_id, '件名': subject, '受信日時': received_time,
                                        '本文(テキスト形式)': body, '本文(ファイル含む)': attachments_text,
                                        'Attachments': ", ".join(attachment_names),
                                    }
                                    data_records.append(record)
                                    extraction_succeeded = True
                                    new_record_counter += 1
                            elif not is_target:
                                print(f"  -> スキップ: 抽出対象外")
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
                                  print(f"  -> INFO: 処理済みマークを付与")
                              except Exception as mark_e:
                                   print(f"  -> 警告: 抽出成功後のマーク付け失敗: {mark_e}")
            
            else:
                 print(f"[{processed_item_count}] スキップ: メールアイテムではありません (Class: {item.Class})")

            # --- ループの最後に必ず次のアイテムを取得 ---
            try:
                item = items.GetNext() 
            except (pythoncom.com_error, Exception) as next_err:
                 print(f"警告: GetNext() でエラー。ループ中断。エラー: {next_err}")
                 break 
        # ▲▲▲【修正】ここまで ▲▲▲

    except pythoncom.com_error as com_outer_err:
         raise RuntimeError(f"Outlook操作エラー (COM): {com_outer_err}\n{traceback.format_exc()}")
    except Exception as e:
        raise RuntimeError(f"Outlook操作エラー: {e}\n{traceback.format_exc()}")
    finally:
        if os.path.exists(temp_dir):
             try:
                 if not os.listdir(temp_dir): os.rmdir(temp_dir)
             except OSError as oe: print(f"警告: 一時フォルダクリーンアップ失敗: {oe}")
        # --- 📌 CoUninitialize() 削除 ---
        pass 

    print(f"INFO: Outlookメール読み込みループ終了。処理アイテム数: {processed_item_count}, 新規抽出件数: {len(data_records)}")
    df = pd.DataFrame(data_records)
    if not df.empty:
        str_cols = [col for col in df.columns if col != '受信日時']
        df[str_cols] = df[str_cols].fillna('N/A').astype(str)
        df['受信日時'] = pd.to_datetime(df['受信日時'], errors='coerce')
        if '受信日時' in df.columns:
            df = df.sort_values(by='受信日時', ascending=False, na_position='last').reset_index(drop=True)

    return df

# ----------------------------------------------------------------------
# 💡 外部公開関数
# ----------------------------------------------------------------------
def run_email_extraction(target_email: str, read_mode: str = "all", days_ago: int = None):
    pass

def delete_old_emails_core(target_email: str, folder_path: str, days_ago: int) -> int:
    pass