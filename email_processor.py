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
# --- 📌 修正ここから ---

# 1. get_attachment_text のデフォルト（代替）定義
def get_attachment_text(*args, **kwargs):
    print("警告: file_processor.py から get_attachment_text を読み込めませんでした。")
    return "ATTACHMENT_CONTENT_IMPORT_FAILED"

# 2. get_outlook_folder のデフォルト（代替）定義
def get_outlook_folder(outlook_ns, account_name, folder_path):
     print(f"警告: config.py から get_outlook_folder を読み込めませんでした。デフォルト処理を使用します。")
     # デフォルトの挙動（もしあれば記述、なければ None を返す）
     # 例: 標準的なフォルダ構造を探すなど。ここでは None を返す
     try:
          # デフォルトのフォルダを探す試み (例)
          return outlook_ns.Folders[account_name].Folders[folder_path]
     except Exception:
          print(f"エラー: デフォルトのフォルダ取得も失敗しました: {account_name}/{folder_path}")
          return None # 失敗したら None

# 3. config.py から設定値と関数を読み込む
try:
    from config import MUST_INCLUDE_KEYWORDS, EXCLUDE_KEYWORDS, SCRIPT_DIR, OUTPUT_CSV_FILE as OUTPUT_FILENAME
    
    # ▼▼▼ 修正点 ▼▼▼
    # get_outlook_folder を config から明示的にインポート
    try:
        from config import get_outlook_folder as real_get_outlook_folder
        get_outlook_folder = real_get_outlook_folder # インポート成功、デフォルト関数を上書き
        print("INFO: config.py から get_outlook_folder を読み込みました。")
    except ImportError:
        print("警告: config.py に get_outlook_folder が定義されていません。デフォルト処理を使用します。")
        # デフォルト関数がそのまま使われる
    # ▲▲▲▲▲▲▲▲▲▲
        
    print("INFO: config.py から設定値を読み込みました。")

except ImportError:
    # config.py 自体が見つからない場合
    print("警告: config.py が見つからないかインポートできませんでした。デフォルト設定を使用します。")
    MUST_INCLUDE_KEYWORDS = [r'スキルシート']
    EXCLUDE_KEYWORDS = [r'案\s*件\s*名',r'案\s*件\s*番\s*号',r'案\s*件:',r'案\s*件：',r'【案\s*件】',r'概\s*要',r'必\s*須']
    SCRIPT_DIR = os.getcwd()
    OUTPUT_FILENAME = 'output_extraction.xlsx'
    # get_outlook_folder は上で定義したデフォルトが使われる

# 4. file_processor.py から関数を読み込む (変更なし)
try:
    from file_processor import get_attachment_text as real_get_attachment_text
    get_attachment_text = real_get_attachment_text
    print("INFO: file_processor.py から get_attachment_text を読み込みました。")
except ImportError:
    print("警告: file_processor.py が見つからないか 'get_attachment_text' が含まれていません。")
except Exception as e:
    print(f"エラー: file_processor.py のインポート中にエラー: {e}")

# --- 📌 修正ここまで ---
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

# email_processor.py (L125 付近の remove_processed_category 関数のみ差し替え)

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
            print(f"警告: カテゴリ解除のRestrict(日付)に失敗しました: {restrict_error}")
            items_to_reset = items

        # 降順 (新しい順) に並び替え
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

# email_processor.py の has_unprocessed_mail 関数 (修正版)

def has_unprocessed_mail(folder_path: str, target_email: str, days_to_check: int = None) -> int:
    """
    指定されたフォルダの未処理メール件数をカウントする。
    days_to_check が指定されていれば、その日数で絞り込む。
    """
    unprocessed_count = 0
    if not folder_path or not target_email: return 0

    valid_days_to_check = None
    if days_to_check is not None:
        try:
            days_to_check_int = int(days_to_check)
            if days_to_check_int >= 0:
                valid_days_to_check = days_to_check_int
            else:
                 print(f"警告(has_unprocessed): 不正日数 {days_to_check}, 全期間チェック")
        except (ValueError, TypeError):
             print(f"警告(has_unprocessed): 日数 {days_to_check} が不正, 全期間チェック")

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

            # --- 日付絞り込み (変更なし) ---
            if valid_days_to_check is not None:
                try:
                    if valid_days_to_check == 0:
                         cutoff_date_dt = datetime.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
                    else:
                         cutoff_date_dt = datetime.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0) - timedelta(days=valid_days_to_check)
                    cutoff_date_str = cutoff_date_dt.strftime('%Y/%m/%d %H:%M')
                    date_filter = f"[ReceivedTime] >= '{cutoff_date_str}'"
                    items = items.Restrict(date_filter)
                    # print(f"DEBUG(has_unprocessed): 過去{valid_days_to_check}日間に絞り込み。")
                except Exception as restrict_error:
                    print(f"警告(has_unprocessed): Restrict失敗。全件スキャン: {restrict_error}")
                    items = folder.Items
            # else:
            #      print("DEBUG(has_unprocessed): 日付指定なし。全期間チェック。")
            
            # --- ソート (変更なし) ---
            try: items.Sort("[ReceivedTime]", True)
            except Exception as sort_error: print(f"警告(has_unprocessed): Sort失敗: {sort_error}")

            # --- アイテムループ ---
            item = items.GetFirst()
            while item:
                mail_entry_id_debug = 'UNKNOWN_ID' # エラー表示用
                try:
                    mail_entry_id_debug = getattr(item, 'EntryID', 'UNKNOWN_ID') # 早めにID取得試行
                    if item.Class == 43:
                        categories = str(getattr(item, 'Categories', ''))
                        if PROCESSED_CATEGORY_NAME not in categories:
                            
                            # ▼▼▼【ここを修正】添付ファイル関連処理を try...except で囲む ▼▼▼
                            has_files = False
                            has_initials_in_filename = False
                            try:
                                if item and hasattr(item, 'Attachments'):
                                     attachments_collection = item.Attachments
                                     attachment_count = attachments_collection.Count
                                     if attachment_count > 0:
                                         has_files = True
                                         # ファイル名リストを取得してイニシャルチェック
                                         attachment_names = [att.FileName for att in attachments_collection if hasattr(att, 'FileName')]
                                         all_filenames_text = " ".join(attachment_names)
                                         if re.search(INITIALS_REGEX, all_filenames_text):
                                             has_initials_in_filename = True
                            except (pythoncom.com_error, AttributeError, Exception) as attach_err:
                                 # エラーが出ても警告表示にとどめ、has_files/has_initials は False のまま
                                 print(f"警告(has_unprocessed): 添付ファイル情報/名前チェック中にエラー (ID: {mail_entry_id_debug}): {attach_err}")
                            # ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

                            # --- 必須キーワード・本文イニシャルチェック (変更なし) ---
                            subject = str(getattr(item, 'Subject', ''))
                            body = str(getattr(item, 'Body', ''))
                            full_search_text = subject + " " + body
                            must_include = any(re.search(kw, full_search_text, re.IGNORECASE) for kw in MUST_INCLUDE_KEYWORDS)
                            has_initials_in_text = re.search(INITIALS_REGEX, full_search_text) 

                            # --- 抽出対象判定 (has_files, has_initials_in_filename は安全に取得済み) ---
                            is_target_for_count = must_include or has_initials_in_text or (has_files and has_initials_in_filename)

                            if is_target_for_count:
                                unprocessed_count += 1
                                
                except pythoncom.com_error as com_err:
                     print(f"警告(has_unprocessed Loop): COMエラー (ID: {mail_entry_id_debug}): {com_err.hresult if hasattr(com_err, 'hresult') else 'N/A'}")
                except Exception as e:
                    print(f"警告(has_unprocessed Loop): アイテム処理中にエラー (ID: {mail_entry_id_debug}): {e}")
                
                # --- 次のアイテムへ (変更なし) ---
                try:
                    item = items.GetNext()
                except: 
                    break

    except Exception as e:
        print(f"警告(has_unprocessed Main): チェック処理中にエラー発生: {e}")
        unprocessed_count = 0 # エラー時は0を返す
    finally:
        pythoncom.CoUninitialize()
        
    return unprocessed_count
# ----------------------------------------------------------------------
# 💡 メイン抽出関数: Outlookからメールを取得 (高精度ロジック + 高速Restrict)
# ----------------------------------------------------------------------

# email_processor.py の get_mail_data_from_outlook_in_memory 関数

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
    if days_ago is not None:
        try:
             days_ago = int(days_ago) 
             if days_ago < 0: raise ValueError("日数は0以上の整数である必要があります")
             if days_ago == 0:
                 today_date = datetime.date.today()
                 start_date_dt = datetime.datetime.combine(today_date, datetime.time.min) 
             else:
                 start_date_dt = datetime.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0) - timedelta(days=days_ago)
        except ValueError as e:
             print(f"警告: 不正な日数 '{days_ago}' が指定されました。全期間を対象とします。エラー: {e}")
             days_ago = None 
             start_date_dt = None
             
    try:
        pythoncom.CoInitialize()
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
            is_processed = False
            mail_entry_id = 'UNKNOWN'
            mail_item = None # ループごとに初期化

            if item.Class == 43: # olMailItem
                extraction_succeeded = False
                try:
                    mail_item = item # mail_item に現在のアイテムを代入
                    mail_entry_id = str(getattr(mail_item, 'EntryID', 'UNKNOWN')) # 早めにID取得を試みる

                    # --- カテゴリチェック ---
                    current_categories = getattr(mail_item, 'Categories', '') # エラーが出にくい getattr を使用
                    if PROCESSED_CATEGORY_NAME in current_categories:
                        is_processed = True

                    # --- モード/日付チェック ---
                    if read_mode == "unprocessed" and is_processed:
                        # item = items.GetNext() # ループ末尾で実行
                        continue 
                    if start_date_dt is not None: 
                         received_time_check = getattr(mail_item, 'ReceivedTime', datetime.datetime.now())
                         if received_time_check.tzinfo is not None:
                             received_time_check = received_time_check.replace(tzinfo=None)
                         if received_time_check < start_date_dt: 
                             # item = items.GetNext() # ループ末尾で実行
                             continue

                    # --- 基本情報取得 ---
                    subject = str(getattr(mail_item, 'Subject', ''))
                    body = str(getattr(mail_item, 'Body', ''))
                    received_time = getattr(mail_item, 'ReceivedTime', datetime.datetime.now()) # 再取得 (整形のため)
                    if received_time is not None and received_time.tzinfo is not None:
                        received_time = received_time.replace(tzinfo=None)
                    elif received_time is None:
                        received_time = datetime.datetime.now().replace(tzinfo=None)

                    # --- 添付ファイルチェック (has_files) ---
                    attachments_text = ""
                    attachment_names = []
                    has_files = False # デフォルト False
                    
                    # ▼▼▼【ここを再修正】hasattr も try の中に入れる ▼▼▼
                    try:
                        # mail_item が有効か再確認
                        if mail_item and hasattr(mail_item, 'Attachments'):
                             # Attachments および Count へのアクセスを try 内で行う
                             attachments_collection = mail_item.Attachments # 変数に入れる
                             attachment_count = attachments_collection.Count 
                             if attachment_count > 0:
                                 has_files = True
                                 # 念のためファイル名リストもここで取得しておく
                                 attachment_names = [att.FileName for att in attachments_collection if hasattr(att, 'FileName')]
                                 
                    except pythoncom.com_error as com_err:
                         print(f"警告(has_files): 添付情報取得中にCOMエラー (ID: {mail_entry_id}): {com_err.hresult if hasattr(com_err, 'hresult') else 'N/A'}")
                    except AttributeError as ae:
                         print(f"警告(has_files): 添付情報取得中にAttributeError (ID: {mail_entry_id}): {ae}")
                    except Exception as e:
                         print(f"警告(has_files): 添付情報取得中に予期せぬエラー (ID: {mail_entry_id}): {e}")
                    # ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

                    # --- 添付ファイル内容抽出 (has_files が True の場合のみ) ---
                    if has_files:
                        if is_processed and mail_entry_id in previous_attachment_content:
                            attachments_text = str(previous_attachment_content.get(mail_entry_id, ""))
                        else:
                            # attachment_names は上で取得済み
                            for attachment in attachments_collection: # 上で取得したコレクションを使用
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
                            attachments_text = attachments_text.strip()

                    # --- キーワード・イニシャル・除外チェック ---
                    body_subject_search_text = str(subject) + " " + str(body)
                    search_text_for_keywords = body_subject_search_text + " " + attachments_text
                    has_must_include_keyword = any(re.search(kw, search_text_for_keywords, re.IGNORECASE) for kw in MUST_INCLUDE_KEYWORDS)
                    
                    has_initials_in_filename = False
                    if has_files: # attachment_names は上で取得済み
                        all_filenames_text = " ".join(attachment_names)
                        if re.search(INITIALS_REGEX, all_filenames_text):
                             has_initials_in_filename = True

                    full_search_text = body_subject_search_text + " " + attachments_text
                    is_excluded = False
                    for kw in EXCLUDE_KEYWORDS:
                        if re.search(kw, full_search_text, re.IGNORECASE):
                            is_excluded = True
                            break
                            
                    if is_excluded:
                         if not is_processed: mark_email_as_processed(mail_item)
                         continue 

                    # --- 抽出対象判定 ---
                    is_target = has_must_include_keyword or (has_files and has_initials_in_filename)

                    # --- 抽出 & マーク付け ---
                    if is_target:
                        record = {
                            'EntryID': mail_entry_id, '件名': subject, '受信日時': received_time,
                            '本文(テキスト形式)': body, '本文(ファイル含む)': attachments_text,
                            'Attachments': ", ".join(attachment_names),
                        }
                        data_records.append(record)
                        extraction_succeeded = True 
                        if not is_processed:
                            # 後続のマーク付け処理へ
                             pass 
                             
                    elif not is_target: 
                        if not is_processed: mark_email_as_processed(mail_item) 
                        continue 

                # --- アイテム処理中の包括的なエラーハンドリング ---
                except pythoncom.com_error as com_err:
                    print(f"警告(Item Loop): COMエラー (ID: {mail_entry_id}): {com_err.hresult if hasattr(com_err, 'hresult') else 'N/A'}")
                    # このアイテムはスキップ (GetNextはfinallyの外)
                except AttributeError as ae:
                     # 特定の属性アクセスでエラーが出た場合
                     print(f"警告(Item Loop): AttributeError (ID: {mail_entry_id}): {ae}")
                except Exception as item_ex:
                    print(f"警告(Item Loop): 予期せぬエラー (ID: {mail_entry_id}): {item_ex}\n{traceback.format_exc(limit=1)}")
                    # エラーが出ても、未処理ならマーク付けを試みる (mail_itemがNoneでない場合)
                    if mail_item and not is_processed:
                        try: mark_email_as_processed(mail_item)
                        except Exception as mark_e: print(f"  警告: エラー後のマーク付け失敗: {mark_e}")
                # ★ try ブロックの最後 (エラーがあってもなくても実行されるべき処理)
                finally:
                      # 抽出成功 かつ 未処理だった場合 -> マーク付け
                      if extraction_succeeded and not is_processed:
                          try:
                              mark_email_as_processed(mail_item)
                          except Exception as mark_e:
                               print(f"警告: 抽出成功後のマーク付け失敗 (ID: {mail_entry_id}): {mark_e}")
                               
            # --- ループの最後に必ず次のアイテムを取得 ---
            try:
                item = items.GetNext() 
            except pythoncom.com_error as next_err:
                 print(f"警告: GetNext() COMエラー。ループ中断。Code: {next_err.hresult if hasattr(next_err, 'hresult') else 'N/A'}")
                 break 
            except Exception as next_ex:
                 print(f"警告: GetNext() 予期せぬエラー。ループ中断。エラー: {next_ex}")
                 break 

    except pythoncom.com_error as com_outer_err:
         raise RuntimeError(f"Outlook操作エラー (COM): {com_outer_err}\n{traceback.format_exc()}")
    except Exception as e:
        raise RuntimeError(f"Outlook操作エラー: {e}\n{traceback.format_exc()}")
    finally:
        if os.path.exists(temp_dir):
             try:
                 if not os.listdir(temp_dir): os.rmdir(temp_dir)
             except OSError as oe: print(f"警告: 一時フォルダクリーンアップ失敗: {oe}")
        pythoncom.CoUninitialize()

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