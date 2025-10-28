# email_processor.py (安定版 - Restrict構文エラー修正)

import pandas as pd
import win32com.client as win32 # 👈 win32 というエイリアスを使用
import pythoncom
import os
import datetime
import re
from datetime import timedelta
import sys
import uuid 
import traceback
from typing import Dict, Any, List
import sqlite3 # 📌 修正1: sqlite3 をインポート

# 外部定数と関数の依存関係を想定 (維持)
try:
    # 📌 修正1: config から OUTPUT_CSV_FILE を OUTPUT_FILENAME としてインポート
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
    
# 📌 修正2: 保存先を .db ファイルに変更
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
            # 必要なカラムだけを高速に読み込む
            df_prev = pd.read_sql_query("SELECT \"EntryID\", \"本文(ファイル含む)\" FROM emails", conn)
            conn.close()
            
            df_prev['EntryID'] = df_prev['EntryID'].str.replace('outlook:', '', regex=False).str.strip()
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
    # ... (変更なし) ...
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
# 💡 処理済みカテゴリの解除 (維持)
# ----------------------------------------------------------------------
# email_processor.py86行目の関数
def remove_processed_category(target_email: str, folder_path: str, days_ago: int = None) -> int:
    """
    指定された日数より古いメールのカテゴリを、その種類に関わらず全て解除するロジックに修正。
    解除基準は、days_ago日前の午前0時0分0秒より古いメールとする。
    """
    reset_count = 0
    query_string = None
   
    if days_ago is not None:
        # 基準日時を「days_ago日前の午前0時0分0秒」に固定する
        today_midnight = datetime.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        cutoff_datetime_naive = today_midnight - timedelta(days=days_ago)
       
        # 📌 修正1: Outlook検索クエリ用の文字列を作成 (ISO 8601形式が推奨)
        # 受信日時が基準日時より古いもののみを検索する
        # ここでは、日付検索の安定性を優先し、日付のみの比較を行う。
        # days_ago日前の午前0時より古い、つまり days_ago+1日前の日付より前のメール
       
        # 検索クエリは日付のみを使うため、最もシンプルな形式（days_ago日前の午前0時）を使う
        search_date_str = cutoff_datetime_naive.strftime('%m/%d/%Y %H:%M %p') # Outlookの標準形式に戻す
        query_string = f"[ReceivedTime] < '{search_date_str}'"
       
    else:
        # days_agoがNoneの場合はカテゴリ解除処理を行わない（必須でないため）
        return 0
 
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
       
        # 📌 修正2: 日付クエリのみでOutlookを検索（カテゴリーが付いていないものも含む）
        items_to_reset = items.Restrict(query_string)
       
        for item in items_to_reset:
            if item.Class == 43: # olMailItem
               
                # 受信日時が検索クエリでフィルタリングされているため、
                # ここではdays_agoによる追加の比較は不要。
               
                # 📌 修正3: カテゴリーが付いているかのチェックを削除し、全てクリアする
                current_categories = getattr(item, 'Categories', '')
               
                # カテゴリーが付いている場合のみ、処理を実行
                if current_categories:
                    item.Categories = "" # カテゴリーを全てクリアする
                    item.Save()
                    reset_count += 1
       
        return reset_count
 
    except Exception as e:
        # エラー処理は元のコードのまま維持
        raise RuntimeError(f"カテゴリマーク解除中にエラーが発生しました。詳細: {e}")
    finally:
        pythoncom.CoUninitialize()
 

# ----------------------------------------------------------------------
# 💡 未処理メールの件数をカウント (維持)
# ----------------------------------------------------------------------
def has_unprocessed_mail(folder_path: str, target_email: str) -> int:
    # ... (変更なし) ...
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
            try:
                items.Sort("[ReceivedTime]", False) 
            except Exception:
                pass
            
            item = items.GetFirst()
            while item:
                try:
                    if item.Class == 43:
                        subject = str(getattr(item, 'Subject', ''))
                        body = str(getattr(item, 'Body', ''))
                        categories = str(getattr(item, 'Categories', ''))
                        full_search_text = subject + " " + body 

                        if PROCESSED_CATEGORY_NAME not in categories:
                            must_include = any(re.search(kw, full_search_text, re.IGNORECASE) for kw in MUST_INCLUDE_KEYWORDS)
                            if must_include:
                                unprocessed_count += 1
                except Exception as e:
                    print(f"警告: アイテムスキャン中にCOMエラー: {e}")
                
                item = items.GetNext() # 次のアイテムへ
                        
    except Exception as e:
        print(f"警告: 未処理メールチェック中にCOMエラー発生: {e}")
        unprocessed_count = 0
        
    finally:
        pythoncom.CoUninitialize()
        
    return unprocessed_count


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
    
    previous_attachment_content = _load_previous_attachment_content()
    
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
        
        # ----------------------------------------------------
        # 📌 修正1: フィルタリングクエリの構築 (Jet構文のみ)
        # ----------------------------------------------------
        filter_query_list = []
        
        # 1. 期間指定フィルタ (Jet構文)
        # 「未処理」モードでも日付を適用するため、read_mode のチェックを削除
        if days_ago is not None:
            start_date = (datetime.datetime.now() - timedelta(days=days_ago)).strftime('%m/%d/%Y %H:%M %p')
            filter_query_list.append(f"[ReceivedTime] >= '{start_date}'")
 
        # 2. 📌 修正2: 「未処理」モードのDASL構文絞り込みを削除
        #    if read_mode == "unprocessed":
        #        category_filter = f"(\"urn:schemas...Keywords\" IS NULL OR ...)"
        #        filter_query_list.append(category_filter)
 
        if filter_query_list:
            query_string = " AND ".join(filter_query_list)
            try:
                # 📌 修正3: 絞り込みを実行（現在は日付フィルタのみ）
                items = items.Restrict(query_string)
            except Exception as restrict_error:
                # 構文が混在してエラーになった場合 (例: "条件が正しくありません")
                print(f"警告: Outlookの絞り込み(Restrict)に失敗しました: {restrict_error}")
                # 失敗した場合は、全件ループで処理 (低速だが安全)
                items = target_folder.Items

        # ----------------------------------------------------
        # 📌 修正4: 安定した GetFirst/GetNext ループ
        # ----------------------------------------------------
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
                    
                    # 処理済みカテゴリチェック (is_processed を設定)
                    if hasattr(item, 'Categories'):
                        current_categories = str(getattr(item, 'Categories', ''))
                        if PROCESSED_CATEGORY_NAME in current_categories:
                            is_processed = True
                            
                    # 📌 修正5: (安全策) Python側でのカテゴリチェック (必須)
                    if read_mode == "unprocessed" and is_processed:
                        item = items.GetNext()
                        continue
 
                    # 属性取得 (str() に強制変換でエラー回避)
                    subject = str(getattr(mail_item, 'Subject', ''))
                    body = str(getattr(mail_item, 'Body', '')) 
                    received_time = getattr(mail_item, 'ReceivedTime', datetime.datetime.now())
                    
                    if received_time is not None and received_time.tzinfo is not None:
                        received_time = received_time.replace(tzinfo=None)
                    elif received_time is None:
                        received_time = datetime.datetime.now().replace(tzinfo=None)
                    
                    attachments_text = ""
                    attachment_names = []
                    
                    # 添付ファイルの読み込みロジック (ファイルI/Oのスキップと復元)
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
                    
                    full_search_text = str(subject) + " " + str(body) + " " + str(attachments_text)
                    
                    # キーワードフィルタリング (MUST/EXCLUDE)
                    must_include = any(re.search(kw, full_search_text, re.IGNORECASE) for kw in MUST_INCLUDE_KEYWORDS)
                    is_excluded = any(re.search(kw, full_search_text, re.IGNORECASE) for kw in EXCLUDE_KEYWORDS)
                    
                    
                    # 抽出対象として残す条件を調整 (キーワードチェック)
                    if is_processed and not must_include:
                        item = items.GetNext()
                        continue
 
                    if is_excluded:
                        item = items.GetNext()
                        continue
                        
                    if not must_include and not is_processed:
                        # 未処理だが、キーワードに該当しないメールは抽出せず、マークだけ付けてスキップ
                        mark_email_as_processed(mail_item)
                        item = items.GetNext()
                        continue
                        
                    # レコードの準備
                    record = {
                        'EntryID': mail_entry_id,
                        '件名': subject,
                        '受信日時': received_time,
                        '本文(テキスト形式)': body,
                        '本文(ファイル含む)': attachments_text, # 復元または新規抽出された本文
                        'Attachments': ", ".join(attachment_names),
                    }
                    data_records.append(record)
                    
                    extraction_succeeded = True
 
                except Exception as item_ex:
                    print(f"警告: メールアイテムの処理中にエラーが発生しました (EntryID: {mail_entry_id}). スキップします。エラー: {item_ex}")
                    # 抽出が失敗した未処理メールは、次回以降のためにマークを付ける
                    if mail_item and not is_processed:
                        try:
                            mark_email_as_processed(mail_item)
                        except Exception as mark_e:
                            print(f"  警告: エラー発生後のマーク付けにも失敗しました: {mark_e}")
                    
                    item = items.GetNext()
                    continue
                
                # 正常な処理フローを通過し、かつ未処理だった場合のみマーク
                if extraction_succeeded and not is_processed:
                    mark_email_as_processed(mail_item)
            
            # ループの最後に次のアイテムを取得
            item = items.GetNext()
 
    except Exception as e:
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