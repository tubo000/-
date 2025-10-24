# email_processor.py (未処理メール抽出の最終版)

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

# 外部定数と関数の依存関係を想定
try:
    from config import MUST_INCLUDE_KEYWORDS, EXCLUDE_KEYWORDS, SCRIPT_DIR
    def get_outlook_folder(outlook_ns, account_name, folder_path):
        """Outlookフォルダオブジェクトを取得する（実装は outlook_api.py にあるものと仮定）"""
        try:
            return outlook_ns.Folders[account_name].Folders[folder_path]
        except Exception:
            return None
    
    # 添付ファイルの中身（テキスト）を取得する関数 (file_processor.py に実装されているはず)
    try:
        from file_processor import get_attachment_text
    except ImportError:
        def get_attachment_text(*args, **kwargs): return "ATTACHMENT_CONTENT_FILE_IO_FAILED" 
    
except ImportError:
    MUST_INCLUDE_KEYWORDS = [r'スキルシート']
    EXCLUDE_KEYWORDS = []
    SCRIPT_DIR = os.getcwd() 
    def get_outlook_folder(*args, **kwargs): return None
    def get_attachment_text(*args, **kwargs): return "ATTACHMENT_CONTENT_FILE_IO_FAILED" 
    
OUTPUT_FILENAME = 'extracted_skills_result.xlsx' 
PROCESSED_CATEGORY_NAME = "スキルシート処理済" 

# ----------------------------------------------------------------------
# 💡 ヘルパー関数: 過去の本文データ復元
# ----------------------------------------------------------------------

def _load_previous_attachment_content() -> Dict[str, str]:
    """
    過去の抽出結果ファイルから EntryID と 本文(ファイル含む) を読み込み、
    本文復元用の辞書を返す。
    """
    script_dir_path = SCRIPT_DIR if 'SCRIPT_DIR' in globals() else os.getcwd()
    output_file_path = os.path.join(os.path.abspath(script_dir_path), OUTPUT_FILENAME)
    
    if os.path.exists(output_file_path):
        try:
            df_prev = pd.read_excel(output_file_path, usecols=['メールURL', '本文(ファイル含む)'], dtype=str)
            
            df_prev['EntryID'] = df_prev['メールURL'].str.replace('outlook:', '', regex=False).str.strip()
            df_prev.set_index('EntryID', inplace=True)
            
            return df_prev['本文(ファイル含む)'].dropna().to_dict()
        
        except Exception as e:
            print(f"警告: 過去ファイルからの本文復元に失敗しました。エラー: {e}")
            return {}
    return {}

# ----------------------------------------------------------------------
# 💡 共通機能: メールアイテムの処理済みマーク
# ----------------------------------------------------------------------

def mark_email_as_processed(mail_item):
    """指定されたメールアイテムに「処理済み」カテゴリを設定する。"""
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
# 💡 共通機能: 処理済みカテゴリの解除
# ----------------------------------------------------------------------

# email_processor.py 内の remove_processed_category 関数

def remove_processed_category(target_email: str, folder_path: str, days_ago: int = None) -> int:
    """
    指定されたフォルダのメールから 'PROCESSED_CATEGORY_NAME' マークを解除する。
    days_agoが指定された場合、その期間【より古い】メールのみを対象とする。
    解除件数を戻り値として返す。
    """
    reset_count = 0
    try:
        pythoncom.CoInitialize()
        
        # win32com.client を使用して Outlook オブジェクトを取得 (維持)
        try:
            outlook = win32.GetActiveObject("Outlook.Application")
        except:
            outlook = win32.Dispatch("Outlook.Application")

        namespace = outlook.GetNamespace("MAPI")
        folder = get_outlook_folder(namespace, target_email, folder_path)
        
        if folder is None:
            raise RuntimeError(f"指定されたフォルダパス '{folder_path}' が見つかりませんでした。")

        items = folder.Items
        
        # フィルタリングクエリの構築
        filter_query_list = []
        
        # 1. カテゴリによるフィルタリング (PROCESSED_CATEGORY_NAMEが付いたアイテム)
        category_filter_query = f"[Categories] = '{PROCESSED_CATEGORY_NAME}'"
        filter_query_list.append(category_filter_query)
        
        # 2. 期間指定フィルタ (days_ago が指定された場合)
        if days_ago is not None:
            # 📌 修正1: 削除機能と統一し、【N日より古い】メールを対象とする (<)
            start_date = (datetime.datetime.now() - timedelta(days=days_ago)).strftime('%m/%d/%Y %H:%M %p')
            filter_query_list.append(f"[ReceivedTime] < '{start_date}'") # 👈 比較演算子を < に変更

        query_string = " AND ".join(filter_query_list)
        items_to_reset = items.Restrict(query_string)
        
        # カテゴリを削除
        for item in items_to_reset:
            if item.Class == 43: # olMailItem
                current_categories = getattr(item, 'Categories', '')
                
                if PROCESSED_CATEGORY_NAME in current_categories:
                    # カテゴリを分割・削除し、再結合
                    categories_list = [c.strip() for c in current_categories.split(',') if c.strip() != PROCESSED_CATEGORY_NAME]
                    item.Categories = ", ".join(categories_list)
                    item.Save()
                    reset_count += 1
        
        return reset_count

    except Exception as e:
        raise RuntimeError(f"カテゴリマーク解除中にエラーが発生しました。詳細: {e}")
    finally:
        pythoncom.CoUninitialize()
# ----------------------------------------------------------------------
# 💡 未処理メールの件数をカウント (デバッグログ削除)
# ----------------------------------------------------------------------

def has_unprocessed_mail(folder_path: str, target_email: str) -> int:
    """
    指定されたフォルダに、処理済みカテゴリがなく、【かつキーワードに合致する】メールの件数をカウントする。
    """
    unprocessed_count = 0
    if not folder_path or not target_email: return 0
        
    outlook = None # COMオブジェクトを初期化
    
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
            
            for item in items:
                if item.Class == 43: # olMailItem
                    
                    subject = str(getattr(item, 'Subject', ''))
                    body = str(getattr(item, 'Body', ''))
                    categories = str(getattr(item, 'Categories', ''))
                    full_search_text = subject + " " + body 

                    # 1. 処理済みではないことをチェック
                    if PROCESSED_CATEGORY_NAME not in categories:
                        
                        # 2. キーワードに合致するかをチェック (ノイズ除外)
                        must_include = any(re.search(kw, full_search_text, re.IGNORECASE) for kw in MUST_INCLUDE_KEYWORDS)
                        
                        if must_include:
                            unprocessed_count += 1
                        
    except Exception as e:
        print(f"警告: 未処理メールチェック中にCOMエラー発生: {e}")
        unprocessed_count = 0
        
    finally:
        pythoncom.CoUninitialize()
        
    # 📌 修正: デバッグログの出力を削除
    
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
    
    # 処理済みメールの本文データを事前に読み込み
    previous_attachment_content = _load_previous_attachment_content()
    
    try:
        pythoncom.CoInitialize()
        # 📌 修正2: win32.client.Dispatch の誤りを win32.Dispatch に修正
        try:
            outlook_app = win32.GetActiveObject("Outlook.Application")
        except:
            outlook_app = win32.Dispatch("Outlook.Application")
        outlook_ns = outlook_app.GetNamespace("MAPI")
        target_folder = get_outlook_folder(outlook_ns, account_name, target_folder_path)
        
        if target_folder is None:
            raise RuntimeError(f"指定されたフォルダパス '{target_folder_path}' が見つかりませんでした。")

        items = target_folder.Items
        
        # 日付フィルタリングクエリの構築
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
                    mail_item = item
                    
                    is_processed = False
                    mail_entry_id = str(getattr(mail_item, 'EntryID', 'UNKNOWN')) 
                    
                    # 処理済みカテゴリチェック (is_processed を設定)
                    if hasattr(item, 'Categories'):
                        current_categories = str(getattr(item, 'Categories', ''))
                        if PROCESSED_CATEGORY_NAME in current_categories:
                            is_processed = True
                            
                    # 'unprocessed' モードの場合、処理済みはスキップ
                    if read_mode == "unprocessed" and is_processed:
                        continue 

                    # 📌 修正1: 件名と本文を str() に強制変換 (float/None エラー回避)
                    subject = str(getattr(mail_item, 'Subject', '')) 
                    body = str(getattr(mail_item, 'Body', ''))       
                    received_time = getattr(mail_item, 'ReceivedTime', datetime.datetime.now())
                    
                    if received_time is not None and received_time.tzinfo is not None:
                        received_time = received_time.replace(tzinfo=None)
                    elif received_time is None:
                        received_time = datetime.datetime.now().replace(tzinfo=None)
                    
                    attachments_text = ""
                    attachment_names = []
                    
                    # ----------------------------------------------------
                    # 📌 修正1: 添付ファイル処理を try/except で完全にラップ
                    # ----------------------------------------------------
                    has_files = hasattr(mail_item, 'Attachments') and mail_item.Attachments.Count > 0
                    
                    if has_files:
                        attachment_names = [att.FileName for att in mail_item.Attachments]
                        
                        if is_processed and mail_entry_id in previous_attachment_content:
                            attachments_text = str(previous_attachment_content.get(mail_entry_id, "")) 
                            
                        else:
                            # 未処理の場合、ファイルI/Oを実行しテキストを抽出
                            for attachment in mail_item.Attachments:
                                
                                safe_filename = re.sub(r'[\\/:*?"<>|]', '_', attachment.FileName)
                                temp_file_path = os.path.join(temp_dir, f"{uuid.uuid4().hex}_{safe_filename}")
                                
                                # 📌 修正2: 個別添付ファイルの処理を try/except で保護
                                try:
                                    attachment.SaveAsFile(temp_file_path)
                                    extracted_content = get_attachment_text(temp_file_path, attachment.FileName)
                                    attachments_text += f"\n--- FILE: {attachment.FileName} ---\n{str(extracted_content)}\n"
                                except Exception as file_ex:
                                    # 抽出失敗ログを本文に残す
                                    attachments_text += f"\n--- ERROR reading {attachment.FileName}: {file_ex} ---\n"
                                finally:
                                    # 一時ファイルを確実に削除
                                    if os.path.exists(temp_file_path):
                                        os.remove(temp_file_path)
                            
                            attachments_text = attachments_text.strip()
                    
                    full_search_text = str(subject) + " " + str(body) + " " + str(attachments_text)
                    
                    # キーワードフィルタリング (MUST/EXCLUDE)
                    must_include = any(re.search(kw, full_search_text, re.IGNORECASE) for kw in MUST_INCLUDE_KEYWORDS)
                    is_excluded = any(re.search(kw, full_search_text, re.IGNORECASE) for kw in EXCLUDE_KEYWORDS)
                    
                    
                    # 抽出対象として残す条件を調整 (キーワードチェック)
                    if is_processed and not must_include:
                         continue

                    # 📌 修正1: 除外キーワードに合致した場合の処理を変更
                    if is_excluded:
                         if not is_processed:
                             # 未処理の除外対象メールであれば、マークを付けてスキップ
                             mark_email_as_processed(mail_item)
                         continue
                        
                    if not must_include and not is_processed:
                         # 未処理だが、キーワードに該当しないメールは抽出せず、マークだけ付けてスキップ
                         mark_email_as_processed(mail_item)
                         continue  
                    # レコードの準備 (抽出結果を DataFrame に追加)
                    record = {
                        'EntryID': mail_entry_id,
                        '件名': subject,
                        '受信日時': received_time, 
                        '本文(テキスト形式)': body, 
                        '本文(ファイル含む)': attachments_text, # 復元または新規抽出された本文
                        'Attachments': ", ".join(attachment_names),
                    }
                    data_records.append(record)
                    
                    # 抽出が成功し、かつ未処理の場合のみ、メールを「処理済み」としてマーク
                    if not is_processed:
                        mark_email_as_processed(mail_item) 
                        
                except Exception as item_ex:
                    # 📌 修正3: エラー発生時でも、未処理メールならマークを付ける（固定化回避）
                    print(f"警告: メールアイテムの処理中にエラーが発生しました (EntryID: {mail_entry_id}). スキップします。エラー: {item_ex}")
                    if not is_processed:
                         # 抽出に失敗したが、次回以降カウントされないようマークを付ける
                         mark_email_as_processed(mail_item) 
                    continue # 次のアイテムへ

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