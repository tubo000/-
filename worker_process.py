# worker_process.py
# (重い処理（抽出・削除）を専門に行う、2号機（別プロセス）)
# (v5e の省略部分をすべて展開した「完全版」)

import sys
import os
import pandas as pd
import pythoncom
import win32com.client as win32
import sqlite3
import traceback
import datetime
from datetime import timedelta
import re # 👈 ★ reorder_output_dataframe で使用

# (main_application.pyから、必要な関数を丸ごと移動)
from config import DATABASE_NAME, MASTER_COLUMNS, SCRIPT_DIR
from email_processor import get_mail_data_from_outlook_in_memory
from extraction_core import extract_skills_data

# ======================================================================
# ★★★ 移植 1: reorder_output_dataframe (main_application.py から移植) ★★★
# ======================================================================

def reorder_output_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """
    (main_application.py [v34] L:182 から移植)
    DataFrameの列順を、定義済みの主要列が先頭に来るように並び替える。
    """
    fixed_leading_cols = [
        'メールURL', '受信日時', '件名', '名前', '信頼度スコア', 
        '本文(テキスト形式)', '本文(ファイル含む)', 'Attachments'
    ]
    # 渡されたDFに存在する列のみを fixed_leading_cols の順序で選別
    fixed_leading_cols = [col for col in fixed_leading_cols if col in df.columns]
    
    # fixed_leading_cols 以外の列を、元の順序を保ったまま取得
    remaining_cols = [col for col in df.columns.tolist() if col not in fixed_leading_cols]
    
    # 最終的な列順序で DataFrame を再インデックス
    return df.reindex(columns=fixed_leading_cols + remaining_cols, fill_value='N/A')

# ======================================================================
# ★★★ 移植 2: delete_processed_records (main_application.py から移植) ★★★
# ======================================================================

def delete_processed_records(days_ago: int, db_path: str, table_name: str) -> str:
    """
    (main_application.py [v34] L:271 から移植)
    指定された日数に基づき、指定されたテーブル内の古いレコードを削除する。
    """
    try:
        days_ago = int(days_ago)
        if days_ago < 0:
             raise ValueError("日数は0以上の整数で指定してください。")
    except ValueError:
        return f"エラー({table_name}): 日数設定が不正です (0以上の整数で指定)。"

    today = datetime.date.today()
    
    if table_name == "skipped_ids" and days_ago != 0:
        return f"INFO(skipped_ids): スキップIDは日付指定削除の対象外です。"

    # ▼▼▼ 修正 1 (メッセージをテーブル名に応じて変更) ▼▼▼
    if table_name == "emails":
        if days_ago == 0:
            where_clause = "" 
            target_message = "抽出済みレコード (すべて)"
        else:
            cutoff_date = today - timedelta(days=(days_ago - 1)) 
            cutoff_datetime = datetime.datetime.combine(cutoff_date, datetime.time.min) 
            cutoff_str = cutoff_datetime.strftime('%Y-%m-%d %H:%M:%S')
            where_clause = f"WHERE \"受信日時\" < '{cutoff_str}'"
            target_message = f"抽出済みレコード ('{days_ago}日前以前')"
    
    elif table_name == "skipped_ids":
        # (このロジックに到達する時点で days_ago == 0 が保証されている)
        where_clause = "" 
        target_message = "スキップID (すべて)"
        
    else: # 念のためフォールバック
        if days_ago == 0:
            where_clause = "" 
            target_message = f"テーブル '{table_name}' のすべての取り込み記録"
        else:
            cutoff_date = today - timedelta(days=(days_ago - 1)) 
            cutoff_datetime = datetime.datetime.combine(cutoff_date, datetime.time.min) 
            cutoff_str = cutoff_datetime.strftime('%Y-%m-%d %H:%M:%S')
            where_clause = f"WHERE \"受信日時\" < '{cutoff_str}'"
            target_message = f"テーブル '{table_name}' の '{days_ago}日前以前' の記録"
    # ▲▲▲ 修正 1 ▲▲▲

    deleted_count = 0
    if not os.path.exists(db_path):
        return f"INFO({table_name}): データベースファイルが見つかりません。スキップします。"

    conn = None
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        count_sql = f"SELECT COUNT(*) FROM {table_name} {where_clause}"
        cursor.execute(count_sql)
        deleted_count = cursor.fetchone()[0]

        if deleted_count > 0:
            delete_sql = f"DELETE FROM {table_name} {where_clause}"
            cursor.execute(delete_sql)
            conn.commit() 
            # ▼▼▼ 修正 1 (メッセージ文言の変更) ▼▼▼
            return f"{target_message} {deleted_count}件 を削除しました。"            
        else:
            return f"{target_message} は見つかりませんでした。削除は行われませんでした。"
            # ▲▲▲ 修正 1 ▲▲▲
    except sqlite3.Error as e: 
        if conn: conn.rollback() 
        if "no such table" in str(e): 
             return f"INFO({table_name}): テーブルが見つかりません。スキップします。"
        return f"エラー({table_name}): DBファイルの処理中にエラーが発生しました ({e})"
    except Exception as e: 
         if conn: conn.rollback()
         return f"エラー({table_name}): 予期せぬエラーが発生しました ({e})"
    finally:
        if conn: conn.close()

# ======================================================================
# ★★★ 移植 3: actual_run_extraction_logic (main_application.py から移植) ★★★
# ======================================================================

def actual_run_extraction_logic(main_elements, target_email, folder_path, read_mode, read_days, gui_queue, stop_flag):
    """ 
    (main_application.py [v5] から移植)
    (★ 引数に gui_queue と stop_flag を追加)
    """
    
    try:
        pythoncom.CoInitialize()
    except Exception:
        pass 
        
    total_new_records_saved = 0 
    total_items_skipped = 0     
    
    try:
        days_ago = None
        if read_days.strip():
            try:
                days_ago = int(read_days)
                if days_ago < 0: raise ValueError("日数は0以上")
            except ValueError:
                gui_queue.put("ERROR:期間指定は 0以上の整数で指定してください。\n(空欄の場合は全期間)")
                return 

        if days_ago == 0: mode_text = "未処理 (今日のみ)"
        elif days_ago is not None and days_ago > 0 : mode_text = f"未処理 (過去{days_ago}日)"
        else: mode_text = "未処理 (全期間)"
            
        gui_queue.put(f"STATUS: {target_email} アカウントからメール取得中 ({mode_text})...")

        # ★ 修正: main_elements ではなく、引数の stop_flag を渡す
        main_elements_for_processor = {"stop_extraction_flag": stop_flag}
        
        df_mail_data_generator = get_mail_data_from_outlook_in_memory(
            folder_path, 
            target_email, 
            read_mode=read_mode, 
            days_ago=days_ago,
            main_elements=main_elements_for_processor # 👈 ★ 修正
        )
        
        gui_queue.put("STATUS: 抽出コアロジック実行中...")
        db_path = os.path.abspath(DATABASE_NAME) 
        
        batch_number = 0
        for batch_data in df_mail_data_generator:
            # (中断フラグのチェックは email_processor 側 で行われている)
            
            batch_number += 300
            df_mail_data_batch = batch_data.get("data_batch", pd.DataFrame())
            skip_ids_batch = batch_data.get("skip_ids", [])
            
            if df_mail_data_batch.empty and not skip_ids_batch:
                continue 

            if df_mail_data_batch.empty and not skip_ids_batch:
                continue 

            # ▼▼▼ 修正 2 (ステータスメッセージの変更) ▼▼▼
            gui_queue.put(f"STATUS: {batch_number}件目抽出中")
            # ▲▲▲ 修正 2 ▲▲▲
            
            df_extracted = pd.DataFrame()
            if not df_mail_data_batch.empty:
                df_extracted = extract_skills_data(df_mail_data_batch)

            conn = None
            
            try:
                conn = sqlite3.connect(db_path)
                try:
                    cursor = conn.cursor()
                    cursor.execute("PRAGMA table_info(emails);")
                    current_columns = {info[1] for info in cursor.fetchall()} 
                    for column_name in MASTER_COLUMNS:
                        if column_name not in current_columns and column_name != 'EntryID':
                            conn.execute(f"ALTER TABLE emails ADD COLUMN \"{column_name}\" TEXT;")
                    conn.commit()
                except Exception:
                     pass
                
                if not df_extracted.empty:
                    df_output = df_extracted.copy()
                    date_key_df = df_mail_data_batch[['EntryID', '受信日時']].copy()
                    if '受信日時' in df_output.columns:
                        df_output.drop(columns=['受信日時'], inplace=True, errors='ignore')
                    df_output = pd.merge(df_output, date_key_df, on='EntryID', how='left')
                    if 'EntryID' in df_output.columns and 'メールURL' not in df_output.columns:
                         df_output.insert(0, 'メールURL', df_output.apply(lambda row: f"outlook:{row['EntryID']}", axis=1))
                    df_output = reorder_output_dataframe(df_output) # 👈 このファイルの関数を呼ぶ
                    final_drop_list = ['宛先メール', '本文(抽出元結合)'] 
                    final_drop_list = [col for col in df_output.columns if col in final_drop_list]
                    df_output = df_output.drop(columns=final_drop_list, errors='ignore')
                    if 'EntryID' not in df_output.columns:
                         raise KeyError("抽出結果バッチに EntryID が含まれていません。")
                    entry_ids_in_this_batch = df_output['EntryID'].tolist()
                    df_output.set_index('EntryID', inplace=True)
                    
                    try:
                        existing_ids_set = set(pd.read_sql_query("SELECT EntryID FROM emails", conn)['EntryID'].tolist())
                    except pd.io.sql.DatabaseError:
                        existing_ids_set = set() 
                        
                    new_ids = [eid for eid in entry_ids_in_this_batch if eid not in existing_ids_set]
                    df_new = df_output.loc[new_ids]
                    
                    if not df_new.empty:
                        df_new.to_sql('emails', conn, if_exists='append', index=True) 
                        newly_saved_count = len(df_new)
                        total_new_records_saved += newly_saved_count 
                        # ▼▼▼ 修正 3 (デバッグ用コンソール出力を無効化) ▼▼▼
                        # print(f"INFO: {newly_saved_count} 件の新規レコードをDBに追記しました。(累計: {total_new_records_saved} 件)")                            gui_queue.put("ENABLE_SEARCH_BUTTON")

                if skip_ids_batch:
                    try:
                        existing_skip_ids = set(pd.read_sql_query("SELECT EntryID FROM skipped_ids", conn)['EntryID'].tolist())
                    except pd.io.sql.DatabaseError:
                        existing_skip_ids = set()
                        conn.execute("CREATE TABLE IF NOT EXISTS skipped_ids (EntryID TEXT PRIMARY KEY)")
                    
                    unique_ids_in_batch = set(skip_ids_batch)
                    new_skip_ids = [eid for eid in unique_ids_in_batch if eid not in existing_skip_ids] 
                    
                    # [修正後]
                    if new_skip_ids:
                            df_skip = pd.DataFrame(new_skip_ids, columns=['EntryID'])
                            df_skip.set_index('EntryID', inplace=True)
                            df_skip.to_sql('skipped_ids', conn, if_exists='append', index=True)
                            total_items_skipped += len(new_skip_ids)
                            # ▼▼▼ 修正 3 (デバッグ用コンソール出力を無効化) ▼▼▼
                            # print(f"INFO: {len(new_skip_ids)} 件のスキップIDをDBに追記しました。(累計: {total_items_skipped} 件)")
                            gui_queue.put("ENABLE_SEARCH_BUTTON")
            except Exception as e:
                print(f"❌ データベース書き込み中にエラー発生: {e}")
                gui_queue.put(f"DB_ERROR:データベースへの書き込み中にエラーが発生しました。\n詳細: {e}")
            finally:
                if conn: conn.close()
            
        if total_new_records_saved == 0 and total_items_skipped == 0:
            gui_queue.put("EXTRACTION_NO_ITEMS_FOUND")
            return 

        final_message = f"抽出処理が正常に完了しました。\n"
        if total_new_records_saved > 0:
             final_message += f"合計 {total_new_records_saved} 件の新規レコードが '{DATABASE_NAME}' に保存されました。\n"
        if total_items_skipped > 0:
             final_message += f"合計 {total_items_skipped} 件のメールが「スキップ対象」としてDBに記録されました。\n"
        
        gui_queue.put(f"EXTRACTION_COMPLETE:{total_new_records_saved}:{final_message}") 
        gui_queue.put(f"STATUS: 処理完了。{total_new_records_saved} 件保存済み。")
        
    except Exception as e:
        error_message_for_user = f"抽出処理中に予期せぬエラーが発生しました。\n詳細: {e}"
        gui_queue.put(f"EXTRACTION_ERROR:{error_message_for_user}")
        traceback.print_exc()
        
    finally:
        gui_queue.put("EXTRACTION_COMPLETE_ENABLE_BUTTON") 
        pythoncom.CoUninitialize()

def actual_run_file_deletion_logic(main_elements, gui_queue):
    """
    (main_application.py [v5] から移植)
    v_single_thread の堅牢なエラーハンドリングと結果レポートロジックを
    v_multi_process の gui_queue 構造に統合した修正版。
    """
    
    try:
        # 1. 入力値の取得
        days_input = main_elements["delete_days_entry"] 
        db_path = os.path.abspath(DATABASE_NAME) 

        # 2. 入力値の検証
        try:
            days_ago = int(days_input)
            if days_ago < 0: 
                raise ValueError("日数は0以上の整数を指定してください。")
        except ValueError as e:
            # ▼▼▼ 修正 (v3で適用済み) ▼▼▼
            # 検証エラーをGUIスレッドに通知
            gui_queue.put(f"MSGBOX_ERROR:入力エラー:削除日数の入力が不正です: {e}\n(0以上の整数で指定)")
            # ▲▲▲ 修正 ▲▲▲
            gui_queue.put("STATUS: 削除失敗 (入力不正)。")
            return 

        # 3. 実行中ステータスをGUIスレッドに通知
        gui_queue.put(f"STATUS: DBレコード削除試行中...")

        # 4. 削除処理
        db_exists = os.path.exists(db_path) 
        delete_result_message = "" 
        delete_result_message_skipped = "" 
        db_had_error = False 

        if db_exists:
            try:
                # 抽出済み(emails)テーブルから削除
                # (引数は検証済みの 'days_ago' を使う)
                delete_result_message = delete_processed_records(days_ago, db_path, "emails")
                # スキップ済み(skipped_ids)テーブルから削除
                delete_result_message_skipped = delete_processed_records(days_ago, db_path, "skipped_ids")
                
                if "エラー:" in delete_result_message or "エラー:" in delete_result_message_skipped:
                    db_had_error = True 
            
            except NameError:
                delete_result_message = "内部エラー: レコード削除関数(delete_processed_records)が見つかりません。"
                db_had_error = True
            except Exception as db_del_err:
                delete_result_message = f"DBレコード削除中に予期せぬエラーが発生しました。\n{db_del_err}" 
                db_had_error = True
        else:
            delete_result_message = f"INFO: データベースファイル '{os.path.basename(db_path)}' が見つかりませんでした。DBレコード削除はスキップされました。"

        # 5. 結果レポートの準備
        
        # 5a. 2つのテーブルの結果メッセージを統合
        final_msg = delete_result_message + "\n" + delete_result_message_skipped.replace("INFO: ", "")
        final_msg = final_msg.strip() 
        
        # 5b. 最終的なメッセージタイトルとステータスを決定
        msg_title = "処理完了"
        msg_icon_type = "MSGBOX_INFO" # GUIキュー用の命令
        final_status_text = "状態: 削除処理完了。"
        
        if db_had_error:
            msg_title = "処理完了 (DB削除エラー)"
            msg_icon_type = "MSGBOX_WARNING"
            final_status_text = "状態: DB削除エラー。"
        elif "INFO:" in delete_result_message: # "DBファイルなし" の場合
            msg_title = "処理スキップ"
            msg_icon_type = "MSGBOX_INFO"
            final_status_text = "状態: DBファイルなし。"
            
        # 6. ★★★ ここが修正点 ★★★
        # 準備した「単一の」結果をGUIスレッドに通知
        gui_queue.put(f"{msg_icon_type}:{msg_title}:{final_msg}")
        gui_queue.put(f"STATUS:{final_status_text}")
        
        # ▼▼▼ 削除 ▼▼▼
        # 以下のブロックは「不要」かつ「間違い」
        # # 1. 'emails' テーブルから削除 (移植した delete_processed_records を使用)
        # result_emails = delete_processed_records(days_input, db_path, "emails")
        # # ... (以下4行すべて削除) ...
        # ▲▲▲ 削除 ▲▲▲
    
    except Exception as outer_err:
        # この関数自体の予期せぬエラー
        try:
            gui_queue.put("STATUS: 削除スレッドで重大なエラー。")
            gui_queue.put(f"MSGBOX_ERROR:重大なエラー:削除処理中に予期せぬエラーが発生しました。\n{outer_err}")
        except:
            pass 
        
    finally:
        # 処理が正常終了でもエラーでも、GUIのボタンを有効化するよう通知
        gui_queue.put("DELETION_COMPLETE_ENABLE_BUTTON")
        pass
