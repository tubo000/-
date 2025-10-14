# main.py
# 目的: アプリケーションの実行フローを制御し、試験モードと本番テストモードの分岐を行う

import os
import sys
import pandas as pd
import win32com.client as win32 # Windows COM機能 (Outlook連携) のためのライブラリ
# 外部ファイルのインポート (システムのコア機能)
from config import INPUT_QUESTION_CSV, MASTER_ANSWERS_PATH, OUTPUT_EVAL_PATH, NUM_RECORDS
from data_generation import generate_raw_data, export_dataframes_to_tsv
from extraction_core import extract_skills_data
from evaluator_core import run_triple_csv_validation, get_question_data_from_csv
from email_processor import run_email_extraction, get_mail_data_from_outlook_in_memory, TARGET_FOLDER_PATH


# ----------------------------------------------------
# ユーティリティ関数: Outlook ID検索 (単体テスト用)
# ----------------------------------------------------

def open_outlook_email_by_id(entry_id: str):
    """
    指定された Entry ID (メール固有の識別子) を使用して、
    Outlookデスクトップアプリで該当メールを直接開く。
    """
    if not entry_id:
        print("エラー: Entry IDが指定されていません。", file=sys.stderr)
        return

    try:
        # Outlookアプリケーションへの接続 (既に起動していればそれを取得、なければ起動)
        try:
            # 起動中のOutlookインスタンスを取得
            outlook_app = win32.GetActiveObject("Outlook.Application")
        except:
            # 起動していない場合、新しいインスタンスを起動
            outlook_app = win32.Dispatch("Outlook.Application")
            
        # MAPI (Messaging API) ネームスペースを取得（Outlook内部へのアクセス経路）
        namespace = outlook_app.GetNamespace("MAPI")
        
        # Entry ID から特定のアイテム（メールオブジェクト）を検索して取得
        olItem = namespace.GetItemFromID(entry_id)
        
        if olItem:
            # アイテム（メール）を新しいウィンドウで表示
            olItem.Display()
            print(f"メールを正常に開きました: {getattr(olItem, 'Subject', '件名なし')}")
        else:
            print("エラー: 指定された Entry ID のメールが見つかりませんでした。", file=sys.stderr)
            
    except Exception as e:
        print(f"Outlook連携中にエラーが発生しました: {e}", file=sys.stderr)
        # ユーザーへの診断メッセージ
        print("Outlookが起動しているか、またはpywin32が正しくインストールされているか確認してください。", file=sys.stderr)


# ----------------------------------------------------
# ユーティリティ関数: インタラクティブテスト (Entry IDの単体確認用)
# ----------------------------------------------------

def interactive_id_search_test():
    """
    すべての処理が完了した後、ユーザーが手動で Entry ID を入力し、
    メールオープン機能をテストするためのプロンプトを出す。
    """
    
    print("\n\n==================================================")
    print("💌 Entry ID 検索機能テスト")
    
    # ユーザーにテスト実行の意思を確認
    test_choice = input("Entry ID 検索テストを実行しますか？ (y/n, nで終了): ").strip().lower()
    
    if test_choice == 'y':
        print("\n--------------------------------------------------")
        print("💡 テスト用の Entry ID をペーストしてEnterを押してください。")
        print(" (例: 00000000D30472EAB8069E4A8A...)")
        
        entry_id = input("Entry ID: ").strip()
        
        if entry_id:
            print(f"\n→ Entry ID [{entry_id[:10]}...] のメールを開きます...")
            # Outlookオープン関数を実行
            open_outlook_email_by_id(entry_id)
        else:
            print("Entry ID が入力されなかったため、テストをスキップします。")
    
    print("==================================================")


# ----------------------------------------------------
# ユーティリティ関数: 出力列の並び替え (論理的な並び替えを定義)
# ----------------------------------------------------

def reorder_output_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """
    出力データフレームの列順を調整し、特定の項目を左側に固定する。
    """
    # ... (既存の関数、省略 - 実際にはこの関数内で列の並び替えロジックが実行されます)
    fixed_leading_cols = []
    
    if 'メールURL' in df.columns:
        fixed_leading_cols.append('メールURL')
        
    fixed_leading_cols.extend(['件名', '名前'])
    
    all_cols = df.columns.tolist()
    
    remaining_cols = [col for col in all_cols if col not in fixed_leading_cols]
    
    final_col_order = fixed_leading_cols + remaining_cols
    
    df_reordered = df.reindex(columns=final_col_order, fill_value='N/A')
    
    return df_reordered


# ----------------------------------------------------
# 試験モード実行関数 (評価と柔軟なデータソース選択)
# ----------------------------------------------------

def main_process_exam_mode():
    """
    試験モードのメイン処理。ダミーデータまたはOutlookデータのどちらかを選択し、
    抽出を実行した後、マスターデータとの比較評価を行う。
    """
    print("★★ 統合スキルシート抽出・評価システム（試験モード）実行 ★★")
    
    # データソースの選択プロンプト
    print("\n--- 試験データの選択 ---")
    print(" [1] ダミーデータ生成 (デフォルト): 新規データを作成しCSVから読み込み")
    print(" [2] Outlookメールから読み込み: 実際のメールデータを使用")
    
    data_source_input = input("データソースを選択してください ([1]で実行): ").strip()
    df_mail_data = pd.DataFrame()

    if not data_source_input or data_source_input == '1':
        # オプション1: ダミーデータ生成と評価CSVの読み込み
        print("\n→ ダミーデータ生成を開始します。")
        df_generated = generate_raw_data(NUM_RECORDS)
        export_dataframes_to_tsv(df_generated)
        df_mail_data = get_question_data_from_csv(INPUT_QUESTION_CSV) # 生成されたCSVを読み込む
        
    elif data_source_input == '2':
        # オプション2: 実際のOutlookデータからの読み込み
        print("\n→ Outlookからの読み込みを開始します。")
        target_email = input("✅ 対象アカウントのメールアドレスを入力してください: ").strip()
        # email_processor.py の Outlook 取得関数を呼び出し、DataFrameを取得
        df_mail_data = get_mail_data_from_outlook_in_memory(TARGET_FOLDER_PATH, target_email)
    
    else:
        print(f"\n無効な入力 '{data_source_input}' です。終了します。")
        return

    # 共通処理: データフレームが空でないかチェック
    if df_mail_data.empty:
        print("処理対象のメールがありませんでした。終了します。")
        return

    # 抽出の実行
    print("\n--- 2. スキル抽出実行 ---")
    df_extracted = extract_skills_data(df_mail_data)
    
    # 評価の実行 (マスターCSVと比較)
    run_triple_csv_validation(df_extracted, MASTER_ANSWERS_PATH, OUTPUT_EVAL_PATH)
    
    print("\n💡 処理が完了しました。")


# ----------------------------------------------------
# メインディスパッチャー (プログラムの起点)
# ----------------------------------------------------

def main_dispatcher():
    """プログラムの開始点。実行モードを選択し、処理を分岐させる。"""
    
    print("\n==================================================")
    print(" 実行モードを選択してください:")
    print(" [1] 試験モード (デフォルト): ダミーデータ生成と評価を実施")
    print(" [2] メールテストモード: Outlookからメールを取得し、抽出結果をXLSXに出力")
    print("==================================================")
    
    try:
        mode_input = input("モード番号を入力してください ([1]で実行): ").strip()
        
        if not mode_input or mode_input == '1':
            # 試験モードの実行
            print("\n→ 試験モード（デフォルト）を開始します。")
            main_process_exam_mode()
            
        elif mode_input == '2':
            # メールテストモード（本番環境模擬）の実行
            print("\n→ メールテストモード（Outlook）を開始します。")
            target_email = input("✅ 対象アカウントのメールアドレスを入力してください (例: user@example.com): ").strip()
            # email_processor.py の run_email_extraction 関数を呼び出す
            run_email_extraction(target_email)
            
        else:
            print(f"\n無効な入力 '{mode_input}' です。処理を終了します。")
            
    except EOFError:
        print("\n→ 入力がないため、試験モード（デフォルト）を開始します。")
        main_process_exam_mode()
    except Exception as e:
        print(f"致命的なエラーが発生しました: {e}")
        
    # 📌 処理完了後にテスト機能を呼び出す
    interactive_id_search_test()


if __name__ == "__main__":
    # プログラムの実行開始
    main_dispatcher()