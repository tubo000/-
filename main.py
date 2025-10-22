# main.py
# 目的: アプリケーションの実行フローを制御し、試験モードと本番テストモードの分岐を行う

import os
import sys
import pandas as pd
import win32com.client as win32 
# 外部ファイルのインポート (GUI/CLI機能に必要な定数と関数)
from config import INPUT_QUESTION_CSV, MASTER_ANSWERS_PATH, OUTPUT_EVAL_PATH, NUM_RECORDS, TARGET_FOLDER_PATH
from data_generation import generate_raw_data, export_dataframes_to_tsv
from extraction_core import extract_skills_data
from evaluator_core import run_triple_csv_validation, get_question_data_from_csv
from email_processor import get_mail_data_from_outlook_in_memory, OUTPUT_FILENAME

import main_application 
import utils 


# ----------------------------------------------------
# ユーティリティ関数群 (ID検索関連は削除)
# ----------------------------------------------------

# 📌 削除: open_outlook_email_by_id 関数は削除 (ID関連)
# 📌 削除: interactive_id_search_test 関数は削除 (ID関連)


def reorder_output_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """出力データフレームの列順を調整し、特定の項目を左側に固定する。"""
    fixed_leading_cols = [
        'メールURL', '件名', '名前', '信頼度スコア', 
        '本文(ファイル含む)', '本文(テキスト形式)', 'Attachments'
    ]
    
    # DataFrameに存在するカラムのみをフィルタリング
    fixed_leading_cols = [col for col in fixed_leading_cols if col in df.columns]
    
    all_cols = df.columns.tolist()
    
    # leading_colsに含まれない残りのカラム
    remaining_cols = [col for col in all_cols if col not in fixed_leading_cols]
    
    final_col_order = fixed_leading_cols + remaining_cols
    
    # 存在しない列は drop されるため、reindexを使用
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
    output_file_abs_path = os.path.abspath(OUTPUT_FILENAME)

    print("★★ 統合スキルシート抽出・評価システム（試験モード）実行 ★★")
    
    # データソースの選択プロンプト
    print("\n--- 試験データの選択 ---")
    print(" [1] ダミーデータ生成 (デフォルト): 新規データを作成しCSVから読み込み")
    print(" [2] Outlookメールから読み込み: 実際のメールデータを使用")
    
    try:
        data_source_input = input("データソースを選択してください ([1]で実行): ").strip()
        df_mail_data = pd.DataFrame()

        if not data_source_input or data_source_input == '1':
            # オプション1: ダミーデータ生成と評価CSVの読み込み
            print("\n→ ダミーデータ生成を開始します。")
            df_generated = generate_raw_data(NUM_RECORDS)
            export_dataframes_to_tsv(df_generated)
            df_mail_data = get_question_data_from_csv(INPUT_QUESTION_CSV) # 生成されたCSVを読み込む
            
        elif data_source_input == '2':
            # GUIアプリケーションのエントリーポイントを呼び出す
            if not os.path.exists(output_file_abs_path):
                print(f"⚠️ 警告: 抽出結果ファイル ('{OUTPUT_FILENAME}') が見つかりません。")
                print("GUIを起動しますが、検索一覧はファイル作成後 ('抽出実行') に利用可能です。")
            
            print("\n→ メールテストモードをGUIで開始します。")
            main_application.main()
            return # GUI起動後は、このCLIプロセスはここで終了（または待機）

        else:
            print(f"\n無効な入力 '{data_source_input}' です。終了します。")
            return
            
    except EOFError:
        print("\n→ 入力がないため、試験モードを開始します。")
        # 処理続行

    except Exception as e:
        print(f"致命的なエラーが発生しました: {e}")
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
    
    # 📌 修正3: デフォルトのモードを [2] GUIモードに変更
    default_mode = '2'
    
    print("\n==================================================")
    print(" 実行モードを選択してください:")
    print(" [1] 試験モード (コンソール実行)")
    print(" [2] メールテストモード (デフォルト): GUIアプリケーションで実行")
    print("==================================================")
    
    try:
        # デフォルト値を '2' に設定
        mode_input = input(f"モード番号を入力してください ([{default_mode}]で実行): ").strip() or default_mode
        
        if mode_input == '1':
            # 試験モードの実行 (コンソールベース)
            print("\n→ 試験モード（コンソール）を開始します。")
            main_process_exam_mode()
            
        elif mode_input == '2':
            # GUIアプリケーションのエントリーポイントを呼び出す
            output_file_abs_path = os.path.abspath(OUTPUT_FILENAME)
            
            if not os.path.exists(output_file_abs_path):
                print(f"⚠️ 警告: 抽出結果ファイル ('{OUTPUT_FILENAME}') が見つかりません。")
                print("         GUIを起動しますが、検索一覧はファイル作成後 ('抽出実行') に利用可能です。")
            
            print("\n→ メールテストモードをGUIで開始します。")
            main_application.main() 
            
        else:
            print(f"\n無効な入力 '{mode_input}' です。処理を終了します。")
            
    except EOFError:
        print("\n→ 入力がないため、GUIモードを開始します。")
        main_application.main()
        
    except Exception as e:
        print(f"致命的なエラーが発生しました: {e}")


if __name__ == "__main__":
    # プログラムの実行開始
    main_dispatcher()