# main.py

import os
import sys
import pandas as pd
# インポート
from config import INPUT_QUESTION_CSV, MASTER_ANSWERS_PATH, OUTPUT_EVAL_PATH, NUM_RECORDS
from data_generation import generate_raw_data, export_dataframes_to_tsv
from extraction_core import extract_skills_data
from evaluator_core import run_triple_csv_validation, get_question_data_from_csv
from email_processor import run_email_extraction # メール抽出機能をインポート


def main_process_exam_mode():
    """試験用（評価）モードのメイン実行フロー。データ生成→抽出→評価を行う。"""
    
    print("★★ 統合スキルシート抽出・評価システム（試験モード）実行 ★★")
    
    # --------------------------------------------------
    # 試験準備フェーズ (データ生成とエクスポート)
    # --------------------------------------------------
    print("\n--- 試験データ生成 ---")
    df_generated = generate_raw_data(NUM_RECORDS)
    export_dataframes_to_tsv(df_generated)
    print("----------------------")
    
    # 1. 問題CSVの読み込み
    print("\n--- 1. 問題CSVの読み込み ---")
    df_mail_data = get_question_data_from_csv(INPUT_QUESTION_CSV)
    
    if df_mail_data.empty:
        print("処理対象のメールがありませんでした。終了します。")
        return

    # 2. 抽出実行
    print("\n--- 2. スキル抽出実行 ---")
    df_extracted = extract_skills_data(df_mail_data)
    
    # 3. 評価と検証
    run_triple_csv_validation(df_extracted, MASTER_ANSWERS_PATH, OUTPUT_EVAL_PATH)
    
    print("\n💡 処理が完了しました。")


def main_dispatcher():
    """実行モードをユーザーに問い合わせ、処理を分岐させる。"""
    
    print("\n==================================================")
    print(" 実行モードを選択してください:")
    print(" [1] 試験モード (デフォルト): ダミーデータ生成と評価を実施")
    print(" [2] メールテストモード: Outlookからメールを取得し、抽出結果をCSVに出力")
    print("==================================================")
    
    try:
        mode_input = input("モード番号を入力してください ([1]で実行): ").strip()
        
        if not mode_input or mode_input == '1':
            print("\n→ 試験モード（デフォルト）を開始します。")
            main_process_exam_mode()
            
        elif mode_input == '2':
            print("\n→ メールテストモード（Outlook）を開始します。")
            
            # ★★★ アカウント入力の追加 ★★★
            target_email = input("✅ 対象アカウントのメールアドレスを入力してください (例: user@example.com): ").strip()
            
            # メール抽出関数にアカウントを渡して実行
            run_email_extraction(target_email)
            
        else:
            print(f"\n無効な入力 '{mode_input}' です。処理を終了します。")
            
    except EOFError:
        # コマンドラインなどでのEOF（入力終了）をキャッチ
        print("\n→ 入力がないため、試験モード（デフォルト）を開始します。")
        main_process_exam_mode()
    except Exception as e:
        print(f"致命的なエラーが発生しました: {e}")

if __name__ == "__main__":
    main_dispatcher()