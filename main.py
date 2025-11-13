# main.py
# 目的: アプリケーションのGUI（メールテストモード）を起動する
# (★ マルチプロセス対応版 ★)

import os
import sys
import multiprocessing # 👈 ★ 1. multiprocessing をインポート
import traceback
# 外部ファイルのインポート (GUI起動に必要なモジュール)
import main_application 
# (utils.py は main_application がインポートするので、ここでは不要)

# 📌 修正1: 抽出結果ファイルのパス定義をインポート
# (v34のコードと変更なし)
try:
    from config import OUTPUT_CSV_FILE as OUTPUT_FILENAME
except ImportError:
    try:
        from email_processor import OUTPUT_FILENAME
    except ImportError:
        OUTPUT_FILENAME = 'extracted_skills_result.xlsx'


# ----------------------------------------------------
# 📌 修正2: 不要な関数を削除
# ----------------------------------------------------
# (v34のコードと変更なし)
# reorder_output_dataframe 関数 (試験モードでのみ使用) は削除されました。
# main_process_exam_mode 関数 (試験モード本体) は削除されました。
# ----------------------------------------------------


# ----------------------------------------------------
# メインディスパッチャー (プログラムの起点)
# ----------------------------------------------------

def main_dispatcher():
    """
    (v34のコードと変更なし)
    プログラムの開始点。メールテストモード(GUI)を直接起動する。
    """
    
    try:
        # GUIアプリケーションのエントリーポイントを呼び出す
        output_file_abs_path = os.path.abspath(OUTPUT_FILENAME)
        
        if not os.path.exists(output_file_abs_path):
            print(f"⚠️ 警告: 抽出結果ファイル ('{OUTPUT_FILENAME}') が見つかりません。")
            print("         GUIを起動しますが、検索一覧はファイル作成後 ('抽出実行') に利用可能です。")
        
        print("\n→ メールテストモードをGUIで開始します。")
        main_application.main() 
        
    except Exception as e:
        print(f"\n--- 致命的なエラーが発生しました ---")
        print(f"エラータイプ: {type(e).__name__}")
        print(f"エラー詳細: {e}")
        traceback.print_exc()
        print("---------------------------------")
        input("Enterキーを押して終了します...")

# ----------------------------------------------------
# プログラム実行
# ----------------------------------------------------
if __name__ == "__main__":
    
    # --- ▼▼▼ ★★★ 修正箇所 ★★★ ▼▼▼
    # マルチプロセス（EXE化）のために、
    # すべての処理の「前」にこの1行が必須です。
    multiprocessing.freeze_support()
    # --- ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲ ---
    
    main_dispatcher()