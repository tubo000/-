# main.py
# 目的: アプリケーションのGUI（メールテストモード）を起動する

import os
import sys
# 外部ファイルのインポート (GUI起動に必要なモジュール)
import main_application 
import utils 

# 📌 修正1: 抽出結果ファイルのパス定義をインポート
# config.py から OUTPUT_CSV_FILE を OUTPUT_FILENAME としてエイリアス
try:
    from config import OUTPUT_CSV_FILE as OUTPUT_FILENAME
except ImportError:
    # config.py に OUTPUT_CSV_FILE がない場合、email_processor.py からフォールバック
    try:
        from email_processor import OUTPUT_FILENAME
    except ImportError:
        # 両方にない場合、デフォルトのファイル名を定義
        OUTPUT_FILENAME = 'extracted_skills_result.xlsx'


# ----------------------------------------------------
# 📌 修正2: 不要な関数を削除
# ----------------------------------------------------
# reorder_output_dataframe 関数 (試験モードでのみ使用) は削除されました。
# main_process_exam_mode 関数 (試験モード本体) は削除されました。
# ----------------------------------------------------


# ----------------------------------------------------
# メインディスパッチャー (プログラムの起点)
# ----------------------------------------------------

def main_dispatcher():
    """プログラムの開始点。メールテストモード(GUI)を直接起動する。"""
    
    # 📌 修正3: モード選択ロジックを削除
    
    try:
        # GUIアプリケーションのエントリーポイントを呼び出す
        output_file_abs_path = os.path.abspath(OUTPUT_FILENAME)
        
        # 起動時のファイル存在チェック (コンソールへの警告)
        if not os.path.exists(output_file_abs_path):
            print(f"⚠️ 警告: 抽出結果ファイル ('{OUTPUT_FILENAME}') が見つかりません。")
            print("         GUIを起動しますが、検索一覧はファイル作成後 ('抽出実行') に利用可能です。")
        
        print("\n→ メールテストモードをGUIで開始します。")
        main_application.main() 
            
    except Exception as e:
        print(f"致命的なエラーが発生しました: {e}")
        traceback.print_exc() # 📌 修正4: エラー詳細を表示するため traceback を追加
        input("エラーのため終了します。Enterキーを押してください...")


if __name__ == "__main__":
    # 📌 修正5: traceback を使用するためにインポート
    import traceback 
    
    # プログラムの実行開始
    main_dispatcher()