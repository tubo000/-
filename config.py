#config.py
#定数、ファイルパス、抽出キーワード、正規表現パターンの管理
#変数の宣言
import os
import sys 

def get_script_dir():
    """実行中のスクリプト（またはexe）のディレクトリパスを確実に取得する。"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))
    

TARGET_FOLDER_PATH = "受信トレイ" # 処理対象とするOutlookフォルダのパス
OUTPUT_EVAL_PATH = 'evaluation_results.xlsx' # 最終出力ファイル名 (ベース)
SCRIPT_DIR = get_script_dir()
CONFIG_FILE_PATH = os.path.join(SCRIPT_DIR, 'name.csv') # Outlookアカウント名を保存する設定ファイル

# メールを選別するためのキーワード（件名または本文に含まれるかチェック）
MUST_INCLUDE_KEYWORDS = [r'スキルシート', r'業務経歴書', r'人材のご紹介', r'リソース'] # 必須キーワード
EXCLUDE_KEYWORDS = [r'請求書', r'セミナー', r'お問い合わせ', r'休暇申請'] # 除外キーワード

# 抽出パターンの定義: 正規表現と信頼度スコア
ITEM_PATTERNS = {
    '氏名': [
        {'pattern': r'(?:名\s*前|名前|氏名)\s*[：:]\s*([^(\n]+)(?:\s*\(.+\))?', 'score': 100}, # 氏名：[値]
        {'pattern': r'様\s*\n\s*([^(\n]+)\s*殿', 'score': 90}, # (本文冒頭で) ○○様 〇〇殿 形式
    ],
    '年齢': [
        {'pattern': r'年\s*齢[：:]\s*(\d+)\s*(?:歳|才)?[^\n]*', 'score': 100, 'cleanup_type': 'int'}, # 年齢：35歳, 年齢:35
        {'pattern': r'\(\s*(\d+)\s*歳', 'score': 80, 'cleanup_type': 'int'}, # (35歳) の括弧内の年齢
    ],
    '単金': [
        {'pattern': r'単\s*金[：:]\s*([^(\n]+)', 'score': 100, 'cleanup_type': 'int'}, # 単金：〇〇
        {'pattern': r'(?:月額|報酬|MP|単価)\s*[：:]\s*([^(\n]+)', 'score': 90, 'cleanup_type': 'int'}, # 月額：〇〇, 単価：〇〇
        {'pattern': r'[\s　]約\s*(\d+)\s*万円', 'score': 70, 'cleanup_type': 'int'}, # 約 70 万円
    ],
    '業務_業種': [
        {'pattern': r'【業\s*務】\s*([^\n]+)', 'score': 100},
        {'pattern': r'(?:業務|業種)\s*[：:]\s*([^\n]+)', 'score': 80},
    ],
    'スキル_言語': [
        {'pattern': r'(?:【言\s*語】|得意言語|使用言語)\s*[：:]\s*([^\n]+)', 'score': 100},
        {'pattern': r'スキル\s*[^:\n]{0,10}[：:]\s*([^\n]+)', 'score': 80},
    ],
    'スキル_OS': [
        {'pattern': r'(?:O\s*S|基\s*盤|基本O\s*S)\s*[：:]\s*([^\n]+)', 'score': 100}, 
    ],
}

OUTPUT_CSV_FILE = OUTPUT_EVAL_PATH.replace('.xlsx', '_final.csv') # 最終出力CSVファイル名
INTERMEDIATE_CSV_FILE = 'intermediate_mail_data.csv' # 一時的な中間ファイル名

# =========================================================
# 💡 ユーティリティ/ファイル処理関数
# =========================================================
#EXE化していても動くように絶対パスを参照するためのコード
def get_script_dir():
    """実行中のスクリプト（またはexe）のディレクトリパスを確実に取得する。"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))