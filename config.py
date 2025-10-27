# config.py

import os
import sys
import re

# =================================================================
# 【定数定義と設定】
# =================================================================

# ----------------------------------------------------
# 💡 ユーティリティ/ファイル処理関数
# ----------------------------------------------------
def get_script_dir():
    if getattr(sys, 'frozen', False):
        # exeファイルとして実行されている場合（PyInstaller環境）
        # sys.executableはexeファイルのフルパスを返す
        return os.path.dirname(sys.executable)
    else:
        # 通常のPythonスクリプトとして実行されている場合
        # __file__は現在のスクリプトのパスを返す
        return os.path.dirname(os.path.abspath(__file__))
    
SCRIPT_DIR = get_script_dir()
# ----------------------------------------------------

# ★★★ ファイルパスと設定 ★★★
INPUT_QUESTION_CSV = 'question_source.csv'
MASTER_ANSWERS_PATH = 'master_answers.csv'
OUTPUT_EVAL_PATH = 'evaluation_results.csv'
NUM_RECORDS = 150 
TARGET_FOLDER_PATH = "受信トレイ" 

# 📌 GUIの依存関係を解消するための定義
OUTPUT_CSV_FILE = OUTPUT_EVAL_PATH 
INTERMEDIATE_CSV_FILE = 'intermediate_mail_data.csv' 
CONFIG_FILE_PATH = os.path.join(SCRIPT_DIR, 'name.csv') 

# ★★★ 評価・出力項目 ★★★
EVALUATION_TARGETS = [ '件名','年齢', '単金', '期間_開始', 'スキルor言語','ポジション', ] 

MASTER_COLUMNS = [
    'EntryID', '件名', '性別', '年齢', '単金','期間_開始',
    'スキルor言語', 'ポジション', '本文(テキスト形式)'
    #'名前', 'マネジメント経験人数', '技術経験年数', '得意技術',
    #'開発工程_要件定義', '開発工程_基本設計', '開発工程_詳細設計', 
    #'開発工程_製造', '開発工程_結合テスト', '開発工程_システムテスト', 
    #'開発工程_運用・保守', '開発工程_その他',
    #'業種', '職務', '人数', '開発ツール', 'OS', 
    #'フレームワーク/ライブラリ', 'データベース', 'その他', '備考', 
    #'社名', '宛先メール', '開発手法', '信頼度スコア',
]

# ★★★ 辞書データ（ダミーデータ生成用） ★★★
LANGUAGES = ['Java', 'Python', 'C#', 'C++', 'JavaScript', 'SQL', 'Go', 'PHP', 'Ruby', 'Kotlin'] 
INDUSTRIES = ['金融', '医療', 'IT/Web開発', '製造業', '物流', '通信', '公共', 'インフラ'] 
NAMES = [
    ('田中', '太郎', 0), ('佐藤', '花子', 0), ('鈴木', '一郎', 0), 
    ('高橋', '恵美', 0), ('中村', '雄大', 0), ('渡辺', '彩乃', 0),
    ('Kim', 'Park', 1), ('John', 'Smith', 1), ('David', 'Lee', 1)
]
SALARY_UNITS = [
    '万円 (固定)', '万 (税込)', '万', '000k', r'円/月'
]
NOISE = [
    "【特記事項】柔軟な勤務が可能です。",
    "これは関係ない備考です。",
    "連絡先は別途送付のPDFを参照ください。",
    ""
]

# ----------------------------------------------------
# 必須キーワード/除外キーワードの定義 (Outlookフィルタリング用)
# ----------------------------------------------------
MUST_INCLUDE_KEYWORDS = [r'スキルシート', r'業務経歴書', r'人材のご紹介', r'リソース']
EXCLUDE_KEYWORDS = [r'請求書', r'セミナー', r'お問い合わせ', r'休暇申請',]
# ----------------------------------------------------

# 年齢の正規表現 
AGE_PATTERN_2_DIGITS = r'([2-9][0-9])'
NEGATIVE_LOOKAHEAD_DATE = r'(?!\s*(?:年|月|日|/|-|\)))'

# 年齢の多様なパターン（必須キーワード付き、単位付き、単位なし）
RE_AGE_KEYWORD = r'[【\[（\(]?\s*年\s*齢\s*[】\]）\)]?[\s:：]*' + AGE_PATTERN_2_DIGITS + r'[歳才]'
RE_AGE_PARENTHESIS = r'[a-zA-Z\s\.]+ *[（\(].*?' + AGE_PATTERN_2_DIGITS + r'[歳才].*?[）\)]' 
RE_AGE_AFTER_NAME = r'[a-zA-Z\s\.]+\s+' + AGE_PATTERN_2_DIGITS + r'[歳才]'
RE_AGE_KEYWORD_NO_UNIT = r'[【\[（\(]?\s*年\s*齢\s*[】\]）\)]?[\s:：]*' + AGE_PATTERN_2_DIGITS + NEGATIVE_LOOKAHEAD_DATE
#優先順位
RE_AGE_PATTERNS = [
RE_AGE_KEYWORD, RE_AGE_PARENTHESIS, RE_AGE_AFTER_NAME, RE_AGE_KEYWORD_NO_UNIT
] # 抽出関数で使用

# 単金の正規表現 
KEYWORD_TANAKA = r'(?:単金|単価|希望単価|【単金】|【単価】|金額)' 
PREFIX_TANAKA = r'.{0,5}' + KEYWORD_TANAKA 

# 1. キーワード付きパターン (優先度順: 範囲指定 > 単一値)
RE_TANAKA_RANGE_KEYWORD = PREFIX_TANAKA + r'[：:\s]*([\d\.,]+万?[～~-][\d\.,]+万?円?)'
RE_TANAKA_VALUE_MAN_KEYWORD = PREFIX_TANAKA + r'[：:\s]*([\d\.,]+万)'
RE_TANAKA_VALUE_YEN_KEYWORD = PREFIX_TANAKA + r'[：:\s]*([\d\.,]+円)'
#優先順位
RE_TANAKA_KW_PATTERNS = [
    RE_TANAKA_RANGE_KEYWORD, RE_TANAKA_VALUE_MAN_KEYWORD, RE_TANAKA_VALUE_YEN_KEYWORD
] # 抽出関数で使用

# 2. キーワードなし (RAW) パターン (優先度順: 00~00万 > 00万 > 00~00 > 00)
RE_TANAKA_RANGE_MAN_RAW = r'(?:\D|^)([\d\.,]+万?[～~-][\d\.,]+万)(?:\D|$)' 
RE_TANAKA_VALUE_MAN_RAW = r'(?:\D|^)([\d\.,]+万)(?:\D|$)' 
RE_TANAKA_RANGE_RAW_NO_MAN = r'(?:\D|^)([\d\.,]+[～~-][\d\.,]+)(?:\D|$)'
RE_TANAKA_VALUE_RAW = r'(?:\D|^)([\d\.,]{5,})(?:\D|$)' 
RE_TANAKA_RAW_PATTERNS = [
RE_TANAKA_RANGE_MAN_RAW, RE_TANAKA_VALUE_MAN_RAW, RE_TANAKA_RANGE_RAW_NO_MAN, RE_TANAKA_VALUE_RAW
] # 抽出関数で使用


# ★★★ 正規表現パターン (旧ロジック。名前とその他の項目は維持) ★★★
# 💡 名前とその他の項目のパターンは、この定数に集約されているため、そのまま使用します。
#    年齢と単金のパターンは、上記の高度な正規表現（RE_AGE_PATTERNS, RE_TANAKA_KW_PATTERNS, etc.）を使用するため、
#    この辞書内のパターンは、**抽出ロジックで無視**されますが、定義は残します。
ITEM_PATTERNS = {
    #'名前': {'pattern': r'(?:名前|氏名)[:：]\s*([^\n\r]+?)(?:\s*(?:年齢|単金|スキル|備考|社名|$))', 'score': 100}, 
    '年齢': {'pattern': r'年齢[:：]\s*([\d\s\-\-～~]+)\s*(?:歳|代)?', 'score': 100}, # ダミー/旧パターン
    '単金': {'pattern': r'単金[:：]\s*([\d,\s\.\-～~]+)\s*(?:万|円/月|\(.*\))?', 'score': 100}, # ダミー/旧パターン
    '期間_開始': {'pattern': r'期間[:：]\s*(?:(\d{4}[\s\/\-]\d{1,2})|(\d{1,2}月))', 'score': 80},
    #'性別': {'pattern': r'(?:名前|氏名)[:：].*?\((男性|女性|Male|Female)\)', 'score': 50}, 
    'スキルor言語': {'pattern': r'(?:スキル|【言語】)[:：]\s*(.*?)(?:【業務】|備考|-{5,}|$)', 'score': 100},
    #'OS': {'pattern': r'(?:【OS】|OS|環境OS)[:：]\s*([^\n\r]+?)(?:データベース|フレームワーク|$)', 'score': 100},
    #'データベース': {'pattern': r'(?:【DB】|データベース)[:：]\s*([^\n\r]+?)(?:フレームワーク|$)', 'score': 100},
    #'フレームワーク/ライブラリ': {'pattern': r'(?:フレームワーク/ライブラリ|FW)[:：]\s*([^\n\r]+?)(?:開発ツール|$)', 'score': 100},
    #'開発ツール': {'pattern': r'(?:ツール|開発ツール)[:：]\s*([^\n\r]+?)(?:得意技術|$)', 'score': 100},
    #'得意技術': {'pattern': r'得意技術[:：]\s*([^\n\r]+?)(?:役割|期間|$)', 'score': 80},
    #'ポジション': {'pattern': r'(?:役割|役職|ポジション|PMO|職種)[:：]\s*([^\n\r]+?)(?:\s*(?:マネジメント経験人数|技術経験年数|期間|$))', 'score': 80},
    #'マネジメント経験人数': {'pattern': r'(?:マネジメント|管理経験|管理人数).*?(\d+)\s*名', 'score': 80},
    #'技術経験年数': {'pattern': r'経験年数[:：]\s*(\d+)\s*年', 'score': 80},
    #'開発手法': {'pattern': r'(?:開発手法|開発プロセス)[:：](Agile|アジャイル|Waterfall|ウォーターフォール)', 'score': 80},
    #'業種': {'pattern': r'(?:【業務】|業務知識|業種)[:：]\s*([^\n\r]+?)(?:\s*(?:職務|期間|備考|$))', 'score': 100},
    #'職務': {'pattern': r'(?:職務|職種|担当業務)[:：]\s*([^\n\r]+?)(?:\s*(?:期間|人数|$))', 'score': 80},
    #'備考': {'pattern': r'(?:備考|特記事項)[:：]\s*([^\n\r]+?)(?:社名|$)', 'score': 50},
    #'社名': {'pattern': r'(?:社名|所属会社|所属)[:：]\s*([^\n\r]+?)(?:宛先メール|$)', 'score': 50},
    #'宛先メール_本文内': {'pattern': r'(?:Email|連絡先)[:：]([^\s\n]+@[^\s\n]+)', 'score': 50},
}

# 開発工程のキーワード定義（フラグ用）
PROCESS_KEYWORDS = {
    #'要件定義': [r'要件定義', r'要求分析'],
    #'基本設計': [r'基本設計', r'外部設計', r'機能設計'],
    #'詳細設計': [r'詳細設計', r'内部設計'],
    #'製造': [r'製造', r'プログラミング', r'コーディング'],
    #'結合テスト': [r'結合テスト', r'IT'],
    #'システムテスト': [r'システムテスト', r'ST', r'総合テスト'],
    #'運用・保守': [r'運用', r'保守', r'メンテナンス'],
    #'その他': [r'キックオフ', r'リリース', '管理'],
}