# file_processor.py (ログ・警告 削除版)

import os
import re
from docx import Document
import pdfplumber
from openpyxl import load_workbook
import unicodedata
import pandas as pd # pd.isna のためにインポート

# --- ▼▼▼ 修正: pdfplumber の警告を抑制 ▼▼▼ ---
import logging
# pdfplumber が内部で使用する pdfminer のログレベルを ERROR (エラー) のみに設定
logging.getLogger("pdfminer").setLevel(logging.ERROR)
# --- ▲▲▲ 修正ここまで ▲▲▲ ---

# ----------------------------------------------------
# ユーティリティ関数: 各ファイル形式のテキスト化
# ----------------------------------------------------

def extract_text_from_xlsx(file_path: str) -> str:
    full_text = []
    wb = None
    try:
        wb = load_workbook(file_path, read_only=True, data_only=True)
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            for row in ws.iter_rows():
                row_text = " ".join([str(cell.value) for cell in row if cell.value is not None])
                full_text.append(row_text.strip())
        return "\n".join(full_text)
    except Exception as e:
        return "" # エラー時は空文字列
    finally:
        if wb:
            try:
                wb.close()
            except Exception as close_error:
                 pass # クローズ時のエラーは無視

def extract_text_from_pdf(file_path: str) -> str:
    text = ""
    try:
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                extracted = page.extract_text(x_tolerance=1, keep_blank_chars=False) or ""
                text += extracted + "\n"
        
        if not text.strip():
            return "" # テキストがなくても空文字列

        return text.strip()
    except Exception as e:
        return "" # エラー時は空文字列


def extract_text_from_docx(file_path: str) -> str:
    full_text = []
    try:
        document = Document(file_path)
        for paragraph in document.paragraphs:
            full_text.append(paragraph.text)
        for i, table in enumerate(document.tables):
            for row in table.rows:
                row_text = " ".join([cell.text.replace('\n', ' ').strip() for cell in row.cells])
                full_text.append(row_text)
        return "\n".join(filter(None, full_text))
    except Exception as e:
        return "" # エラー時は空文字列


def get_attachment_text(temp_file_path: str, filename: str) -> str:
    file_extension = os.path.splitext(filename)[1].lower()
    
    if file_extension in ['.xlsx', '.xls']:
        raw_text = extract_text_from_xlsx(temp_file_path)
    elif file_extension == '.pdf':
        raw_text = extract_text_from_pdf(temp_file_path)
    elif file_extension == '.docx':
        raw_text = extract_text_from_docx(temp_file_path)
    else:
        # print(f"警告: 非対応の添付ファイル形式です: {filename}") # ログ削除
        return "" # 非対応なら空文字列

    # --- 抽出後の最終クリーンアップ ---
    if not raw_text or pd.isna(raw_text):
         return ""

    cleaned_text = str(raw_text).strip()
    
    # エラーメッセージが含まれていたら、空文字列を返す
    if cleaned_text.startswith("[ERROR:") or cleaned_text.startswith("[WARN:"):
         return ""

    try:
        cleaned_text = unicodedata.normalize('NFKC', cleaned_text)
        control_chars = ''.join(map(chr, list(range(0, 9)) + list(range(11, 13)) + list(range(14, 32)) + [127]))
        cleaned_text = re.sub(f'[{control_chars}\u200B\uFEFF]', '', cleaned_text)
        cleaned_text = re.sub(r'[\s\u3000]+', ' ', cleaned_text)
        cleaned_text = re.sub(r'(\s*\n\s*)+', '\n', cleaned_text)
        cleaned_text = cleaned_text.strip()
    except Exception as e:
         # print(f"エラー: テキストクリーニング中にエラーが発生しました: {e}") # ログ削除
         return str(raw_text).strip() # クリーニング失敗時は元のテキスト

    return cleaned_text