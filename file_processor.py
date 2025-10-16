# file_processor.py
# 責務: 添付ファイル (XLSX, PDF, DOCX) をプレーンテキストに変換する。

import os
import re
from docx import Document # DOCXファイル用
import pdfplumber  # PDFファイル用 (テキストベース抽出)
from openpyxl import load_workbook # XLSXファイル用
import unicodedata # テキストの正規化（NFKC）に必須

# ----------------------------------------------------
# ユーティリティ関数: 各ファイル形式のテキスト化
# ----------------------------------------------------

def extract_text_from_xlsx(file_path: str) -> str:
    """XLSXファイルから全シートのテキストを結合して抽出する。"""
    full_text = []
    wb = None # 初期化
    try:
        # read_only=True で読み込み専用にし、ロックを防ぐ
        wb = load_workbook(file_path, read_only=True)
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            for row in ws.rows:
                # Noneではないセルの値を文字列化し、スペース区切りで結合
                row_text = " ".join([str(cell.value) for cell in row if cell.value is not None])
                full_text.append(row_text)
        
        # 📌 修正1: 抽出完了後、明示的にワークブックを閉じる
        if wb:
            wb.close()
            
        return "\n".join(full_text) # 行ごとに改行を追加

    except Exception as e:
        # 📌 修正2: エラー発生時もクローズを試みる
        if wb:
            try: wb.close()
            except: pass
            
        return f"[ERROR: XLSX処理失敗: {e}]"


def extract_text_from_pdf(file_path: str) -> str:
    """PDFファイルから全ページのテキストを抽出する。(テキストベース抽出のみ)"""
    text = ""
    try:
        # pdfplumberはwith句を使用しているため安全
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                text += page.extract_text(x_tolerance=1) or ""
        
        if not text.strip():
            return "[ERROR: PDF処理失敗: テキストベースのコンテンツが見つかりません。OCR機能は無効化されています。]"

        return text
            
    except Exception as e:
        return f"[ERROR: PDF処理失敗: {e}]"


def extract_text_from_docx(file_path: str) -> str:
    """DOCXファイルから全段落と表のテキストを抽出する。（Word対応強化）"""
    full_text = []
    try:
        # python-docxはファイルハンドルをすぐに解放するため、通常ロックの問題は発生しにくい
        document = Document(file_path)
        # 段落の抽出
        for paragraph in document.paragraphs:
            full_text.append(paragraph.text)
            
        # 表の抽出
        for i, table in enumerate(document.tables):
            full_text.append(f"\n--- TABLE_{i+1} START ---")
            for row in table.rows:
                row_text = " ".join([cell.text.replace('\n', ' ') for cell in row.cells])
                full_text.append(row_text)
            full_text.append("--- TABLE END ---\n")
            
        return "\n".join(full_text)
    except Exception as e:
        return f"[ERROR: DOCX処理失敗: {e}]"


def get_attachment_text(temp_file_path: str, filename: str) -> str:
    """
    一時保存された添付ファイルのパスを受け取り、拡張子に応じてテキストを抽出する。
    抽出後、LLM処理に適した形に簡単なクリーニングを行う。
    """
    file_extension = os.path.splitext(filename)[1].lower()
    
    # ... (元のコードを維持)
    if file_extension in ['.xlsx', '.xls']:
        raw_text = extract_text_from_xlsx(temp_file_path)
    elif file_extension == '.pdf':
        raw_text = extract_text_from_pdf(temp_file_path)
    elif file_extension == '.docx':
        raw_text = extract_text_from_docx(temp_file_path)
    else:
        return f"[WARN: 非対応ファイル形式: {file_extension}]"
        
    # 抽出後の最終クリーンアップ
    cleaned_text = raw_text.strip()
    cleaned_text = unicodedata.normalize('NFKC', cleaned_text)
    # 不要な制御文字と空白文字をスペースに置換
    cleaned_text = re.sub(r'[\r\n\t\u0020\u00A0\uFEFF\u3000\u0000-\u001F]', ' ', cleaned_text)
    # 連続するスペースを単一のスペースに置換
    cleaned_text = re.sub(r'\s+', ' ', cleaned_text).strip()
    
    return cleaned_text