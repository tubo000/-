# file_processor.py
import pandas as pd # ← ここでインポートされています
import os
import re
from docx import Document # DOCXファイル用
import pdfplumber # PDFファイル用 (テキストベース抽出)
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
        wb = load_workbook(file_path, read_only=True, data_only=True) # data_only=True で数式の代わりに計算結果を取得
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            for row in ws.iter_rows(): # iter_rows() を使う方が効率的な場合がある
                # Noneではないセルの値を文字列化し、スペース区切りで結合
                row_text = " ".join([str(cell.value) for cell in row if cell.value is not None])
                full_text.append(row_text.strip()) # 各行の前後の空白を削除
        
        # 結合する前に空行を除去する場合
        # full_text = [line for line in full_text if line] 
        
        return "\n".join(full_text) # 行ごとに改行を追加

    except Exception as e:
        # エラーメッセージを返す
        return f"[ERROR: XLSX処理失敗: {e}]"
        
    finally:
        # 📌 修正: finallyブロックで確実にファイルを閉じる
        if wb:
            try:
                wb.close()
            except Exception as close_error:
                 # クローズ時のエラーはログに出力するなどしても良い
                 print(f"警告: XLSXファイルのクローズ中にエラー: {close_error}")


def extract_text_from_pdf(file_path: str) -> str:
    # ... (変更なし) ...
    """PDFファイルから全ページのテキストを抽出する。(テキストベース抽出のみ)"""
    text = ""
    try:
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                # x_tolerance を調整して、スペースが過剰に入るのを抑制できる場合がある
                extracted = page.extract_text(x_tolerance=1, keep_blank_chars=False) or ""
                text += extracted + "\n" # ページごとに改行を追加
        
        if not text.strip():
            return "[WARN: PDFからテキストを抽出できませんでした。画像ベースのPDFか空のファイルの可能性があります。]"

        return text.strip() # 最後にもう一度strip
            
    except Exception as e:
        return f"[ERROR: PDF処理失敗: {e}]"


def extract_text_from_docx(file_path: str) -> str:
    # ... (変更なし) ...
    """DOCXファイルから全段落と表のテキストを抽出する。（Word対応強化）"""
    full_text = []
    try:
        document = Document(file_path)
        # 段落の抽出
        for paragraph in document.paragraphs:
            full_text.append(paragraph.text)
            
        # 表の抽出
        for i, table in enumerate(document.tables):
            # full_text.append(f"\n--- TABLE_{i+1} START ---") # 抽出テキストとしては不要かも
            for row in table.rows:
                # セル内の改行もスペースに置換し、前後の空白を削除
                row_text = " ".join([cell.text.replace('\n', ' ').strip() for cell in row.cells])
                full_text.append(row_text)
            # full_text.append("--- TABLE END ---\n")
            
        return "\n".join(filter(None, full_text)) # 空行を除去して結合
    except Exception as e:
        return f"[ERROR: DOCX処理失敗: {e}]"


def get_attachment_text(temp_file_path: str, filename: str) -> str:
    # ... (変更なし、ただしクリーニング処理を少し調整) ...
    """
    一時保存された添付ファイルのパスを受け取り、拡張子に応じてテキストを抽出する。
    抽出後、LLM処理に適した形に簡単なクリーニングを行う。
    """
    file_extension = os.path.splitext(filename)[1].lower()
    
    if file_extension in ['.xlsx', '.xls']:
        raw_text = extract_text_from_xlsx(temp_file_path)
    elif file_extension == '.pdf':
        raw_text = extract_text_from_pdf(temp_file_path)
    elif file_extension == '.docx':
        raw_text = extract_text_from_docx(temp_file_path)
    # 📌 '.doc' ファイルへの対応を追加する場合は、pywin32など別のライブラリが必要
    # elif file_extension == '.doc':
    #     raw_text = extract_text_from_doc(temp_file_path) # 別途定義が必要
    else:
        # 📌 修正: 警告メッセージをより具体的にする
        print(f"警告: 非対応または拡張子不明な添付ファイル形式です。ファイル名: '{filename}', 検出された拡張子: '{file_extension}'")
        return "" # 空文字列を返す
        # return f"[WARN: 非対応ファイル形式: {file_extension}]"

    # --- 抽出後の最終クリーンアップ ---
    
    # 最初にNoneや空でないことを確認
    if not raw_text or pd.isna(raw_text):
         return ""

    cleaned_text = str(raw_text).strip()
    
    # エラーメッセージが含まれていたら、そのまま返す
    if cleaned_text.startswith("[ERROR:") or cleaned_text.startswith("[WARN:"):
         return cleaned_text

    try:
        # 全角英数字、記号などを半角に正規化 (NFKC)
        cleaned_text = unicodedata.normalize('NFKC', cleaned_text)
        
        # 制御文字（改行、タブを除く）とゼロ幅スペースなどを削除
        # \r, \n, \t は保持し、他の制御文字 (\x00-\x1F, \x7F) を削除
        # U+200B (ゼロ幅スペース), U+FEFF (BOM) なども削除対象に含める
        control_chars = ''.join(map(chr, list(range(0, 9)) + list(range(11, 13)) + list(range(14, 32)) + [127]))
        cleaned_text = re.sub(f'[{control_chars}\u200B\uFEFF]', '', cleaned_text)
        
        # 複数の改行やスペースが混在している場合、単一の改行に置換
        cleaned_text = re.sub(r'[\s\u3000]+', ' ', cleaned_text) # 全角スペースも半角スペースに
        cleaned_text = re.sub(r'(\s*\n\s*)+', '\n', cleaned_text) # 連続する改行（前後のスペース含む）を単一改行に
        
        # 文頭・文末の空白・改行を削除
        cleaned_text = cleaned_text.strip()

    except Exception as e:
         print(f"エラー: テキストクリーニング中にエラーが発生しました: {e}")
         # クリーニング失敗時は、元のテキスト（strip済み）を返す
         return str(raw_text).strip()

    return cleaned_text

# --- .doc ファイル用の関数（pywin32が必要） ---
# def extract_text_from_doc(file_path: str) -> str:
#     """DOCファイルからテキストを抽出する (Windows + Word必須)"""
#     try:
#         import win32com.client as win32
#         import pythoncom
#         pythoncom.CoInitialize()
#         word = None
#         doc = None
#         try:
#             # フルパスに変換
#             abs_path = os.path.abspath(file_path)
#             word = win32.Dispatch("Word.Application")
#             word.Visible = False # Wordを画面に表示しない
#             # ファイルを開く
#             doc = word.Documents.Open(abs_path, ReadOnly=True)
#             text = doc.Content.Text
#             doc.Close(False) # 保存せずに閉じる
#             word.Quit()
#             pythoncom.CoUninitialize()
#             return text
#         except Exception as e:
#              if doc:
#                   try: doc.Close(False)
#                   except: pass
#              if word:
#                   try: word.Quit()
#                   except: pass
#              pythoncom.CoUninitialize()
#              return f"[ERROR: DOC処理失敗 (win32com): {e}]"
#     except ImportError:
#         return "[ERROR: DOC処理失敗: pywin32 がインストールされていません]"
#     except Exception as e:
#          return f"[ERROR: DOC処理失敗: {e}]"