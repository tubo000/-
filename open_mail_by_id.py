# open_mail_by_id.py
# 責務: 外部コマンド（Excelなど）から渡された Entry ID を使用し、
#       Python (win32com) を介して Outlook で特定のメールを直接開く。

import win32com.client as win32 # Windows COM機能 (Outlook連携) のための必須ライブラリ
import sys                      # コマンドライン引数 (Entry ID) の取得に使用
import os                       # (このスクリプトでは未使用だが、環境依存の処理に使う可能性あり)


def open_outlook_email_by_id(entry_id: str):
    """
    指定された Entry ID を使用して、Outlookデスクトップアプリでメールを開くメイン関数。
    """
    # 1. 入力値の基本チェック
    if not entry_id:
        print("エラー: Entry IDが指定されていません。", file=sys.stderr)
        return

    try:
        # 2. Outlook アプリケーションへの接続 (COMオブジェクトの取得)
        try:
            # 既に起動している Outlook インスタンスを取得 (GetActiveObject)
            outlook_app = win32.GetActiveObject("Outlook.Application")
        except:
            # 起動していない場合、新しい Outlook インスタンスを起動 (Dispatch)
            outlook_app = win32.Dispatch("Outlook.Application")
            
        # 3. MAPI ネームスペースの取得 (Outlookのデータストアへアクセスするための窓口)
        namespace = outlook_app.GetNamespace("MAPI")
        
        # 4. Entry ID から特定のアイテム（メール）を取得
        # GetItemFromID() が Entry ID の「鍵」を使ってメールオブジェクトを特定する
        olItem = namespace.GetItemFromID(entry_id)
        
        # 5. 取得結果の表示
        if olItem:
            # アイテム（メール）を画面に表示
            olItem.Display()
            print(f"メールを正常に開きました: {getattr(olItem, 'Subject', '件名なし')}")
        else:
            print("エラー: 指定された Entry ID のメールが見つかりませんでした。", file=sys.stderr)
            
    except Exception as e:
        # 実行時エラーが発生した場合の診断メッセージ
        print(f"Outlook連携中にエラーが発生しました: {e}", file=sys.stderr)
        print("Outlookが起動しているか、またはpywin32が正しくインストールされているか確認してください。", file=sys.stderr)


# ----------------------------------------------------
# コマンドライン引数処理 (スクリプトが直接実行されたときの動作)
# ----------------------------------------------------

if __name__ == "__main__":
    # スクリプト実行時の引数の数をチェック (最低でもスクリプト名 + Entry ID の計2つ必要)
    if len(sys.argv) > 1:
        # 最初の引数 (インデックス1) が Entry ID 本体
        email_id = sys.argv[1]
        # メールオープン関数を実行
        open_outlook_email_by_id(email_id)
    else:
        # 引数が不足している場合の使用法を表示
        print("使用法: python open_mail_by_id.py [Entry ID]", file=sys.stderr)