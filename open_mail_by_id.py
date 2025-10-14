# open_mail_by_id.py

import win32com.client as win32
import sys
import os

def open_outlook_email_by_id(entry_id: str):
    """
    指定された Entry ID を使用して、Outlookデスクトップアプリでメールを開く関数
    """
    if not entry_id:
        print("エラー: Entry IDが指定されていません。", file=sys.stderr)
        return

    try:
        # Outlook アプリケーションへの接続 (起動していなければ起動する)
        # GetActiveObjectで起動中のOutlookを取得し、失敗したらDispatchで起動を試みる
        try:
            outlook_app = win32.GetActiveObject("Outlook.Application")
        except:
            outlook_app = win32.Dispatch("Outlook.Application")
            
        # MAPI ネームスペースを取得
        namespace = outlook_app.GetNamespace("MAPI")
        
        # Entry ID から特定のアイテム (メール) を取得
        olItem = namespace.GetItemFromID(entry_id)
        
        if olItem:
            # アイテム (メール) を表示
            olItem.Display()
            print(f"メールを正常に開きました: {getattr(olItem, 'Subject', '件名なし')}")
        else:
            print("エラー: 指定された Entry ID のメールが見つかりませんでした。", file=sys.stderr)
            
    except Exception as e:
        print(f"Outlook連携中にエラーが発生しました: {e}", file=sys.stderr)
        print("Outlookが起動しているか、またはpywin32が正しくインストールされているか確認してください。", file=sys.stderr)

# ----------------------------------------------------
# コマンドライン引数処理
# ----------------------------------------------------

if __name__ == "__main__":
    # Python スクリプトの引数をチェック
    if len(sys.argv) > 1:
        # 最初の引数 (argv[1]) が Entry ID
        email_id = sys.argv[1]
        open_outlook_email_by_id(email_id)
    else:
        print("使用法: python open_mail_by_id.py [Entry ID]", file=sys.stderr)