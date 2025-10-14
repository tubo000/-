# main2.py

import tkinter as tk
import gui_elements
import gui_search_window

def main():
    """
    アプリケーションのメインウィンドウを作成し、実行します。
    """
    # 1. メインウィンドウの設定
    root = tk.Tk()
    root.title("Outlook Mail Search Tool")
    root.geometry("800x600")

    # 2. GUI要素の作成
    # gui_elements.py にある create_main_window_elements 関数を呼び出し、
    # メイン画面のウィジェット（entry, labelなど）を取得する
    main_elements = gui_elements.create_main_window_elements(root)
    
    # 取得したウィジェットを変数に格納 (設定画面への引数として使用するため)
    account_entry = main_elements["account_entry"]
    status_label = main_elements["status_label"]
    
    # 3. コールバック関数（ボタン押下時の処理）の定義
    
    # a. 設定ウィンドウを開く処理を定義
    # account_entry と status_label を引数として渡すことで、
    # 設定ウィンドウ内でメイン画面のウィジェットを更新できるようにします。
    def open_settings():
        # gui_elements.py で定義された関数を呼び出す
        gui_elements.open_settings_window(root, account_entry, status_label)

    # b. 検索ウィンドウを開く処理を定義
    def open_search():
        # gui_search_window.py で定義された関数を呼び出す
        gui_search_window.open_search_window(root)
    
    # 4. ボタンのコマンドを上書きまたは再設定
    # gui_elements.py で作成されたボタンに、ここで定義したコマンドを設定し直します。
    # gui_elements.py のコードに依存しますが、ここではsettings_buttonが存在すると仮定します。
    
    # "設定"ボタンのコマンドを設定
    if "settings_button" in main_elements:
        main_elements["settings_button"].config(command=open_settings)
    
    # "検索"ボタンのコマンドを設定 (open_search_windowが直接呼ばれる想定)
    if "search_button" in main_elements:
        main_elements["search_button"].config(command=open_search)
        
    # 5. アプリケーションの開始
    root.mainloop()
