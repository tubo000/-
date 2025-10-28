# gui_elements.py (open_settings_window の定義を含む)
 
import tkinter as tk
from tkinter import Frame, messagebox
import os
import pandas as pd
import utils # load_config_csv, save_config_csv に依存
 
from config import OUTPUT_EVAL_PATH ,OUTPUT_CSV_FILE, SCRIPT_DIR
import gui_search_window # open_search_window は callbackとして渡される
 
 
# --- open_settings_window の定義（設定画面）---
def open_settings_window(root, account_entry_main, status_label_main):
    """
    設定ウィンドウを開き、アカウント名を保存する処理。
    メインウィンドウのウィジェット (account_entry_main, status_label_main) を更新する。
    """
    settings_window = tk.Toplevel(root)
    settings_window.title("アカウント名初期設定")
   
    # ウィンドウサイズの計算（UIロジックの維持）
    WINDOW_WIDTH = 400; WINDOW_HEIGHT = 150
    screen_width = root.winfo_screenwidth(); screen_height = root.winfo_screenheight()
    x = int((screen_width / 2) - 200); y = int((screen_height / 2) - 75)
    settings_window.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}+{x}+{y}")
   
    tk.Label(settings_window, text="名前の初期設定が可能です。", font=("Arial", 10, "bold")).pack(pady=10)
    current_account, _ = utils.load_config_csv()
    tk.Label(settings_window, text="Outlookアカウント (メールアドレス):").pack(padx=10, pady=(5, 0), anchor='w')
   
    email_entry = tk.Entry(settings_window, width=50)
    email_entry.insert(0, current_account)
    email_entry.pack(padx=10, fill='x')
   
    def save_and_close():
        new_account = email_entry.get().strip()
        if not new_account: messagebox.showerror("エラー", "メールアドレスは空にできません。"); return
       
        success, message = utils.save_config_csv(new_account)
        if success:
            account_entry_main.delete(0, tk.END); account_entry_main.insert(0, new_account)
            status_label_main.config(text=f" 設定: アカウント名が '{new_account}' に更新されました。", fg="purple")
            messagebox.showinfo("保存完了", message); settings_window.destroy()
        else:
            messagebox.showerror("保存エラー", message)
           
    save_button = tk.Button(settings_window, text="上書き保存", command=save_and_close, bg="#FFA07A")
    save_button.pack(pady=10)
 
 
# --- create_main_window_elements の定義（メイン画面の構成）---
def create_main_window_elements(root, setting_frame, saved_account, saved_folder,
                                run_extraction_callback, open_settings_callback, open_search_callback):
    """
    メイン画面のウィジェットを定義し、配置する。
    """
   
    # 1. アカウント設定要素 (配置ロジックは以前の通り)
    tk.Label(setting_frame, text="Outlookアカウント (メールアドレス/表示名) (必須):").grid(row=0, column=0, sticky="w", pady=5)
    account_entry = tk.Entry(setting_frame, width=40)
    account_entry.insert(0, saved_account)
    account_entry.grid(row=0, column=1, sticky="ew", padx=5)
    setting_frame.grid_columnconfigure(1, weight=1)
   
    tk.Label(setting_frame, text="対象フォルダパス (必須):").grid(row=1, column=0, sticky="w", pady=5)
    folder_entry = tk.Entry(setting_frame, width=40)
    folder_entry.insert(0, saved_folder)
    folder_entry.grid(row=1, column=1, sticky="ew", padx=5)
 
    tk.Label(setting_frame, text="入力例: 受信トレイ\\プロジェクトXのスキルシート", fg="gray", font=("Arial", 9)).grid(row=2, column=1, sticky="w", padx=5)
 
    # 2. ステータスラベル
    status_label = tk.Label(root, text="準備完了。設定を確認し、ボタンを押してください。", fg="black", padx=10, pady=15, font=("Arial", 10))
    status_label.pack(fill='x')
 
    # 3. メインボタンフレーム
    main_button_frame = Frame(root)
    main_button_frame.pack(pady=5)
   
    # 📌 抽出ボタン (read_button)
    read_button = tk.Button(
        main_button_frame,
        text="抽出を実行",
        command=run_extraction_callback,
        bg="#4CAF50", fg="white", font=("Arial", 12, "bold"), width=20
    )
    read_button.pack(side=tk.LEFT, padx=10)
 
    # 📌 検索ボタン (search_button)
    search_button = tk.Button(
        main_button_frame,
        text="検索・結果一覧表示",
        command=open_search_callback,
        bg="#2196F3", fg="white", font=("Arial", 12), width=20
    )
    search_button.pack(side=tk.LEFT, padx=10)
 
    # 4. CSV存在チェックに基づいたボタンの有効化/無効化
    output_csv_path = os.path.join(SCRIPT_DIR, OUTPUT_CSV_FILE)
    is_csv_exists_and_not_empty = False
    if os.path.exists(output_csv_path):
        try:
            if not pd.read_excel(output_csv_path).empty:
                 is_csv_exists_and_not_empty = True
        except Exception:
            pass
 
    if not is_csv_exists_and_not_empty:
        search_button.config(state=tk.DISABLED)
 
    # 5. 設定ボタン
    settings_button_frame = Frame(root)
    settings_button_frame.pack(pady=(0, 5))
 
    tk.Label(settings_button_frame, text="アカウント名の初期設定はこちらから", fg="gray").pack(side=tk.TOP, pady=(5, 0))
 
    settings_button = tk.Button(
        settings_button_frame,
        text="設定",
        command=open_settings_callback,
        bg="#808080", fg="white", font=("Arial", 10), width=15
    )
    settings_button.pack(side=tk.TOP, padx=10, pady=5)
   
    # 必要なウィジェットを返す
    return {
        "account_entry": account_entry,
        "folder_entry": folder_entry,
        "status_label": status_label,
        "settings_button": settings_button,
        "search_button": search_button
    }
 