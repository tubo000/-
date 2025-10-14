# gui_elements.py
#ボタン配置
import tkinter as tk
from tkinter import Frame
import os 
import pandas as pd

from config import OUTPUT_CSV_FILE,SCRIPT_DIR
from utils import save_config_csv ,load_config_csv
from gui_search_window import open_search_window
from data_processor import run_extraction_workflow
from gui_main_window import root , setting_frame ,saved_account ,saved_folder ,open_settings_window




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

status_label = tk.Label(root, text="準備完了。設定を確認し、ボタンを押してください。", fg="black", padx=10, pady=15, font=("Arial", 10))
status_label.pack(fill='x')

main_button_frame = Frame(root)
main_button_frame.pack(pady=5)

read_button = tk.Button(
    main_button_frame, 
    text="抽出を実行", 
    # ★ command の引数を修正: status_label と search_button を追加 ★
    command=lambda: run_extraction_workflow(root, account_entry, folder_entry, status_label, search_button), 
    bg="#4CAF50", fg="white", font=("Arial", 12, "bold"), width=20 
)
read_button.pack(side=tk.LEFT, padx=10)

search_button = tk.Button(
    main_button_frame, 
    text="検索・結果一覧表示", 
    command=lambda: open_search_window(root), 
    bg="#2196F3", fg="white", font=("Arial", 12), width=20 
)
search_button.pack(side=tk.LEFT, padx=10)

output_csv_path = os.path.join(SCRIPT_DIR, OUTPUT_CSV_FILE)
is_csv_exists_and_not_empty = False
if os.path.exists(output_csv_path):
    try:
        if not pd.read_csv(output_csv_path, encoding='utf-8-sig').empty:
            is_csv_exists_and_not_empty = True
    except Exception:
        pass 

if not is_csv_exists_and_not_empty:
    search_button.config(state=tk.DISABLED)

settings_button_frame = Frame(root)
settings_button_frame.pack(pady=(0, 5)) 

settings_button_label = tk.Label(settings_button_frame, text="アカウント名の初期設定はこちらから", fg="gray")
settings_button_label.pack(side=tk.TOP, pady=(5, 0))

settings_button = tk.Button(
    settings_button_frame, 
    text="設定", 
    command=lambda: open_settings_window(root), 
    bg="#808080", fg="white", font=("Arial", 10), width=15
)
settings_button.pack(side=tk.TOP, padx=10, pady=5)

root.mainloop()

