# gui_main_windowgui_main_window.py 
#開いてすぐの画面のフレームの設定

from tkinter import Frame
import tkinter as tk
from tkinter import messagebox
from gui_utils import load_config_csv ,save_config_csv


root = tk.Tk()
root.title("スキルシートデータ化システム")    
WINDOW_WIDTH = 500; WINDOW_HEIGHT = 380
root.update_idletasks()
screen_width = root.winfo_screenwidth(); screen_height = root.winfo_screenheight()
x = int((screen_width / 2) - 250); y = int((screen_height / 2) - 190)
root.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}+{x}+{y}")
# 以前に保存されたアカウント情報とフォルダパスを読み込む（name.csvから）
saved_account, saved_folder = load_config_csv()

setting_frame = Frame(root, padx=10, pady=10, borderwidth=1, relief="solid")
setting_frame.pack(pady=10, padx=10, fill='x')
tk.Label(setting_frame, text="対象フォルダパス (必須):").grid(row=1, column=0, sticky="w", pady=5)
folder_entry = tk.Entry(setting_frame, width=40) 
folder_entry.insert(0, saved_folder) 
folder_entry.grid(row=1, column=1, sticky="ew", padx=5)

tk.Label(setting_frame, text="入力例: 受信トレイ\\プロジェクトXのスキルシート", fg="gray", font=("Arial", 9)).grid(row=2, column=1, sticky="w", padx=5)

status_label = tk.Label(root, text="準備完了。設定を確認し、ボタンを押してください。", fg="black", padx=10, pady=15, font=("Arial", 10))
status_label.pack(fill='x')

main_button_frame = Frame(root)
main_button_frame.pack(pady=5)

tk.Label(setting_frame, text="Outlookアカウント (メールアドレス/表示名) (必須):").grid(row=0, column=0, sticky="w", pady=5)
account_entry = tk.Entry(setting_frame, width=40)
account_entry.insert(0, saved_account) 
account_entry.grid(row=0, column=1, sticky="ew", padx=5)
setting_frame.grid_columnconfigure(1, weight=1)

#設定ボタンの押した後の画面
def open_settings_window(root):
    settings_window = tk.Toplevel(root)
    settings_window.title("アカウント名初期設定")
    window_width = 400; window_height = 150
    screen_width = root.winfo_screenwidth(); screen_height = root.winfo_screenheight()
    x = int((screen_width / 2) - 200); y = int((screen_height / 2) - 75)
    settings_window.geometry(f"{window_width}x{window_height}+{x}+{y}")
    tk.Label(settings_window, text="名前の初期設定が可能です。", font=("Arial", 10, "bold")).pack(pady=10)
    current_account, _ = load_config_csv()
    tk.Label(settings_window, text="Outlookアカウント (メールアドレス):").pack(padx=10, pady=(5, 0), anchor='w')
    
    #email_entry: 設定ウィンドウ内のアカウント入力フィールド
    email_entry = tk.Entry(settings_window, width=50) 
    email_entry.insert(0, current_account)
    email_entry.pack(padx=10, fill='x')
    
    #設定ウィンドウで「保存して閉じる」ボタンが押された時の処理。
    def save_and_close():
        nonlocal email_entry 
        
        new_account = email_entry.get().strip()
        if not new_account: messagebox.showerror("エラー", "メールアドレスは空にできません。"); return
        
        global account_entry, status_label
        #新しいアカウント名をCSVに保存
        success, message = save_config_csv(new_account)
        if success:
            # メイン画面の account_entry を更新
            account_entry.delete(0, tk.END); account_entry.insert(0, new_account)
            status_label.config(text=f" 設定: アカウント名が '{new_account}' に更新されました。", fg="purple")
            messagebox.showinfo("保存完了", message); settings_window.destroy()
        else:
            messagebox.showerror("保存エラー", message)
            
    save_button = tk.Button(settings_window, text="上書き保存", command=save_and_close, bg="#FFA07A")
    save_button.pack(pady=10)


