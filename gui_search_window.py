# gui_search_window.py (バグ修正版)

import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import pandas as pd
import os
import sqlite3 # 📌 DB接続のために追加
from config import DATABASE_NAME # 📌 DB名を取得するために追加
import traceback # ← ★★★ この行を追加 ★★★
# import main_application # 循環インポート防止

# ==============================================================================
# 0. 共通ユーティリティ（データ処理ロジック）
# ==============================================================================

def filter_skillsheets_by_keywords(df: pd.DataFrame, keywords: list) -> pd.DataFrame:
    """
    DataFrameをキーワードで絞り込む（軽量版）。
    '本文' と '添付ファイル内容' は軽量読み込みでは存在しないため、
    存在する列 ('スキル', '件名') だけで検索する。
    """
    if df.empty or not keywords: return df
    
    # 📌 修正: 検索対象列を、軽量読み込みで存在する可能性のある列のみにする
    search_cols = [col for col in df.columns if col in ['スキル','件名']] 
    
    if not search_cols: return df # 検索対象列がない
    
    df_search = df[search_cols].astype(str).fillna(' ').agg(' '.join, axis=1).str.lower()
    filter_condition = pd.Series([True] * len(df), index=df.index)
    for keyword in keywords:
        lower_keyword = keyword.lower().strip()
        if lower_keyword:
            filter_condition = filter_condition & df_search.str.contains(lower_keyword, na=False)
    return df[filter_condition]


def filter_skillsheets(df: pd.DataFrame, keywords: list, range_data: dict) -> pd.DataFrame:
    # (変更なし)
    if df.empty: return df 
    df_filtered = df.copy()
    df_filtered = filter_skillsheets_by_keywords(df_filtered, keywords)
    if df_filtered.empty: return df_filtered
    for key, limits in range_data.items():
        lower = limits['lower']
        upper = limits['upper']
        if not lower and not upper: continue
        col_name = {'age': '年齢', 'price': '単価', 'start': '実働開始'}.get(key)
        
        if col_name not in df_filtered.columns: continue

        if col_name in ['年齢', '単価']:
            try:
                col = df_filtered[col_name]
                col_numeric = pd.to_numeric(col, errors='coerce') 
                is_not_nan = col_numeric.notna()
                min_val = col_numeric.min() if is_not_nan.any() else 0
                max_val = col_numeric.max() if is_not_nan.any() else float('inf')
                lower_val = int(lower) if lower and str(lower).isdigit() else min_val
                upper_val = int(upper) if upper and str(upper).isdigit() else max_val
                valid_range_filter = is_not_nan & (col_numeric >= lower_val) & (col_numeric <= upper_val)
                df_filtered = df_filtered[valid_range_filter]
            except Exception as e:
                print(f"🚨 データ型エラー: '{col_name}'の入力値またはデータが無効です。{e}")
                continue
                
        elif key == 'start' and '実働開始' in df_filtered.columns:
            start_col = df_filtered['実働開始']
            is_nan_or_nat = pd.to_datetime(start_col, errors='coerce').isna()
            df_target = df_filtered[~is_nan_or_nat].copy()
            if df_target.empty:
                 df_filtered = df_target
                 continue 
            start_col_target_str = df_target['実働開始'].astype(str).str.replace(r'[^0-9]', '', regex=True)
            filter_condition = pd.Series([True] * len(df_target), index=df_target.index)
            if lower: 
                lower_norm = str(lower).replace(r'[^0-9]', '', regex=True)
                filter_condition = filter_condition & (start_col_target_str >= lower_norm)
            if upper:
                upper_norm = str(upper).replace(r'[^0-9]', '', regex=True)
                filter_condition = filter_condition & (start_col_target_str <= upper_norm)
            df_filtered = df_target[filter_condition]
            
    return df_filtered


# ==============================================================================
# 1. メインアプリケーション（データと画面遷移の管理）
# ==============================================================================

class App(tk.Toplevel):
    # (変更なし)
    def __init__(self, parent, data_frame: pd.DataFrame, open_email_callback, db_has_new_data_var: tk.BooleanVar):
        super().__init__(parent) 
        self.master = parent 
        self.open_email_callback = open_email_callback
        self.db_has_new_data_var = db_has_new_data_var # ← ★ この行を追加 ★
        self.title("スキルシート検索アプリ")
        self.keywords = []      
        self.range_data = {'age': {'lower': '', 'upper': ''}, 'price': {'lower': '', 'upper': ''}, 'start': {'lower': '', 'upper': ''}} 
        self.all_cands = {
            'age': [str(i) for i in range(20, 71, 5)], 
            'price': [str(i) for i in range(50, 101, 10)],
            'start': ['202401', '202404', '202407', '202410', '202501', '202504']
        }
        self.df_all_skills = self._clean_data(data_frame) 
        self.df_filtered_skills = self.df_all_skills.copy() if not self.df_all_skills.empty else pd.DataFrame()
        self.current_frame = None 
        self.screen1 = None
        self.screen2 = None
        
        window_width = 900
        window_height = 700
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        center_x = int(screen_width/2 - window_width/2)
        center_y = int(screen_height/2 - window_height/2)
        self.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        if self.df_all_skills.empty:
             pass

        self.show_screen1()
        self.protocol("WM_DELETE_WINDOW", self.on_closing) 
        #self.grab_set()

    def on_closing(self):
        self.grab_release() 
        try: self.master.destroy() 
        except tk.TclError: pass 
        try: self.destroy()
        except tk.TclError: pass
            
    def on_return_to_main(self):
        self.grab_release()
        self.master.deiconify() 
        self.destroy()

    def _clean_data(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        main_appから渡された軽量DataFrameをクリーンアップする。
        (本文 と 添付ファイル内容 はこの時点では含まれていない)
        """
        if df.empty: return pd.DataFrame()
        try:
            df.columns = df.columns.str.strip()
            # 📌 修正: '本文' と '添付ファイル内容' のリネームを削除 (渡されないため)
            rename_map = {
                '単金': '単価', 
                'スキルor言語': 'スキル', 
                '名前': '氏名', 
                '期間_開始':'実働開始',
                # '本文(テキスト形式)':'本文', # 軽量ロードでは除外
                # '本文(ファイル含む)':'添付ファイル内容', # 軽量ロードでは除外
                'メールURL': 'ENTRY_ID'
            }
            if 'EntryID' in df.columns and 'ENTRY_ID' not in df.columns:
                 df = df.rename(columns={'EntryID': 'ENTRY_ID'}, errors='ignore')
            elif 'メールURL' in df.columns and 'ENTRY_ID' not in df.columns:
                 df = df.rename(columns={'メールURL': 'ENTRY_ID'}, errors='ignore')

            if '期間_開始' in df.columns:
                df = df.rename(columns={'期間_開始': '実働開始'}, errors='ignore')
            elif '実働開始' not in df.columns:
                df['実働開始'] = 'N/A' 
                
            df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns and v != 'ENTRY_ID'}, errors='ignore')
            
            if 'ENTRY_ID' in df.columns:
                df['ENTRY_ID'] = df['ENTRY_ID'].astype(str).str.replace('outlook:', '', regex=False).str.strip()
                df = df[df['ENTRY_ID'].astype(str).str.len() > 10].reset_index(drop=True)
            else:
                raise ValueError("必要な 'ENTRY_ID' 列が見つかりません。")
            
            return df

        except Exception as e:
            print(f"🚨 エラー: データクリーンアップに失敗しました。詳細: {e}") 
            messagebox.showerror("データエラー", f"データの読み込みまたは整形に失敗しました。\n詳細: {e}\n\n空のテーブルを表示します。")
            return pd.DataFrame()

    def show_screen1(self):
        # (変更なし)
        if self.current_frame: self.current_frame.destroy()
        self.screen1 = Screen1(self)
        self.current_frame = self.screen1
        self.current_frame.grid(row=0, column=0, sticky='nsew')
        current_keywords_str = ", ".join(self.keywords)
        self.after(10, lambda: self._set_screen1_keywords(current_keywords_str))

    def _set_screen1_keywords(self, keywords_str):
        # (変更なし)
        if self.screen1 and hasattr(self.screen1, 'keyword_entry'):
            try:
                self.screen1.keyword_entry.delete(0, tk.END) 
                self.screen1.keyword_entry.insert(0, keywords_str)
            except tk.TclError:
                pass

    def show_screen2(self):
        # (変更なし)
        if self.current_frame: 
            if isinstance(self.current_frame, Screen1): 
                self.current_frame.save_state()
            self.current_frame.destroy()
        
        if not self.df_all_skills.empty:
            self.df_filtered_skills = filter_skillsheets(
                self.df_all_skills, self.keywords, self.range_data)
        else:
            self.df_filtered_skills = pd.DataFrame()
        
        self.screen2 = Screen2(self)
        self.current_frame = self.screen2
        self.current_frame.grid(row=0, column=0, sticky='nsew')


# ==============================================================================
# 2. 画面1: 検索条件の入力
# ==============================================================================

class Screen1(ttk.Frame):
    # (変更なし)
    def __init__(self, master):
        super().__init__(master)
        self.master = master
        self.lower_widgets = {} 
        self.upper_widgets = {} 
        self.columnconfigure(0, weight=1)
        self.columnconfigure(1, weight=1)
        ttk.Label(self, text="カンマ区切り（5個まで）：キーワード検索").grid(row=0, column=0, columnspan=2, padx=10, pady=(10, 0), sticky='w')
        self.keyword_entry = ttk.Entry(self) 
        self.keyword_entry.grid(row=1, column=0, columnspan=2, padx=10, pady=(0, 10), sticky='ew')
        ttk.Label(self, text="単価 (万円) 範囲指定").grid(row=2, column=0, columnspan=2, padx=10, pady=(10, 0), sticky='w')
        self.create_range_input('単価 (万円) 範囲指定', 'price', row=2)
        ttk.Label(self, text="年齢 (歳) 範囲指定").grid(row=4, column=0, columnspan=2, padx=10, pady=(10, 0), sticky='w')
        self.create_range_input('年齢 (歳) 範囲指定', 'age', row=4)
        ttk.Label(self, text="実働開始 範囲指定 (YYYYMM)").grid(row=6, column=0, columnspan=2, padx=10, pady=(10, 0), sticky='w')
        self.create_range_input('実働開始 範囲指定 (YYYYMM)', 'start', row=6)
        
        # --- ▼▼▼ ここから修正 ▼▼▼ ---
        
        # 伸縮する空きスペース (ボタンフレームの上)
        self.rowconfigure(8, weight=1) 
        
        # --- ボタンフレーム ---
        button_frame = ttk.Frame(self)
        # 📌 修正: ボタンフレームを row=9 に配置 (row=7 が入力欄のため)
        button_frame.grid(row=9, column=0, columnspan=2, padx=10, pady=10, sticky='sew') 

        # 📌 修正: ボタンフレーム内の列設定を変更
        button_frame.columnconfigure(0, weight=0) # 列0: 「戻る」ボタン用
        button_frame.columnconfigure(1, weight=1) # 列1: 伸縮する空きスペース
        button_frame.columnconfigure(2, weight=0) # 列2: 「リセット」ボタン用
        button_frame.columnconfigure(3, weight=0) # 列3: 「検索」ボタン用

        # 抽出画面に戻るボタン (列0, 左下寄せ)
        ttk.Button(button_frame, text="抽出画面に戻る", command=self.master.on_return_to_main).grid(row=0, column=0, padx=5, sticky='sw')

        # リセットボタン (列2, 右下寄せ)
        ttk.Button(button_frame, text="リセット", command=self.reset_fields).grid(row=0, column=2, padx=5, sticky='se')

        # 検索ボタン (列3, 右下寄せ)
        ttk.Button(button_frame, text="検索", command=master.show_screen2).grid(row=0, column=3, padx=5, sticky='se')
        
        # 📌 修正: 伸縮する空きスペース (ボタンフレームの下)
        self.rowconfigure(10, weight=1)
        # --- ▲▲▲ 修正ここまで ▲▲▲ ---

    def create_range_input(self, label_text, key, row):
        # (変更なし)
        is_combobox = (key != 'start')
        ttk.Label(self, text="下限:").grid(row=row+1, column=0, padx=(10, 0), pady=5, sticky='w')
        if is_combobox:
            widget_lower = ttk.Combobox(self, values=self.master.all_cands.get(key, []))
            widget_lower.bind('<KeyRelease>', lambda e, k=key, c=widget_lower: self.update_combobox_list(e, k, c))
        else:
            widget_lower = ttk.Entry(self)
        widget_lower.grid(row=row+1, column=0, padx=(50, 10), pady=5, sticky='ew')
        initial_lower_val = self.master.range_data[key]['lower']
        widget_lower.insert(0, initial_lower_val)
        self.lower_widgets[key] = widget_lower 
        ttk.Label(self, text="上限:").grid(row=row+1, column=1, padx=(10, 0), pady=5, sticky='w')
        if is_combobox:
            widget_upper = ttk.Combobox(self, values=self.master.all_cands.get(key, []))
            widget_upper.bind('<KeyRelease>', lambda e, k=key, c=widget_upper: self.update_combobox_list(e, k, c))
        else:
            widget_upper = ttk.Entry(self)
        widget_upper.grid(row=row+1, column=1, padx=(50, 10), pady=5, sticky='ew')
        initial_upper_val = self.master.range_data[key]['upper']
        widget_upper.insert(0, initial_upper_val)
        self.upper_widgets[key] = widget_upper
        
    def update_combobox_list(self, event, key, combo):
        # (変更なし)
        typed = combo.get().lower()
        all_candidates = self.master.all_cands.get(key, [])
        new_values = [item for item in all_candidates if item.lower().startswith(typed)]
        combo['values'] = new_values

    def save_state(self):
        # (変更なし)
        new_keywords = [k.strip() for k in self.keyword_entry.get().split(',') if k.strip()]
        self.master.keywords = list(set(new_keywords))[:5]
        for key in ['age', 'price', 'start']:
            if key in self.lower_widgets and self.lower_widgets[key].winfo_exists():
                 self.master.range_data[key]['lower'] = self.lower_widgets[key].get().strip()
            if key in self.upper_widgets and self.upper_widgets[key].winfo_exists():
                 self.master.range_data[key]['upper'] = self.upper_widgets[key].get().strip()
                 
    def reset_fields(self):
        # (変更なし)
        self.keyword_entry.delete(0, tk.END)
        for key in ['age', 'price', 'start']:
            if key in self.lower_widgets and self.lower_widgets[key].winfo_exists():
                 if isinstance(self.lower_widgets[key], ttk.Combobox):
                      self.lower_widgets[key].set('')
                 else:
                      self.lower_widgets[key].delete(0, tk.END) 
            if key in self.upper_widgets and self.upper_widgets[key].winfo_exists():
                 if isinstance(self.upper_widgets[key], ttk.Combobox):
                      self.upper_widgets[key].set('')
                 else:
                      self.upper_widgets[key].delete(0, tk.END)
        self.master.keywords = []
        self.master.range_data = {'age': {'lower': '', 'upper': ''}, 'price': {'lower': '', 'upper': ''}, 'start': {'lower': '', 'upper': ''}}
        print("INFO: 検索条件をリセットしました。") 


# ==============================================================================
# 3. 画面2: タグ表示とTreeview
# ==============================================================================

class Screen2(ttk.Frame):
    
    def __init__(self, master):
        super().__init__(master)
        self.master = master
        self.columnconfigure(0, weight=1) 
        self.rowconfigure(6, weight=3) # Treeview
        self.rowconfigure(8, weight=1) # Text area
        
        ttk.Label(self, text="追加のキーワード検索:").grid(row=0, column=0, columnspan=2, padx=10, pady=(10, 0), sticky='w')
        self.add_keyword_entry = ttk.Entry(self)
        self.add_keyword_entry.grid(row=1, column=0, padx=10, pady=(10, 0), sticky='ew')
        ttk.Button(self, text="適応", command=self.apply_new_keywords).grid(row=1, column=1, padx=10, pady=(10, 0), sticky='e')
        
        self.tag_frame = ttk.Frame(self)
        self.tag_frame.grid(row=2, column=0, columnspan=2, padx=10, pady=5, sticky='w')
        self.draw_tags()

        ttk.Label(self, text="IDからメールをOutlookで開く:").grid(row = 3, column=0, columnspan=2, padx=10, pady=(10, 0), sticky='w')
        self.id_entry = ttk.Entry(self)
        self.id_entry.grid(row = 4,column=0, padx=10, pady=5, sticky='ew')
        ttk.Button(self, text="Outlookで開く", command=self.open_email_from_entry).grid(row=4, column=1, padx=10, pady=5, sticky='e')

        self.setup_treeview() # 修正あり
        self.display_search_results()
        
        # --- ▼▼▼【ここから修正】ボタンフレームのレイアウトを .grid() に変更 ▼▼▼ ---
        button_frame = ttk.Frame(self)
        button_frame.grid(row=7, column=0, columnspan=2, padx=10, pady=(10, 0), sticky='ew')
        
        # --- .grid() のための列設定 ---
        button_frame.columnconfigure(0, weight=0) # 本文表示
        button_frame.columnconfigure(1, weight=0) # 添付ファイル
        button_frame.columnconfigure(2, weight=0) # 一覧更新
        button_frame.columnconfigure(3, weight=1) # 伸縮する空きスペース
        button_frame.columnconfigure(4, weight=0) # 戻る
        # ---

        ttk.Button(button_frame, text="本文表示", 
                   command=lambda: self.update_display_area("本文(テキスト形式)")
        ).grid(row=0, column=0, sticky='w', padx=(0, 10)) # .grid() に変更
        
        self.btn_attachment_content = ttk.Button(
            button_frame, text="添付ファイル内容表示", 
            command=lambda: self.update_display_area("本文(ファイル含む)"), state='disabled'
        )
        self.btn_attachment_content.grid(row=0, column=1, sticky='w') # .grid() に変更
        
        # --- ▼▼▼【ここに追加】(20行) ▼▼▼ ---
        self.btn_refresh = ttk.Button(
            button_frame, 
            text="一覧更新", 
            command=self.refresh_data_from_db,
            state='disabled' # デフォルトは無効
        )
        self.btn_refresh.grid(row=0, column=2, sticky='w', padx=(10, 0))
        
        # 「旗」の状態が変わったら、ボタンの状態も変えるように設定
        if self.master.db_has_new_data_var:
            def update_refresh_button_state(*args):
                try:
                    if self.master.db_has_new_data_var.get(): # 旗が True (更新あり) なら
                        self.btn_refresh.config(state=tk.NORMAL) # ボタンを有効化
                    else:
                        self.btn_refresh.config(state=tk.DISABLED) # 旗が False (最新) なら
                except tk.TclError:
                    pass # ウィンドウが閉じた後など
            
            # 旗 (BooleanVar) の変更を監視
            self.master.db_has_new_data_var.trace_add("write", update_refresh_button_state)
            update_refresh_button_state() # 初期状態をセット
        # --- ▲▲▲ 追加ここまで ▲▲▲ ---
        ttk.Button(button_frame, text="戻る (検索条件へ)", command=master.show_screen1
        ).grid(row=0, column=4, sticky='e', padx=10) # .grid() に変更
        # --- ▲▲▲ 修正ここまで ▲▲▲ ---
        self.body_text = tk.Text(self, wrap='word', height=10, state='disabled')
        self.body_text.grid(row=8, column=0, columnspan=2, padx=10, pady=(0, 10), sticky='nsew')


    def open_email_from_entry(self):
        # (変更なし)
        entry_id = self.id_entry.get().strip()
        if hasattr(self.master, 'open_email_callback') and callable(self.master.open_email_callback):
            self.master.open_email_callback(entry_id)
        else:
             print("エラー: open_email_callback が設定されていません。")
             messagebox.showerror("内部エラー", "Outlookを開く機能が正しく設定されていません。")

    # --- ▼▼▼ 修正: バグ2対応 (check_attachment_content) ▼▼▼ ---
    def check_attachment_content(self, item_id):
        """
        Treeviewで選択された行の 'Attachments' 列 (非表示) を読み取り、
        値があればボタンを有効化する。
        """
        if not item_id:
            self.btn_attachment_content.config(state='disabled')
            return
        
        is_content_available = False
        try:
            tree_columns = list(self.tree['columns'])
            
            # 'Attachments' 列が Treeview に含まれているか確認
            if 'Attachments' not in tree_columns:
                 self.btn_attachment_content.config(state='disabled')
                 return 
                 
            attachments_col_index = tree_columns.index('Attachments')
            tree_values = self.tree.item(item_id, 'values')
            
            if len(tree_values) <= attachments_col_index: return
            
            attachments_data = tree_values[attachments_col_index] 
            
            # 'Attachments' 列にファイル名(N/Aや空以外) があれば有効化
            if attachments_data and str(attachments_data).strip() not in ['', 'N/A']:
                is_content_available = True
                
        except (ValueError, IndexError, KeyError) as e: 
             print(f"check_attachment_content でエラー: {e}")
             pass 
             
        if is_content_available:
            self.btn_attachment_content.config(state='normal') 
        else:
            self.btn_attachment_content.config(state='disabled') 
    # --- ▲▲▲ 修正ここまで ▲▲▲ ---

    def _debug_keyword_extraction(self, entry_id, col_name, text_content):
        # (変更なし)
        keywords = self.master.keywords
        if not keywords or not text_content:
            return
        print("="*70)
        print(f"✅ ENTRY_ID: {entry_id} の [{col_name}] ヒット箇所検索:")
        full_text = str(text_content).replace('_x000D_', '\n')
        full_text_lower = full_text.lower()
        for keyword in keywords:
            lower_keyword = keyword.lower()
            if not lower_keyword: continue
            current_search_pos = 0
            while True:
                start_index = full_text_lower.find(lower_keyword, current_search_pos)
                if start_index == -1: break
                end_index = start_index + len(lower_keyword)
                current_search_pos = end_index
                start_context = max(0, start_index - 3)
                end_context = min(len(full_text), end_index + 3)
                extracted_text = full_text[start_context:end_context].replace('\n', ' ')
                print(f"  - '{keyword}' -> '{extracted_text}' ({start_index})")
        print("="*70)

    def update_display_area(self, content_type: str):
        # (変更なし - DBオンデマンド読み込み)
        selected_items = self.tree.selection()
        if not selected_items: return
        item_id = selected_items[0]
        display_text = "[データ取得中...]"
        entry_id = ""
        self.body_text.config(state='normal') 
        self.body_text.delete(1.0, tk.END) 
        self.body_text.insert(tk.END, display_text)
        self.body_text.config(state='disabled')
        self.master.update_idletasks() 

        try:
            tree_columns = list(self.tree['columns'])
            if 'ENTRY_ID' not in tree_columns:
                raise ValueError("TreeviewにENTRY_ID列がありません。")
            id_index = tree_columns.index('ENTRY_ID')
            tree_values = self.tree.item(item_id, 'values')
            if len(tree_values) <= id_index:
                raise IndexError("選択行の値リストが短すぎます。")
            entry_id = str(tree_values[id_index])
            if not entry_id or entry_id == 'N/A':
                 raise ValueError("有効な EntryID が取得できませんでした。")

            db_path = os.path.abspath(DATABASE_NAME)
            if not os.path.exists(db_path):
                 raise FileNotFoundError(f"データベース {DATABASE_NAME} が見つかりません。")
            
            conn = None
            text_content = ""
            try:
                conn = sqlite3.connect(db_path)
                cursor = conn.cursor()
                # 📌 修正: allowed_cols に '本文' (古い名前) も含めておく (安全策)
                allowed_cols = ["本文(テキスト形式)", "本文(ファイル含む)", "スキル", "件名", "本文"] 
                if content_type not in allowed_cols:
                     raise ValueError(f"不正なカラム名 {content_type} が指定されました。")
                
                query = f"SELECT \"{content_type}\" FROM emails WHERE \"EntryID\" = ?"
                cursor.execute(query, (entry_id,))
                row = cursor.fetchone()
                
                if row:
                    full_data = row[0]
                    if pd.notna(full_data) and str(full_data).strip() != '':
                        full_text_content = str(full_data).replace('_x000D_', '\n')
                        display_text = full_text_content[:1000]
                        if len(full_text_content) > 1000:
                            display_text += "...\n\n[--- 1000文字以降は省略 ---]"
                        self._debug_keyword_extraction(entry_id, content_type, full_text_content)
                    else:
                        display_text = f"{content_type} のデータが空です。"
                else:
                    display_text = f"データベースで EntryID '{entry_id}' が見つかりません。"
            except Exception as db_err:
                 print(f"DB読み込みエラー (update_display_area): {db_err}")
                 display_text = f"データベースからのテキスト取得中にエラーが発生しました。\n詳細: {db_err}"
            finally:
                if conn: conn.close()
        except (ValueError, IndexError, FileNotFoundError) as e:
            display_text = f"データ取得エラー: {e}"
            print(f"update_display_area でエラー: {e}") 

        self.body_text.config(state='normal') 
        self.body_text.delete(1.0, tk.END) 
        self.body_text.insert(tk.END, display_text)
        self.body_text.config(state='disabled')
        
    def draw_tags(self):
        # (変更なし)
        for widget in self.tag_frame.winfo_children(): widget.destroy()
        for keyword in self.master.keywords: self.create_tag(keyword, is_keyword=True)
        range_map = {'age': '年齢', 'price': '単価', 'start': '実働開始'}
        for key, label in range_map.items():
            lower = self.master.range_data[key]['lower']
            upper = self.master.range_data[key]['upper']
            if lower or upper: 
                tag_text = f"{label}: {lower or '下限なし'}~{upper or '上限なし'}"
                self.create_tag(tag_text, is_keyword=False) 

    def create_tag(self, text, is_keyword):
        # (変更なし)
        tag_container = ttk.Frame(self.tag_frame, relief='solid', borderwidth=1)
        tag_container.pack(side='left', padx=(5, 0), pady=2)
        ttk.Label(tag_container, text=text, padding=(5, 2)).pack(side='left')
        if is_keyword:
            ttk.Button(tag_container, text='×', width=2, command=lambda k=text: self.remove_tag(k)).pack(side='right')

    def remove_tag(self, keyword):
        # (変更なし)
        if keyword in self.master.keywords:
            self.master.keywords.remove(keyword)
            self.draw_tags()
            if not self.master.df_all_skills.empty:
                 self.master.df_filtered_skills = filter_skillsheets(self.master.df_all_skills, self.master.keywords, self.master.range_data)
            else:
                 self.master.df_filtered_skills = pd.DataFrame()
            self.display_search_results()

    def apply_new_keywords(self):
        # (変更なし)
        new_input = [k.strip() for k in self.add_keyword_entry.get().split(',') if k.strip()]
        combined_keywords = self.master.keywords + new_input
        self.master.keywords = list(set(combined_keywords))[:5]
        self.draw_tags()
        self.add_keyword_entry.delete(0, 'end') 
        if not self.master.df_all_skills.empty:
            self.master.df_filtered_skills = filter_skillsheets(self.master.df_all_skills, self.master.keywords, self.master.range_data)
        else:
            self.master.df_filtered_skills = pd.DataFrame()
        self.display_search_results()
        
    # --- ▼▼▼ 修正: バグ2対応 (setup_treeview) ▼▼▼ ---
    def setup_treeview(self):
        if not self.master.df_all_skills.empty:
             cols_available = self.master.df_all_skills.columns.tolist()
             
             # 📌 修正: 'Attachments' を表示対象ベースリストに追加
             cols_to_display_base = ['受信日時','件名' ,'スキル', '年齢', '単価', '実働開始', 'Attachments'] 
             
             cols_to_display = [col for col in cols_to_display_base if col in cols_available]
             all_columns = ['ENTRY_ID'] + cols_to_display
        else:
             cols_to_display = []
             all_columns = ['ENTRY_ID']

        self.tree = ttk.Treeview(self, columns=all_columns, show='headings')
        
        for col in cols_to_display:
            self.tree.heading(col, text=col)
            width_val = 100
            if col in ['年齢', '単価']: width_val = 40
            elif col in ['実働開始']: width_val = 50
            elif col in ['スキル','件名']: width_val = 150
            elif col == '受信日時': width_val = 80
            
            # 📌 修正: 'Attachments' 列を非表示にする
            elif col == 'Attachments': width_val = 0 
            
            # 📌 修正: 'Attachments' 列は伸縮させない
            self.tree.column(col, width=width_val, anchor='w', stretch=(col != 'Attachments'))
            if col == 'Attachments':
                 self.tree.column(col, stretch=tk.NO)
                 
        self.tree.column('ENTRY_ID', width=0, stretch=tk.NO) 
        self.tree.heading('ENTRY_ID', text='')
            
        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.grid(row=6, column=0, padx=10, pady=10, sticky='nsew')
        vsb.grid(row=6, column=1, sticky='nse', padx=(0, 10), pady=10)
        self.tree.bind('<Double-Button-1>', self.treeview_double_click)
        self.tree.bind('<<TreeviewSelect>>', lambda event: self.check_attachment_content(self.tree.focus()))
    # --- ▲▲▲ 修正ここまで ▲▲▲ ---
        
    def display_search_results(self):
        # (変更なし)
        for item in self.tree.get_children(): self.tree.delete(item)
        if self.master.df_filtered_skills.empty or not all(col in self.master.df_filtered_skills.columns for col in self.tree['columns']):
             print("表示するデータがないか、必要な列が不足しています。") 
             return 
        for row in self.master.df_filtered_skills.itertuples(index=False):
            values = []
            for col in self.tree['columns']:
                val = getattr(row, col, 'N/A') 
                if pd.isna(val): val = '' 
                elif col == '年齢' or col == '単価':
                    try: val = int(float(val))
                    except (ValueError, TypeError): val = str(val)
                elif col == '受信日時':
                     try: val = str(val).split(' ')[0]
                     except: val = str(val)
                else: val = str(val)
                values.append(val)
            try:
                self.tree.insert('', 'end', values=values)
            except Exception as e:
                print(f"🚨 Treeview挿入エラー: 行データ {values} の挿入に失敗しました: {e}")
                
    def search_by_id(self):
        # (変更なし)
        search_id = self.id_entry.get().strip()
        if not self.master.df_all_skills.empty and 'ENTRY_ID' in self.master.df_all_skills.columns:
            if not search_id:
                self.master.df_filtered_skills = filter_skillsheets(self.master.df_all_skills, self.master.keywords, self.master.range_data)
            else:
                self.master.df_filtered_skills = self.master.df_all_skills[
                    self.master.df_all_skills['ENTRY_ID'].astype(str).str.contains(search_id, case=False, na=False)
                ]
        else:
             self.master.df_filtered_skills = pd.DataFrame()
        self.display_search_results()
        
    # --- ▼▼▼ 修正: バグ1対応 (treeview_double_click) ▼▼▼ ---
    def treeview_double_click(self, event):
        item_id = self.tree.identify_row(event.y)
        if not item_id: return
        self.tree.selection_set(item_id)
        self.copy_id_to_entry(item_id)
        # 📌 修正: '本文' -> '本文(テキスト形式)' に変更
        self.update_display_area('本文(テキスト形式)') 
    # --- ▲▲▲ 修正ここまで ▲▲▲ ---

    def copy_id_to_entry(self, item_id):
        # (変更なし)
        try:
            tree_columns = list(self.tree['columns'])
            if 'ENTRY_ID' not in tree_columns: return
            id_index = tree_columns.index('ENTRY_ID')
            values = self.tree.item(item_id, 'values')
            if not values or id_index >= len(values): return
            id_value = str(values[id_index])
            self.master.clipboard_clear()
            self.master.clipboard_append(id_value)
            self.id_entry.delete(0, 'end')
            self.id_entry.insert('end', id_value)
        except (ValueError, IndexError, tk.TclError):
            pass
    # --- ▼▼▼【このメソッドを丸ごと追加】▼▼▼ ---
    # --- ▼▼▼【このメソッドを丸ごと置き換え】▼▼▼ ---
    def refresh_data_from_db(self):
        """
        データベースから最新の「軽量」データを再読み込みし、
        ソート、フィルタリングして Treeview を更新する。
        """
        
        # print("\n" + "="*40) # ログ削除
        # print("DEBUG: [Refresh] 「一覧更新」ボタンが押されました。")
        
        if hasattr(self, 'btn_refresh'):
            self.btn_refresh.config(state=tk.DISABLED)
        self.master.update_idletasks()

        try:
            db_path = os.path.abspath(DATABASE_NAME)
            if not os.path.exists(db_path):
                raise FileNotFoundError(f"データベース {DATABASE_NAME} が見つかりません。")
            
            conn = None
            new_df = pd.DataFrame()
            try:
                conn = sqlite3.connect(db_path)
                cursor = conn.cursor()
                
                cursor.execute("PRAGMA table_info(emails)")
                all_columns = [info[1] for info in cursor.fetchall()]
                heavy_columns = ['本文(テキスト形式)', '本文(ファイル含む)']
                light_columns = [col for col in all_columns if col not in heavy_columns]
                
                if not light_columns:
                     raise Exception("DBに列が見つかりません。")
                     
                light_columns_sql = ", ".join([f'"{col}"' for col in light_columns])
                query = f"SELECT {light_columns_sql} FROM emails"
                
                # print(f"DEBUG: [Refresh] DB読み込み実行: {query}")
                
                new_df = pd.read_sql_query(query, conn)
                
            finally:
                if conn: conn.close()

            # print(f"DEBUG: [Refresh] DBから {len(new_df)} 件の軽量データを読み込みました。")

            # --- ▼▼▼【修正】ここでソートを実行 ▼▼▼ ---
            if not new_df.empty and '受信日時' in new_df.columns:
                try:
                    new_df['受信日時'] = pd.to_datetime(new_df['受信日時'], errors='coerce')
                    new_df = new_df.sort_values(by='受信日時', ascending=False, na_position='last').reset_index(drop=True)
                    # print("DEBUG: [Refresh] 読み込みデータのソート完了。") 
                except Exception as sort_err:
                     print(f"警告: [Refresh] DB読み込み後のソート失敗: {sort_err}")
            # --- ▲▲▲ ソート追加ここまで ▲▲▲ ---

            # 5. App(self.master) のデータを更新
            self.master.df_all_skills = self.master._clean_data(new_df)
            
            # print(f"DEBUG: [Refresh] 現在のフィルタキーワード: {self.master.keywords}")

            # 6. 現在のフィルタ(Appが保持)を再適用
            self.master.df_filtered_skills = filter_skillsheets(
                self.master.df_all_skills, 
                self.master.keywords,
                self.master.range_data
            )
            
            # print(f"DEBUG: [Refresh] フィルタ適用後の件数: {len(self.master.df_filtered_skills)} 件")
            
            # 7. Treeview を再描画
            self.display_search_results()
            
            # 8. 「旗」を倒す
            if self.master.db_has_new_data_var:
                self.master.db_has_new_data_var.set(False)

            # 9. 画面下のテキストエリアもクリアする
            self.body_text.config(state='normal') 
            self.body_text.delete(1.0, tk.END) 
            self.body_text.insert(tk.END, f"一覧を更新しました。 (DB全件: {len(self.master.df_all_skills)} 件 / フィルタ後: {len(self.master.df_filtered_skills)} 件)")
            self.body_text.config(state='disabled')
            
            # print("INFO: 検索一覧をDBから更新しました。")

        except Exception as e:
            messagebox.showerror("更新エラー", f"一覧の更新中にエラーが発生しました。\n詳細: {e}")
            traceback.print_exc()
        finally:
            # 10. ボタンを再度有効化
            # (旗がFalseになったので、trace_add のコールバックで自動的に Disabled になる)
            if hasattr(self, 'btn_refresh'):
                try:
                    if self.btn_refresh.winfo_exists():
                        # self.btn_refresh.config(state=tk.NORMAL) # ← 不要
                        pass
                except tk.TclError:
                    pass
        
        # print("="*40 + "\n")
    # --- ▲▲▲ 置き換えここまで ▲▲▲ ---

# ==============================================================================
# 4. 実行エントリポイント
# ==============================================================================

def main():
    # (変更なし - 軽量読み込みに合わせてダミーデータを修正)
    root = tk.Tk()
    root.withdraw() 
    df_dummy = pd.DataFrame({ 
         'ENTRY_ID': ['outlook:dummy1', 'outlook:dummy2'], 
         '受信日時': ['2025-10-29 10:00:00', '2025-10-29 09:00:00'],
         '件名': ['テスト件名1', 'テスト件名2'],
         'スキル': ['Python', 'Java'],
         # 📌 本文 と 添付ファイル内容 は軽量読み込みで除外される
         # '本文(テキスト形式)': ['本文1','本文2'],
         # '本文(ファイル含む)': ['添付1',''],
         '年齢': [30, None],
         '単価': [60, 70],
         '実働開始': ['202501', ''],
         'Attachments': ['file1.xlsx', ''] # 📌 Attachments (ファイル名) は含まれる
    })
    
    def dummy_open_email_callback(entry_id):
        print(f"--- [TEST CALLBACK] Outlookでメールを開きます: {entry_id} ---")
        messagebox.showinfo("テストコールバック", f"Outlookを開く関数が呼ばれました。\nID: {entry_id}")
        
    app = App(
        root, 
        data_frame=df_dummy, 
        open_email_callback=dummy_open_email_callback
    ) 
    app.mainloop()

if __name__ == "__main__":
    main()