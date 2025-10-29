# gui_search_window.py (create_sample_data 削除版)

import tkinter as tk
from tkinter import ttk
# 📌 修正: messagebox をインポート
from tkinter import messagebox
import pandas as pd
import os
# 📌 修正: 循環インポート解消のためコメントアウト (前回修正済み)
# import main_application
from config import OUTPUT_CSV_FILE as OUTPUT_FILENAME

# ==============================================================================
# 0. 共通ユーティリティ（データ処理ロジック）
# ==============================================================================

# 📌 削除: create_sample_data 関数を削除
# def create_sample_data():
#     """ (削除) """
#     # ... (関数定義全体を削除) ...

def filter_skillsheets_by_keywords(df: pd.DataFrame, keywords: list) -> pd.DataFrame:
    # ... (変更なし) ...
    if df.empty or not keywords: return df
    search_cols = [col for col in df.columns if col in ['スキル','件名','本文','添付ファイル内容']]
    if not search_cols: return df # 検索対象列がない場合はそのまま返す
    df_search = df[search_cols].astype(str).fillna(' ').agg(' '.join, axis=1).str.lower()
    filter_condition = pd.Series([True] * len(df), index=df.index)
    for keyword in keywords:
        lower_keyword = keyword.lower().strip()
        if lower_keyword:
            filter_condition = filter_condition & df_search.str.contains(lower_keyword, na=False)
    return df[filter_condition]


def filter_skillsheets(df: pd.DataFrame, keywords: list, range_data: dict) -> pd.DataFrame:
    # ... (変更なし) ...
    if df.empty: return df # 空のDataFrameなら即座に返す
    df_filtered = df.copy()
    df_filtered = filter_skillsheets_by_keywords(df_filtered, keywords)
    if df_filtered.empty: return df_filtered
    for key, limits in range_data.items():
        lower = limits['lower']
        upper = limits['upper']
        if not lower and not upper: continue
        col_name = {'age': '年齢', 'price': '単価', 'start': '実働開始'}.get(key)
        
        # 抽出結果に列が存在しない場合スキップ
        if col_name not in df_filtered.columns: continue

        if col_name in ['年齢', '単価']:
            try:
                col = df_filtered[col_name]
                col_numeric = pd.to_numeric(col, errors='coerce') 
                is_not_nan = col_numeric.notna()
                # min/max は NaN を除外して計算
                min_val = col_numeric.min() if is_not_nan.any() else 0
                max_val = col_numeric.max() if is_not_nan.any() else float('inf')
                
                lower_val = int(lower) if lower and str(lower).isdigit() else min_val
                upper_val = int(upper) if upper and str(upper).isdigit() else max_val
                
                valid_range_filter = is_not_nan & (col_numeric >= lower_val) & (col_numeric <= upper_val)
                # filter_condition = valid_range_filter | (~is_not_nan) # NaNも常に含める場合
                df_filtered = df_filtered[valid_range_filter] # NaNを除外する場合
            except Exception as e:
                print(f"🚨 データ型エラー: '{col_name}'の入力値またはデータが無効です。{e}")
                continue
                
        elif key == 'start' and '実働開始' in df_filtered.columns:
            start_col = df_filtered['実働開始']
            is_nan_or_nat = pd.to_datetime(start_col, errors='coerce').isna() # NaT も NaN として扱う
            
            df_target = df_filtered[~is_nan_or_nat].copy()
            if df_target.empty: # 全て NaN/NaT ならスキップ
                 df_filtered = df_target # 空にする
                 continue 
                 
            start_col_target_str = df_target['実働開始'].astype(str).str.replace(r'[^0-9]', '', regex=True) # 数字のみ抽出
            
            filter_condition = pd.Series([True] * len(df_target), index=df_target.index)
            if lower: 
                # YYYYMM形式の文字列比較
                lower_norm = str(lower).replace(r'[^0-9]', '', regex=True)
                filter_condition = filter_condition & (start_col_target_str >= lower_norm)
            if upper:
                upper_norm = str(upper).replace(r'[^0-9]', '', regex=True)
                filter_condition = filter_condition & (start_col_target_str <= upper_norm)
                
            # NaNだった行はフィルタリング条件に関わらず除外される
            df_filtered = df_target[filter_condition]
            
    return df_filtered


# ==============================================================================
# 1. メインアプリケーション（データと画面遷移の管理）
# ==============================================================================

class App(tk.Toplevel):
    """メインウィンドウとアプリケーションの状態を管理するクラス"""
    
    # 📌 修正: __init__ シグネチャ変更 (前回修正済み)
    def __init__(self, parent, data_frame: pd.DataFrame, open_email_callback):
        super().__init__(parent) 
        self.master = parent 
        self.open_email_callback = open_email_callback
        self.title("スキルシート検索アプリ")
        
        # --- 属性の初期化 (変更なし) ---
        self.keywords = []      
        self.range_data = {'age': {'lower': '', 'upper': ''}, 'price': {'lower': '', 'upper': ''}, 'start': {'lower': '', 'upper': ''}} 
        self.all_cands = {
            'age': [str(i) for i in range(20, 71, 5)], 
            'price': [str(i) for i in range(50, 101, 10)],
            'start': ['202401', '202404', '202407', '202410', '202501', '202504']
        }
        # 📌 修正: _clean_data が空のDataFrameを返す可能性あり
        self.df_all_skills = self._clean_data(data_frame) 
        self.df_filtered_skills = self.df_all_skills.copy() if not self.df_all_skills.empty else pd.DataFrame()
        
        self.current_frame = None 
        self.screen1 = None
        self.screen2 = None
        
        # --- ウィンドウサイズと位置の設定 (変更なし) ---
        window_width = 900
        window_height = 700
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        center_x = int(screen_width/2 - window_width/2)
        center_y = int(screen_height/2 - window_height/2)
        self.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        # 最初の画面表示
        # 📌 df_all_skills が空の場合の処理を追加しても良い (例: Screen1 表示前にエラーを出すなど)
        if self.df_all_skills.empty:
            # messagebox.showerror("起動エラー", "表示できるデータがありません。")
            # self.on_return_to_main() # メイン画面に戻る
            # または、空の Screen1 を表示する (現状維持)
             pass

        self.show_screen1()
        self.protocol("WM_DELETE_WINDOW", self.on_closing) 
        #self.grab_set()

    def on_closing(self):
        # ... (変更なし) ...
        self.grab_release() 
        try: self.master.destroy() 
        except tk.TclError: pass 
        try: self.destroy()
        except tk.TclError: pass
            
    def on_return_to_main(self):
        # ... (変更なし) ...
        self.grab_release()
        self.master.deiconify() 
        self.destroy()

    def _clean_data(self, df: pd.DataFrame) -> pd.DataFrame:
        """main_appから渡されたDataFrameをクリーンアップし、UIで使えるようにする"""
        if df.empty: return pd.DataFrame() # 最初から空なら空を返す
        try:
            df.columns = df.columns.str.strip()
            rename_map = {
                '単金': '単価', 
                'スキルor言語': 'スキル', 
                '名前': '氏名', 
                '期間_開始':'実働開始',
                '本文(テキスト形式)':'本文',
                '本文(ファイル含む)':'添付ファイル内容',
                'メールURL': 'ENTRY_ID' # DBから読み込むときは 'メールURL'
            }
            # 'EntryID' も考慮 (DB保存時にindex=Trueで保存した場合)
            if 'EntryID' in df.columns and 'ENTRY_ID' not in df.columns:
                 df = df.rename(columns={'EntryID': 'ENTRY_ID'}, errors='ignore')
            elif 'メールURL' in df.columns and 'ENTRY_ID' not in df.columns:
                 df = df.rename(columns={'メールURL': 'ENTRY_ID'}, errors='ignore')


            if '期間_開始' in df.columns:
                df = df.rename(columns={'期間_開始': '実働開始'}, errors='ignore')
            elif '実働開始' not in df.columns:
                df['実働開始'] = 'N/A' # ない場合は列を追加
                
            # rename_map に基づくリネーム
            df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns and v != 'ENTRY_ID'}, errors='ignore')
            
            # ENTRY_ID 列の整形
            if 'ENTRY_ID' in df.columns:
                df['ENTRY_ID'] = df['ENTRY_ID'].astype(str).str.replace('outlook:', '', regex=False).str.strip()
                # IDとして有効そうな行のみ残す (例: 10文字以上)
                df = df[df['ENTRY_ID'].astype(str).str.len() > 10].reset_index(drop=True)
            else:
                # ENTRY_ID がないと動作しないためエラー
                raise ValueError("必要な 'ENTRY_ID' 列が見つかりません。")
            
            return df

        except Exception as e:
            # 📌 修正: エラー時にメッセージを表示し、空のDataFrameを返す
            print(f"🚨 エラー: データクリーンアップに失敗しました。詳細: {e}") 
            messagebox.showerror("データエラー", f"データの読み込みまたは整形に失敗しました。\n詳細: {e}\n\n空のテーブルを表示します。")
            return pd.DataFrame()

    def show_screen1(self):
        # ... (変更なし) ...
        if self.current_frame: self.current_frame.destroy()
        self.screen1 = Screen1(self)
        self.current_frame = self.screen1
        self.current_frame.grid(row=0, column=0, sticky='nsew')
        current_keywords_str = ", ".join(self.keywords)
        self.after(10, lambda: self._set_screen1_keywords(current_keywords_str))

    def _set_screen1_keywords(self, keywords_str):
        # ... (変更なし) ...
        if self.screen1 and hasattr(self.screen1, 'keyword_entry'):
            try:
                self.screen1.keyword_entry.delete(0, tk.END) 
                self.screen1.keyword_entry.insert(0, keywords_str)
            except tk.TclError: # ウィンドウが閉じられた後などのエラーを無視
                pass

    def show_screen2(self):
        """検索結果表示画面（Screen2）に遷移する。"""
        if self.current_frame: 
            if isinstance(self.current_frame, Screen1): 
                self.current_frame.save_state()
            self.current_frame.destroy()
        
        # 📌 修正: df_all_skills が空でないかチェック
        if not self.df_all_skills.empty:
            self.df_filtered_skills = filter_skillsheets(
                self.df_all_skills, self.keywords, self.range_data)
        else:
            self.df_filtered_skills = pd.DataFrame() # 空のDFを渡す
        
        self.screen2 = Screen2(self)
        self.current_frame = self.screen2
        self.current_frame.grid(row=0, column=0, sticky='nsew')


# ==============================================================================
# 2. 画面1: 検索条件の入力
# ==============================================================================

class Screen1(ttk.Frame):
    # ... (このクラスは変更なし) ...
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
        self.rowconfigure(8, weight=1) 
        button_frame = ttk.Frame(self)
        button_frame.grid(row=8, column=0, columnspan=2, padx=10, pady=10, sticky='ew')
        ttk.Button(button_frame, text="検索", command=master.show_screen2).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="抽出画面に戻る", command=self.master.on_return_to_main).pack(side=tk.LEFT, padx=5)
        self.rowconfigure(9, weight=1)

    def create_range_input(self, label_text, key, row):
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
        typed = combo.get().lower()
        all_candidates = self.master.all_cands.get(key, [])
        new_values = [item for item in all_candidates if item.lower().startswith(typed)]
        combo['values'] = new_values

    def save_state(self):
        new_keywords = [k.strip() for k in self.keyword_entry.get().split(',') if k.strip()]
        self.master.keywords = list(set(new_keywords))[:5]
        for key in ['age', 'price', 'start']:
            # ウィジェットが存在するか確認してから .get() を呼ぶ
            if key in self.lower_widgets and self.lower_widgets[key].winfo_exists():
                 self.master.range_data[key]['lower'] = self.lower_widgets[key].get().strip()
            if key in self.upper_widgets and self.upper_widgets[key].winfo_exists():
                 self.master.range_data[key]['upper'] = self.upper_widgets[key].get().strip()


# ==============================================================================
# 3. 画面2: タグ表示とTreeview
# ==============================================================================

class Screen2(ttk.Frame):
    # ... (このクラスの __init__ と open_email_from_entry 以外は変更なし) ...
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

        self.setup_treeview()
        self.display_search_results()
        
        button_frame = ttk.Frame(self)
        button_frame.grid(row=7, column=0, columnspan=2, padx=10, pady=(10, 0), sticky='ew')
        ttk.Button(button_frame, text="本文表示", command=lambda: self.update_display_area('本文')).pack(side='left', padx=(0, 10))
        self.btn_attachment_content = ttk.Button(
            button_frame, text="添付ファイル内容表示", 
            command=lambda: self.update_display_area('添付ファイル内容'), state='disabled'
        )
        self.btn_attachment_content.pack(side='left')
        ttk.Button(button_frame, text="戻る (検索条件へ)", command=master.show_screen1).pack(side='right', padx=10)
        
        self.body_text = tk.Text(self, wrap='word', height=10, state='disabled')
        self.body_text.grid(row=8, column=0, columnspan=2, padx=10, pady=(0, 10), sticky='nsew')


    def open_email_from_entry(self):
        """ID入力欄の値をENTRY_IDとして取得し、Appに保存されたコールバック関数を呼び出す。"""
        entry_id = self.id_entry.get().strip()
        # 📌 修正: self.master (Appインスタンス) 経由でコールバックを呼び出す
        if hasattr(self.master, 'open_email_callback') and callable(self.master.open_email_callback):
            self.master.open_email_callback(entry_id)
        else:
             print("エラー: open_email_callback が設定されていません。")
             messagebox.showerror("内部エラー", "Outlookを開く機能が正しく設定されていません。")


    def check_attachment_content(self, item_id):
        # ... (変更なし) ...
        if not item_id:
            self.btn_attachment_content.config(state='disabled')
            return
        is_content_available = False
        try:
            # Treeviewのカラム名リストを取得
            tree_columns = list(self.tree['columns'])
            if 'ENTRY_ID' not in tree_columns: return # ENTRY_ID列がなければ何もしない
            
            entry_id_col_index = tree_columns.index('ENTRY_ID')
            tree_values = self.tree.item(item_id, 'values')
            
            # tree_valuesが十分な長さを持っているか確認
            if len(tree_values) <= entry_id_col_index: return
            
            entry_id = tree_values[entry_id_col_index]
            
            # df_all_skillsが空でないか、'ENTRY_ID'列を持っているか確認
            if self.master.df_all_skills.empty or 'ENTRY_ID' not in self.master.df_all_skills.columns:
                 return

            content_row = self.master.df_all_skills[self.master.df_all_skills['ENTRY_ID'].astype(str) == str(entry_id)]
            
            if not content_row.empty and '添付ファイル内容' in content_row.columns:
                content = content_row['添付ファイル内容'].iloc[0]
                content_str = str(content).strip().lower()
                if pd.notna(content) and content_str not in ['', 'nan', 'n/a']:
                    is_content_available = True
        except (ValueError, IndexError, KeyError) as e: 
             print(f"check_attachment_content でエラー: {e}") # デバッグ用
             pass 
        if is_content_available:
            self.btn_attachment_content.config(state='normal') 
        else:
            self.btn_attachment_content.config(state='disabled') 

    def _debug_keyword_extraction(self, entry_id, body_row):
        # ... (変更なし) ...
        search_cols = ['スキル', '件名', '本文', '添付ファイル内容']
        keywords = self.master.keywords
        if not keywords or body_row.empty: return
        print("="*70)
        print(f"✅ ENTRY_ID: {entry_id} のキーワードヒット箇所を検索中...")
        for col_name in search_cols:
            if col_name not in body_row.columns: continue
            full_data = body_row[col_name].iloc[0]
            if pd.isna(full_data) or str(full_data).strip() == '': continue 
            full_text = str(full_data).replace('_x000D_', '\n') # 改行コードを変換
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
                    extracted_text = full_text[start_context:end_context].replace('\n', ' ') # 表示用に改行をスペースに
                    print(f"  - [{col_name}] キーワード '{keyword}'")
                    print(f"    -> 抽出: '{extracted_text}' (文字位置: {start_index})")
        print("="*70)


    def update_display_area(self, content_type):
        # ... (変更なし) ...
        selected_items = self.tree.selection()
        if not selected_items: return
        item_id = selected_items[0]
        display_text = "データを取得できませんでした。"
        full_text_content = ""
        entry_id = "" 
        body_row = pd.DataFrame() 
        try:
            # Treeviewのカラム名リストを取得し、'ENTRY_ID' のインデックスを確認
            tree_columns = list(self.tree['columns'])
            if 'ENTRY_ID' not in tree_columns: raise ValueError("TreeviewにENTRY_ID列がありません。")
            id_index = tree_columns.index('ENTRY_ID')
            
            tree_values = self.tree.item(item_id, 'values')
            if len(tree_values) <= id_index: raise IndexError("選択行の値リストが短すぎます。")
            
            entry_id = tree_values[id_index]
            
            # df_all_skillsのチェック
            if self.master.df_all_skills.empty or 'ENTRY_ID' not in self.master.df_all_skills.columns:
                 raise ValueError("元のDataFrameが空か、ENTRY_ID列がありません。")

            body_row = self.master.df_all_skills[self.master.df_all_skills['ENTRY_ID'].astype(str) == str(entry_id)]
            
            if not body_row.empty and content_type in body_row.columns:
                full_data = body_row[content_type].iloc[0]
                if pd.notna(full_data) and str(full_data).strip() != '':
                    full_text_content = str(full_data).replace('_x000D_', '\n') # 改行コード変換
                    display_text = full_text_content[:1000] # 1000文字制限
                    if len(full_text_content) > 1000:
                        display_text += "...\n\n[--- 1000文字以降は省略 ---]"
                else:
                    display_text = f"{content_type} のデータが空です。"
            else:
                 display_text = f"選択されたメールに '{content_type}' のデータが見つかりません。" # 列がない場合

            # キーワード抽出デバッグ呼び出し
            self._debug_keyword_extraction(entry_id, body_row)
            
        except (ValueError, IndexError, KeyError) as e:
            display_text = f"データ取得エラー: {e}"
            print(f"update_display_area でエラー: {e}") # デバッグ用

        self.body_text.config(state='normal') 
        self.body_text.delete(1.0, tk.END) 
        self.body_text.insert(tk.END, display_text)
        self.body_text.config(state='disabled')
        
    def draw_tags(self):
        # ... (変更なし) ...
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
        # ... (変更なし) ...
        tag_container = ttk.Frame(self.tag_frame, relief='solid', borderwidth=1)
        tag_container.pack(side='left', padx=(5, 0), pady=2)
        ttk.Label(tag_container, text=text, padding=(5, 2)).pack(side='left')
        if is_keyword:
            ttk.Button(tag_container, text='×', width=2, command=lambda k=text: self.remove_tag(k)).pack(side='right')

    def remove_tag(self, keyword):
        # ... (変更なし) ...
        if keyword in self.master.keywords:
            self.master.keywords.remove(keyword)
            self.draw_tags()
            if not self.master.df_all_skills.empty:
                 self.master.df_filtered_skills = filter_skillsheets(self.master.df_all_skills, self.master.keywords, self.master.range_data)
            else:
                 self.master.df_filtered_skills = pd.DataFrame()
            self.display_search_results()

    def apply_new_keywords(self):
        # ... (変更なし) ...
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
        
    def setup_treeview(self):
        # ... (変更なし) ...
        # 表示するカラムを動的に決定 (df_all_skills が空でない場合)
        if not self.master.df_all_skills.empty:
             cols_available = self.master.df_all_skills.columns.tolist()
             cols_to_display_base = ['受信日時','件名' ,'スキル', '年齢', '単価', '実働開始'] 
             # 利用可能な列のみを抽出順序を維持しつつ選択
             cols_to_display = [col for col in cols_to_display_base if col in cols_available]
             # ENTRY_ID は常に内部的に必要
             all_columns = ['ENTRY_ID'] + cols_to_display
        else:
             # データがない場合はデフォルトのカラム構造 (ENTRY_IDのみでも良い)
             cols_to_display = []
             all_columns = ['ENTRY_ID']

        self.tree = ttk.Treeview(self, columns=all_columns, show='headings')
        
        for col in cols_to_display:
            self.tree.heading(col, text=col)
            # 幅の設定 (デフォルト値を用意)
            width_val = 100 # デフォルト幅
            if col in ['年齢', '単価']: width_val = 40
            elif col in ['実働開始']: width_val = 50
            elif col in ['スキル','件名']: width_val = 150
            elif col == '受信日時': width_val = 80 # 少し狭く
            self.tree.column(col, width=width_val, anchor='w', stretch=True)

        # ENTRY_ID 列は非表示
        self.tree.column('ENTRY_ID', width=0, stretch=tk.NO) 
        self.tree.heading('ENTRY_ID', text='')
            
        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.grid(row=6, column=0, padx=10, pady=10, sticky='nsew')
        vsb.grid(row=6, column=1, sticky='nse', padx=(0, 10), pady=10)
        self.tree.bind('<Double-Button-1>', self.treeview_double_click)
        self.tree.bind('<<TreeviewSelect>>', lambda event: self.check_attachment_content(self.tree.focus()))
        
    def display_search_results(self):
        # ... (変更なし) ...
        for item in self.tree.get_children(): self.tree.delete(item)
        
        # df_filtered_skills が空でないか、必要な列があるか確認
        if self.master.df_filtered_skills.empty or not all(col in self.master.df_filtered_skills.columns for col in self.tree['columns']):
             print("表示するデータがないか、必要な列が不足しています。") # デバッグ用
             return 

        for row in self.master.df_filtered_skills.itertuples(index=False):
            values = []
            for col in self.tree['columns']:
                val = getattr(row, col, 'N/A') # 列が存在しない場合に備える
                
                # データ型の整形 (NaNやNoneを空文字にするなど)
                if pd.isna(val):
                     val = '' 
                elif col == '年齢' or col == '単価':
                    try: val = int(float(val)) # 整数に変換
                    except (ValueError, TypeError): val = str(val) # 変換失敗時は文字列
                elif col == '受信日時':
                     try: val = str(val).split(' ')[0] # 日付のみ
                     except: val = str(val) # 失敗時はそのまま
                else:
                     val = str(val) # 他は文字列

                values.append(val)
                
            try:
                self.tree.insert('', 'end', values=values)
            except Exception as e:
                print(f"🚨 Treeview挿入エラー: 行データ {values} の挿入に失敗しました: {e}")
                
    def search_by_id(self):
        # ... (変更なし) ...
        search_id = self.id_entry.get().strip()
        if not self.master.df_all_skills.empty and 'ENTRY_ID' in self.master.df_all_skills.columns:
            if not search_id:
                self.master.df_filtered_skills = filter_skillsheets(self.master.df_all_skills, self.master.keywords, self.master.range_data)
            else:
                self.master.df_filtered_skills = self.master.df_all_skills[
                    self.master.df_all_skills['ENTRY_ID'].astype(str).str.contains(search_id, case=False, na=False)
                ]
        else:
             self.master.df_filtered_skills = pd.DataFrame() # 元データがなければ空
             
        self.display_search_results()
        
    def treeview_double_click(self, event):
        # ... (変更なし) ...
        item_id = self.tree.identify_row(event.y)
        if not item_id: return
        self.tree.selection_set(item_id)
        self.copy_id_to_entry(item_id)
        self.update_display_area('本文') 

    def copy_id_to_entry(self, item_id):
        # ... (変更なし) ...
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
        except (ValueError, IndexError, tk.TclError): # TclErrorを追加
            pass

# ==============================================================================
# 4. 実行エントリポイント
# ==============================================================================

def main():
    """アプリケーションのメイン実行関数。この関数が呼び出されるとGUIが起動する。"""
    
    root = tk.Tk()
    root.withdraw() 
    
    # 📌 修正: create_sample_data() 削除に伴い、直接ダミーデータ作成
    df_dummy = pd.DataFrame({ 
         'ENTRY_ID': ['outlook:dummy1', 'outlook:dummy2'], # outlook: プレフィックス付きでテスト
         '受信日時': ['2025-10-29 10:00:00', '2025-10-29 09:00:00'],
         '件名': ['テスト件名1', 'テスト件名2'],
         'スキル': ['Python', 'Java'],
         '本文': ['本文1','本文2'],
         '添付ファイル内容': ['添付1',''],
         '年齢': [30, None],
         '単価': [60, 70],
         '実働開始': ['202501', ''] # Noneや空文字をテスト
    })
    
    # テスト実行時にもダミーのコールバック関数を渡す
    def dummy_open_email_callback(entry_id):
        print(f"--- [TEST CALLBACK] Outlookでメールを開きます: {entry_id} ---")
        messagebox.showinfo("テストコールバック", f"Outlookを開く関数が呼ばれました。\nID: {entry_id}")
        
    # App の呼び出しに open_email_callback を追加
    app = App(
        root, 
        data_frame=df_dummy, 
        open_email_callback=dummy_open_email_callback
    ) 
    app.mainloop()

if __name__ == "__main__":
    main()