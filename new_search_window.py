#gui_new_search_window.py
import tkinter as tk
from tkinter import ttk
import pandas as pd
import numpy as np

# ==============================================================================
# 0. 共通ユーティリティ（テストデータ、検索ロジックなど）
# ==============================================================================
#このコードは使用しない
def create_sample_data(file_path="skillsheets.csv"):
    """テスト用のスキルシートデータを作成する"""
    data = {
        'ENTRY_ID': [f'ID{i:03}' for i in range(1, 11)],
        '氏名': [f'テスト太郎{i}' for i in range(1, 11)],
        'スキル': ['JAVA, Python, DB', 'C#, Azure', 'Python, AWS', 'JAVA, AWS, PG', 'C#, Unity', 
                 'Python, AI', 'DB, SQL', 'JAVA, DB', 'C#, .NET', 'Python, Django'],
        '本文': [f'これはメール本文{i}です。詳細情報や経歴はこの本文に記述されています。非常に長いメール本文を想定しています。ユーザーがダブルクリックした際、Treeviewの代わりにこの本文が下のテキストエリアに表示されます。' for i in range(1, 11)],
        '年齢': [25, 30, 45, 33, 28, 50, 40, 37, 22, 35], 
        '単価': [50, 65, 70, 55, 60, 80, 75, 50, 60, 70],
        '実働開始': ['2024/05', '2025/01', '2024/07', '2024/03', '2025/06', 
                   '2024/01', '2025/03', '2024/11', '2024/02', '2025/02'],
    }
    df = pd.DataFrame(data)
    df.to_csv(file_path, index=False)
    return df

def filter_skillsheets(df: pd.DataFrame, keywords: list, range_data: dict) -> pd.DataFrame:
    """キーワードと範囲指定の両方でフィルタリングを行う (エラー処理強化)"""
    df_filtered = df.copy()

    # 1. キーワードフィルタリング (AND条件)
    df_filtered = filter_skillsheets_by_keywords(df_filtered, keywords)
    if df_filtered.empty: return df_filtered

    # 2. 範囲指定フィルタリング
    for key, limits in range_data.items():
        lower = limits['lower']
        upper = limits['upper']
        
        if not lower and not upper:
            continue

        col_name = None
        if key == 'age': col_name = '年齢'
        elif key == 'price': col_name = '単価'
        elif key == 'start': col_name = '実働開始'

        if col_name in ['年齢', '単価']:
            # 数値フィルタリング (エラー処理をtry-exceptで実装)
            try:
                col = df_filtered[col_name]
                
                lower_val = int(lower) if lower else col.min()
                upper_val = int(upper) if upper else col.max()
                
                df_filtered = df_filtered[(col >= lower_val) & (col <= upper_val)]
            except ValueError:
                # 無効な入力（例: 'あいうえお'）があった場合
                print(f"🚨 データ型エラー: '{col_name}'の入力値が無効です ({lower} / {upper})。この項目はスキップします。")
                continue # 次の項目へ
            except KeyError:
                # 列名がDataFrameに存在しない場合
                print(f"🚨 KeyError: 列名 '{col_name}' がDataFrameに見つかりません。")
                continue

        elif key == 'start':
            # 実働開始 (ここでは文字列比較)
            if '実働開始' in df_filtered.columns:
                start_col = df_filtered['実働開始'].astype(str)
                if lower:
                    df_filtered = df_filtered[start_col >= lower]
                if upper:
                    df_filtered = df_filtered[start_col <= upper]

    return df_filtered

def filter_skillsheets_by_keywords(df: pd.DataFrame, keywords: list) -> pd.DataFrame:
    """キーワードによるAND検索を実行する"""
    if df.empty or not keywords:
        return df
    
    # 検索を効率化するため、検索対象の列を結合
    search_cols = [col for col in df.columns if col not in ['本文', '年齢', '単価']]
    df_search = df[search_cols].astype(str).agg(' '.join, axis=1).str.lower()
    
    filter_condition = pd.Series([True] * len(df), index=df.index)
    
    for keyword in keywords:
        lower_keyword = keyword.lower().strip()
        if lower_keyword:
            filter_condition = filter_condition & df_search.str.contains(lower_keyword, na=False)
            
    return df[filter_condition]

# ==============================================================================
# 1. メインアプリケーション（データと画面遷移の管理）
# ==============================================================================

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("スキルシート検索アプリ")
        
        # --- 🌟 ウィンドウを中央に配置するロジック 🌟 ---
        window_width = 900
        window_height = 700
        
        # 画面の幅と高さを取得
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        
        # 中央に配置するための座標を計算
        center_x = int(screen_width/2 - window_width/2)
        center_y = int(screen_height/2 - window_height/2)
        
        # ウィンドウサイズと位置を設定
        self.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
        # ------------------------------------------------

        
        # --- 共有データ ---
        self.keywords = []      
        self.range_data = {
            'age': {'lower': '', 'upper': ''}, 
            'price': {'lower': '', 'upper': ''}, 
            'start': {'lower': '', 'upper': ''}
        } 
        
        self.all_cands = {
            'age': [str(i) for i in range(20, 71, 5)], 
            'price': [str(i) for i in range(50, 101, 10)],
            'start': ['2024/01', '2024/04', '2024/07', '2024/10', '2025/01', '2025/04']
        }
        
        # --- データフレーム ---
        self.df_all_skills = create_sample_data() 
        self.df_filtered_skills = self.df_all_skills.copy()
        
        # --- 画面遷移 ---
        self.current_frame = None
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        self.show_screen1()

    def show_screen1(self):
        if self.current_frame: self.current_frame.destroy()
        self.current_frame = Screen1(self)
        self.current_frame.grid(row=0, column=0, sticky='nsew')

    def show_screen2(self):
        if self.current_frame: 
            self.current_frame.save_state()
            self.current_frame.destroy()
            
        self.df_filtered_skills = filter_skillsheets(
            self.df_all_skills, self.keywords, self.range_data)
        
        self.current_frame = Screen2(self)
        self.current_frame.grid(row=0, column=0, sticky='nsew')

# ==============================================================================
# 2. 画面1: 検索条件の入力
# ==============================================================================

class Screen1(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.master = master
        
        self.lower_vars = {}
        self.upper_vars = {}
        
        self.columnconfigure(0, weight=1)
        self.columnconfigure(1, weight=1)
        
        # --- キーワード検索欄（カンマ区切り） ---
        ttk.Label(self, text="カンマ区切り（5個まで）：キーワード検索").grid(
            row=0, column=0, columnspan=2, padx=10, pady=(10, 0), sticky='w')
        
        self.keyword_var = tk.StringVar(value=", ".join(master.keywords))
        self.keyword_entry = ttk.Entry(self, textvariable=self.keyword_var)
        self.keyword_entry.grid(row=1, column=0, columnspan=2, padx=10, pady=(0, 10), sticky='ew')
        
        # --- 範囲指定入力欄 (単価、年齢、実働開始) ---
        self.create_range_input('単価 (万円) 範囲指定', 'price', row=2)
        self.create_range_input('年齢 (歳) 範囲指定', 'age', row=4)
        self.create_range_input('実働開始 範囲指定 (YYYY年MM月)', 'start', row=6)

        # --- 検索ボタン ---
        ttk.Button(self, text="検索 (画面2へ)", command=master.show_screen2).grid(
            row=8, column=0, columnspan=2, padx=10, pady=20)

    def create_range_input(self, label_text, key, row):
        """下限・上限のCombobox または Entry を生成するヘルパー関数"""
        
        is_combobox = (key != 'start')
        
        ttk.Label(self, text=label_text).grid(row=row, column=0, columnspan=2, padx=10, pady=(10, 0), sticky='w')
        
        # --- 下限 ---
        ttk.Label(self, text="下限:").grid(row=row+1, column=0, padx=(10, 0), pady=5, sticky='w')
        
        # tk.StringVarをインスタンス変数に保存
        self.lower_vars[key] = tk.StringVar(value=self.master.range_data[key]['lower']) 
        lower_var = self.lower_vars[key]
        
        if is_combobox:
            widget_lower = ttk.Combobox(self, textvariable=lower_var, values=self.master.all_cands.get(key, []))
            widget_lower.bind('<KeyRelease>', lambda e, k=key, c=widget_lower: self.update_combobox_list(e, k, c))
        else:
            widget_lower = ttk.Entry(self, textvariable=lower_var)
            
        widget_lower.grid(row=row+1, column=0, padx=(50, 10), pady=5, sticky='ew')
        setattr(self, f'{key}_lower_entry', widget_lower) 

        # --- 上限 ---
        ttk.Label(self, text="上限:").grid(row=row+1, column=1, padx=(10, 0), pady=5, sticky='w')
        
        # tk.StringVarをインスタンス変数に保存
        self.upper_vars[key] = tk.StringVar(value=self.master.range_data[key]['upper'])
        upper_var = self.upper_vars[key]
        
        if is_combobox:
            widget_upper = ttk.Combobox(self, textvariable=upper_var, values=self.master.all_cands.get(key, []))
            widget_upper.bind('<KeyRelease>', lambda e, k=key, c=widget_upper: self.update_combobox_list(e, k, c))
        else:
            widget_upper = ttk.Entry(self, textvariable=upper_var)
            
        widget_upper.grid(row=row+1, column=1, padx=(50, 10), pady=5, sticky='ew')
        setattr(self, f'{key}_upper_entry', widget_upper) 

    def update_combobox_list(self, event, key, combo):
        """Comboboxのオートコンプリートロジック"""
        typed = combo.get().lower()
        all_candidates = self.master.all_cands.get(key, [])
        new_values = [item for item in all_candidates if item.lower().startswith(typed)]
        combo['values'] = new_values

    def save_state(self):
        """画面遷移前に状態を保存"""
        new_keywords = [k.strip() for k in self.keyword_entry.get().split(',') if k.strip()]
        unique_keywords = list(set(new_keywords))[:5]
        self.master.keywords = unique_keywords
        
        for key in ['age', 'price', 'start']:
            # 保存されている tk.StringVar の現在の値を App.range_data に保存
            lower_value = self.lower_vars[key].get().strip()
            upper_value = self.upper_vars[key].get().strip()
            
            self.master.range_data[key]['lower'] = lower_value
            self.master.range_data[key]['upper'] = upper_value


# ==============================================================================
# 3. 画面2: タグ表示とTreeview
# ==============================================================================

class Screen2(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.master = master
        
        self.columnconfigure(0, weight=1) 
        self.rowconfigure(6, weight=3)
        self.rowconfigure(8, weight=1)

        ttk.Label(self, text="追加のキーワード検索:").grid(
            row=0, column=0, columnspan=2, padx=10, pady=(10, 0), sticky='w')
        
        self.add_keyword_entry = ttk.Entry(self)
        self.add_keyword_entry.grid(row=1, column=0, padx=10, pady=(10, 0), sticky='ew')
        ttk.Button(self, text="適応", command=self.apply_new_keywords).grid(
            row=1, column=1, padx=10, pady=(10, 0), sticky='e')
        
        #IDのフレームと配置
        self.tag_frame = ttk.Frame(self)
        self.tag_frame.grid(row=2, column=0, columnspan=2, padx=10, pady=5, sticky='w')

        self.draw_tags()
        ttk.Label(self, text="IDからメールを検索:").grid(
        row = 3, column=0, columnspan=2, padx=10, pady=(10, 0), sticky='w')

        self.id_entry = ttk.Entry(self)
        self.id_entry.grid(row = 4,column=0, padx=10, pady=5, sticky='ew')
        ttk.Button(self, text="検索", command=self.search_by_id).grid(
            row=4, column=1, padx=10, pady=5, sticky='e')

        self.setup_treeview()
        self.display_search_results()

        #本文のフレーム設置
        ttk.Label(self, text="選択行の本文:").grid(row=7, column=0, padx=10, pady=(10, 0), sticky='w')
        self.body_text = tk.Text(self, wrap='word', height=10, state='disabled')
        self.body_text.grid(row=8, column=0, columnspan=2, padx=10, pady=(0, 10), sticky='nsew')

        ttk.Button(self, text="戻る (画面1へ)", command=master.show_screen1).grid(
            row=9, column=0, columnspan=2, padx=10, pady=10)

    # === タグ管理 ===
    def draw_tags(self):
        for widget in self.tag_frame.winfo_children():
            widget.destroy()
        for keyword in self.master.keywords:
            self.create_tag(keyword)
    
    def create_tag(self, keyword):
        tag_container = ttk.Frame(self.tag_frame, relief='solid', borderwidth=1)
        tag_container.pack(side='left', padx=(5, 0), pady=2)
        ttk.Label(tag_container, text=keyword, padding=(5, 2)).pack(side='left')
        ttk.Button(tag_container, text='×', width=2, command=lambda k=keyword: self.remove_tag(k)).pack(side='right')

    def remove_tag(self, keyword):
        if keyword in self.master.keywords:
            self.master.keywords.remove(keyword)
            self.draw_tags()
            self.master.df_filtered_skills = filter_skillsheets(self.master.df_all_skills, self.master.keywords, self.master.range_data)
            self.display_search_results()

    def apply_new_keywords(self):
        new_input = [k.strip() for k in self.add_keyword_entry.get().split(',') if k.strip()]
        combined_keywords = self.master.keywords + new_input
        unique_keywords = list(set(combined_keywords))[:5]
        
        self.master.keywords = unique_keywords
        self.draw_tags()
        self.add_keyword_entry.delete(0, 'end') 
        
        self.master.df_filtered_skills = filter_skillsheets(self.master.df_all_skills, self.master.keywords, self.master.range_data)
        self.display_search_results()

    # === Treeviewと検索 ===
    def setup_treeview(self):
        cols_to_display = [col for col in self.master.df_all_skills.columns if col != '本文']
        self.tree = ttk.Treeview(self, columns=cols_to_display, show='headings')
        
        for col in cols_to_display:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, anchor='w')
            
        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        
        self.tree.grid(row=6, column=0, padx=10, pady=10, sticky='nsew')
        vsb.grid(row=6, column=1, sticky='nse', padx=(0, 10), pady=10)
        
        self.tree.bind('<Double-Button-1>', self.treeview_double_click)

    def display_search_results(self):
        """フィルタリングされたデータをTreeviewに表示する"""
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        for row in self.master.df_filtered_skills.itertuples(index=False):
            values = [getattr(row, col) for col in self.tree['columns']]
            try:
                self.tree.insert('', 'end', values=values)
            except Exception as e:
                print(f"🚨 Treeview挿入エラー: 行データ {values} の挿入に失敗しました: {e}")

    def search_by_id(self):
        """ID入力欄の値を使ってTreeviewをフィルタリングし直す"""
        search_id = self.id_entry.get().strip()
        
        if not search_id:
            self.master.df_filtered_skills = filter_skillsheets(self.master.df_all_skills, self.master.keywords, self.master.range_data)
        else:
            self.master.df_filtered_skills = self.master.df_all_skills[
                self.master.df_all_skills['ENTRY_ID'].astype(str).str.contains(search_id, case=False, na=False)
            ]
            
        self.display_search_results()
        
    # === ダブルクリック処理 (本文表示とIDコピー) ===
    def treeview_double_click(self, event):
        item_id = self.tree.identify_row(event.y)
        if not item_id: return

        self.tree.selection_set(item_id)
        
        self.copy_id_to_entry(item_id)
        self.show_email_body(item_id)

    def copy_id_to_entry(self, item_id):
        try:
            id_index = list(self.tree['columns']).index('ENTRY_ID')
        except ValueError:
            return
            
        values = self.tree.item(item_id, 'values')
        if not values or id_index >= len(values): return
        
        id_value = str(values[id_index])
        
        self.master.clipboard_clear()
        self.master.clipboard_append(id_value)
        
        self.id_entry.delete(0, 'end')
        self.id_entry.insert('end', id_value)

    def show_email_body(self, item_id):
        try:
            entry_id_col_index = list(self.tree['columns']).index('ENTRY_ID')
            tree_values = self.tree.item(item_id, 'values')
            entry_id = tree_values[entry_id_col_index]
            
            body_row = self.master.df_all_skills[self.master.df_all_skills['ENTRY_ID'].astype(str) == str(entry_id)]
            if body_row.empty:
                 email_body = f"ID: {entry_id} の本文データが元のリストに見つかりません。"
            else:
                email_body = body_row['本文'].iloc[0]
            
        except (ValueError, IndexError):
            email_body = "本文データを取得できませんでした。データ構造を確認してください。"
        
        self.body_text.config(state='normal') 
        self.body_text.delete(1.0, tk.END) 
        self.body_text.insert(tk.END, email_body)
        self.body_text.config(state='disabled') 


if __name__ == "__main__":
    if not pd.io.common.file_exists("skillsheets.csv"):
        create_sample_data()
        
    app = App()
    app.mainloop()