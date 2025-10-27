# gui_search_window.py
# 責務: 抽出結果（Excelファイル）を読み込み、Treeviewで表示し、
#       各種フィルタリング（範囲、チェックボックス、キーワード）をリアルタイムで適用する。

import tkinter as tk
from tkinter import ttk
import pandas as pd
import os
import main_application
from email_processor import OUTPUT_FILENAME # 👈 config.py ではなく、email_processor からインポート

# ==============================================================================
# 0. 共通ユーティリティ（データ処理ロジック）
# ==============================================================================

# バックアップ用のテストデータ作成関数
def create_sample_data():
    """CSVファイルが見つからない場合に、代わりに使用するテスト用のDataFrameを作成する。"""
    data = {
        'ENTRY_ID': [f'ID{i:03}' for i in range(1, 11)],
        '氏名': [f'テスト太郎{i}' for i in range(1, 11)],
        'スキル': ['JAVA, Python, C言語, DB', 'C#, Azure', 'Python, AWS', 'JAVA, AWS', 'C#, Unity', 
                 'Python, AI', 'DB, SQL', 'JAVA, DB', 'C#, .NET', 'Python, Django'],
        '本文': [f'これはメール本文{i}です。詳細情報や経歴はこの本文に記述されています。非常に長いメール本文を想定しています。' for i in range(1, 11)],
        '年齢': [25, 30, pd.NA, 33, 28, 50, 40, 37, 22, 35], # NaNを含む
        '単価': [50, 65, 70, pd.NA, 60, 80, 75, 50, 60, 70], # NaNを含む
        '実働開始': ['202405', '202501', '202407', '202403', '202506', 
                   '2024年01', pd.NA, '202411', '202402', '202502'], # NaNを含む
    }
    return pd.DataFrame(data)

def filter_skillsheets_by_keywords(df: pd.DataFrame, keywords: list) -> pd.DataFrame:
    """キーワードリストを用いて、指定された列に対してAND検索を実行する。"""
    if df.empty or not keywords: return df
    search_cols = [col for col in df.columns if col  in ['スキル','件名','本文','添付ファイル内容']]
    df_search = df[search_cols].astype(str).fillna(' ').agg(' '.join, axis=1).str.lower()
    
    filter_condition = pd.Series([True] * len(df), index=df.index)
    
    for keyword in keywords:
        lower_keyword = keyword.lower().strip()
        if lower_keyword:
            filter_condition = filter_condition & df_search.str.contains(lower_keyword, na=False)
            
    return df[filter_condition]

def filter_skillsheets(df: pd.DataFrame, keywords: list, range_data: dict) -> pd.DataFrame:
    """キーワード（AND検索）と範囲指定（年齢/単価/実働開始）の両方でデータをフィルタリングするメインロジック。"""
    df_filtered = df.copy()
    
    # 1. キーワードフィルタリング (AND条件)
    df_filtered = filter_skillsheets_by_keywords(df_filtered, keywords)
    if df_filtered.empty: return df_filtered
    
    # 2. 範囲指定フィルタリング
    for key, limits in range_data.items():
        lower = limits['lower']
        upper = limits['upper']
        if not lower and not upper: continue

        col_name = {'age': '年齢', 'price': '単価', 'start': '実働開始'}.get(key)
        
        if col_name in ['年齢', '単価']:
            try:
                col = df_filtered[col_name]
                col_numeric = pd.to_numeric(col, errors='coerce') 
                
                is_not_nan = col_numeric.notna()
                
                lower_val = int(lower) if lower and str(lower).isdigit() else col_numeric.min()
                upper_val = int(upper) if upper and str(upper).isdigit() else col_numeric.max()
                
                valid_range_filter = is_not_nan & (col_numeric >= lower_val) & (col_numeric <= upper_val)
                
                filter_condition = valid_range_filter | (~is_not_nan) 
                df_filtered = df_filtered[filter_condition]
                
            except Exception as e:
                print(f"🚨 データ型エラー: '{col_name}'の入力値またはデータが無効です。{e}")
                continue
                
        elif key == 'start' and '実働開始' in df_filtered.columns:
            is_nan_or_nat = df_filtered['実働開始'].isna()
            
            df_target = df_filtered[~is_nan_or_nat].copy()
            start_col_target_str = df_target['実働開始'].astype(str)
            
            filter_condition = pd.Series([True] * len(df_target), index=df_target.index)
            
            if lower: 
                filter_condition = filter_condition & (start_col_target_str >= lower)
            if upper:
                filter_condition = filter_condition & (start_col_target_str <= upper)
                
            df_filtered = pd.concat([
                df_target[filter_condition],
                df_filtered[is_nan_or_nat] # NaNだった行を無条件で追加
            ]).drop_duplicates(keep='first').sort_index()
            
    return df_filtered


# ==============================================================================
# 1. メインアプリケーション（データと画面遷移の管理）
# ==============================================================================

# 📌 修正1: tk.Tk から tk.Toplevel に変更
class App(tk.Toplevel):
    """メインウィンドウとアプリケーションの状態を管理するクラス"""
    
    # 📌 修正2: __init__ で親 (parent) を受け取る
    def __init__(self, parent, file_path=OUTPUT_FILENAME):
        super().__init__(parent) # 👈 親ウィンドウを Toplevel に渡す
        self.master = parent # 👈 親 (root) への参照を保持
        
        self.title("スキルシート検索アプリ")
        #ウィンドウを中央に配置するロジック
        window_width = 900
        window_height = 700
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        center_x = int(screen_width/2 - window_width/2)
        center_y = int(screen_height/2 - window_height/2)
        self.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
        
        # --- 共有データ ---
        self.keywords = []      
        self.range_data = {'age': {'lower': '', 'upper': ''}, 'price': {'lower': '', 'upper': ''}, 'start': {'lower': '', 'upper': ''}} 
        self.all_cands = {
            'age': [str(i) for i in range(20, 71, 5)], 
            'price': [str(i) for i in range(50, 101, 10)],
            'start': ['202401', '202404', '202407', '202410', '202501', '202504']
        }
        
        # データ読み込み
        self.df_all_skills = self._load_data(file_path)
        self.df_filtered_skills = self.df_all_skills.copy()
        
        self.current_frame = None
        self.screen1 = None
        self.screen2 = None
        
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        
        self.show_screen1()
        
        # 📌 修正: 呼び出す関数名を 'on_closing_app' から 'on_closing' に変更
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # Toplevel をモーダルにする
        self.grab_set()

    # ----------------------------------------------------
    # 📌 修正: 'on_closing_app' の定義を 'on_closing' に変更
    # ----------------------------------------------------
    def on_closing(self):
        """「×」ボタン用：ウィンドウを閉じ、メインアプリケーション全体を終了させる"""
        self.grab_release() 
        
        try:
            self.master.destroy() 
        except tk.TclError:
            pass 
            
        try:
            self.destroy()
        except tk.TclError:
            pass
            
    # 📌 修正: 「戻る」ボタン用の 'on_return_to_main' メソッド
    def on_return_to_main(self):
        """「戻る」ボタン用：このToplevelウィンドウのみを閉じ、親を再表示する"""
        self.grab_release()
        self.master.deiconify() 
        self.destroy()



    def _load_data(self, file_path):
        """データファイルを読み込み、必要な列名をリネーム・クリーンアップする"""
        if not os.path.exists(file_path):
            print(f"警告: ファイル '{file_path}' が見つかりません。テストデータを作成します。")
            return create_sample_data()

        try:
            # 📌 修正: engine='openpyxl' を明示的に指定
            df = pd.read_excel(file_path, engine='openpyxl') 
            print(f"ファイル '{file_path}' をXLSX/XLS形式で読み込みました。")
            
            df.columns = df.columns.str.strip()
            
            rename_map = {
                '単金': '単価', 
                'スキルor言語': 'スキル', 
                '名前': '氏名', 
                '期間_開始':'実働開始',
                '本文(テキスト形式)':'本文',
                '本文(ファイル含む)':'添付ファイル内容',
                'メールURL': 'ENTRY_ID'
            }
            
            # その他のリネームを適用
            df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns}, errors='ignore')
            
            # ENTRY_ID列のクリーンアップ
            if 'ENTRY_ID' in df.columns:
                df['ENTRY_ID'] = df['ENTRY_ID'].astype(str).str.replace('outlook:', '', regex=False).str.strip()
                df = df[df['ENTRY_ID'].astype(str).str.len() > 10].reset_index(drop=True)
                
            return df

        except Exception as e:
            print(f"🚨 エラー: データ読み込みに失敗しました。詳細: {e}。テストデータを作成します。")
            return create_sample_data()

    # 📌 修正3: 重複していた show_screen1 の定義を削除
    def show_screen1(self):
        """検索条件入力画面（Screen1）に遷移する"""
        if self.current_frame: self.current_frame.destroy()
        
        self.screen1 = Screen1(self)
        self.current_frame = self.screen1
        self.current_frame.grid(row=0, column=0, sticky='nsew')
        
        # キーワード設定用の文字列を準備
        current_keywords_str = ", ".join(self.keywords)
        
        # キーワード設定を遅延実行
        self.after(10, lambda: self._set_screen1_keywords(current_keywords_str))

    def _set_screen1_keywords(self, keywords_str):
        """after()で遅延実行される、キーワード設定処理"""
        # 📌 修正7: self.screen1 が None でないことを確認
        # 📌 修正7: self.screen1 が None でないことを確認
        if self.screen1:
            self.screen1.keyword_entry.delete(0, tk.END) 
            self.screen1.keyword_entry.insert(0, keywords_str)

    def show_screen2(self):
        """検索結果表示画面（Screen2）に遷移する。"""
        if self.current_frame: 
            if isinstance(self.current_frame, Screen1): 
                self.current_frame.save_state()
            self.current_frame.destroy()
            
        self.df_filtered_skills = filter_skillsheets(
            self.df_all_skills, self.keywords, self.range_data)
        
        self.screen2 = Screen2(self)
        self.current_frame = self.screen2
        self.current_frame.grid(row=0, column=0, sticky='nsew')

# ==============================================================================
# 2. 画面1: 検索条件の入力
# ==============================================================================

class Screen1(ttk.Frame):
    """キーワード、年齢、単価、実働開始の検索条件を入力する画面"""
    def __init__(self, master):
        super().__init__(master)
        self.master = master
        
        # 📌 修正5: ウィジェット本体を保持する辞書を定義
        self.lower_widgets = {} 
        self.upper_widgets = {} 
        
        self.columnconfigure(0, weight=1)
        self.columnconfigure(1, weight=1)
        
        # --- UI部品の配置（Row 0 - Row 7 まで） ---
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
        
        # 📌 修正1: 
        # self.rowconfigure(8, weight=1) # 以前は row 8 が伸縮していた
        
        # ボタンフレームを row=8 に配置（以前は row=9）
        button_frame = ttk.Frame(self)
        button_frame.grid(row=8, column=0, columnspan=2, padx=10, pady=10, sticky='ew')
        
        # 検索ボタン (右寄せ)
        ttk.Button(button_frame, text="検索", command=master.show_screen2).pack(side=tk.RIGHT, padx=5)
        
        # 📌 修正7: 「抽出画面に戻る」ボタンを追加し、Appの on_return_to_main を呼び出す
        ttk.Button(button_frame, text="抽出画面に戻る", command=self.master.on_return_to_main).pack(side=tk.LEFT, padx=5)
        # 📌 修正2: row 9 を伸縮する空きスペースにする
        self.rowconfigure(9, weight=1)

    def create_range_input(self, label_text, key, row):
        """範囲指定用の入力フィールド（ComboboxまたはEntry）を作成する"""
        is_combobox = (key != 'start')

        # 下限
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

        # 上限
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
        # ... (変更なし) ...
        typed = combo.get().lower()
        all_candidates = self.master.all_cands.get(key, [])
        new_values = [item for item in all_candidates if item.lower().startswith(typed)]
        combo['values'] = new_values

    def save_state(self):
        """画面遷移前に現在の入力状態をAppオブジェクトに保存する"""
        
        new_keywords = [k.strip() for k in self.keyword_entry.get().split(',') if k.strip()]
        self.master.keywords = list(set(new_keywords))[:5]
        
        # 📌 修正7: tk.StringVarではなく、ウィジェット本体から値を直接取得
        for key in ['age', 'price', 'start']:
            self.master.range_data[key]['lower'] = self.lower_widgets[key].get().strip()
            self.master.range_data[key]['upper'] = self.upper_widgets[key].get().strip()
# ==============================================================================
# 3. 画面2: タグ表示とTreeview
# ==============================================================================

class Screen2(ttk.Frame):
    """検索結果をTreeviewで表示し、追加検索や本文表示を行う画面"""
    def __init__(self, master):
        super().__init__(master)
        self.master = master
        self.columnconfigure(0, weight=1) 
        self.rowconfigure(6, weight=3)
        self.rowconfigure(8, weight=1)
        
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
        
        # ボタンのコマンドが open_email_from_entry メソッドになっている
        ttk.Button(self, text="Outlookで開く", command=self.open_email_from_entry).grid(row=4, column=1, padx=10, pady=5, sticky='e')

        self.setup_treeview()
        self.display_search_results()

        # ----------------------------------------------------
        # 📌 修正3: ボタンフレーム (row 7) に「戻る」ボタンを移動
        # ----------------------------------------------------
        button_frame = ttk.Frame(self)
        button_frame.grid(row=7, column=0, columnspan=2, padx=10, pady=(10, 0), sticky='ew')
        
        # 本文表示ボタン
        ttk.Button(button_frame, text="本文表示", 
                   command=lambda: self.update_display_area('本文')).pack(side='left', padx=(0, 10))
        
        # 添付ファイル内容表示ボタン
        self.btn_attachment_content = ttk.Button(
            button_frame, text="添付ファイル内容表示", 
            command=lambda: self.update_display_area('添付ファイル内容'),
            state='disabled'
        )
        self.btn_attachment_content.pack(side='left')
        
        # 「戻る (検索条件へ)」ボタンを右端に配置
        ttk.Button(button_frame, text="戻る (検索条件へ)", command=master.show_screen1).pack(side='right', padx=10)
        # ----------------------------------------------------
        
        # 本文/添付ファイル内容表示エリア (row 8)
        self.body_text = tk.Text(self, wrap='word', height=10, state='disabled')
        self.body_text.grid(row=8, column=0, columnspan=2, padx=10, pady=(0, 10), sticky='nsew')
       
        # 📌 修正4: row 9 の古い「戻る」ボタンを削除
        # ttk.Button(self, text="戻る (検索条件へ)", command=master.show_screen1).grid(row=9, ...)


    def open_email_from_entry(self):
        """ID入力欄の値をENTRY_IDとして取得し、外部のOutlook連携関数を呼び出す。"""
        entry_id = self.id_entry.get().strip()
        main_application.open_outlook_email_by_id(entry_id) # I. ロジックから呼び出し

    def check_attachment_content(self, item_id):
        """選択行の添付ファイル内容を確認し、ボタンを有効/無効化する。"""
        # 選択がない場合は無効化して終了
        if not item_id:
            self.btn_attachment_content.config(state='disabled')
            return

        is_content_available = False
        try:
            # 1. 選択行のEntry IDを取得
            entry_id_col_index = list(self.tree['columns']).index('ENTRY_ID')
            tree_values = self.tree.item(item_id, 'values')
            entry_id = tree_values[entry_id_col_index]
            
            # 2. DataFrameから対応する行を検索
            content_row = self.master.df_all_skills[self.master.df_all_skills['ENTRY_ID'].astype(str) == str(entry_id)]
            
            # 3. 添付ファイル内容のデータを確認
            if not content_row.empty and '添付ファイル内容' in content_row.columns:
                content = content_row['添付ファイル内容'].iloc[0]
                
                content_str = str(content).strip().lower()
                
                if pd.notna(content) and content_str not in ['', 'nan', 'n/a']:
                    is_content_available = True
            
        except (ValueError, IndexError, KeyError): 
            pass # エラー時は無効化のまま

        # 4. ボタンの状態を切り替え
        if is_content_available:
            self.btn_attachment_content.config(state='normal') # 有効化
        else:
            self.btn_attachment_content.config(state='disabled') # 無効化

    def update_display_area(self, content_type):
        """本文または添付ファイル内容を下のテキストエリアに表示する"""
        selected_items = self.tree.selection()
        if not selected_items: return

        item_id = selected_items[0]
        email_body = "データを取得できませんでした。"
        full_text = ""
        
        try:
            id_index = list(self.tree['columns']).index('ENTRY_ID')
            tree_values = self.tree.item(item_id, 'values')
            entry_id = tree_values[id_index]
            
            body_row = self.master.df_all_skills[self.master.df_all_skills['ENTRY_ID'].astype(str) == str(entry_id)]
            if not body_row.empty and content_type in body_row.columns:
                full_data = body_row[content_type].iloc[0]
                
                if pd.notna(full_data) and str(full_data).strip() != '':
                    full_text = str(full_data)
                full_text = full_text.replace('_x000D_', '')
                # 1000文字に制限
                email_body = str(full_text)[:1000]
                if len(full_text) > 1000:
                    email_body += "...\n\n[--- 1000文字以降は省略 ---]"
            else:
                email_body = f"{content_type} のデータが空です。"

            
        except (ValueError, IndexError):
            email_body = "選択された行からIDを取得できませんでした。"

        self.body_text.config(state='normal') 
        self.body_text.delete(1.0, tk.END) 
        self.body_text.insert(tk.END, email_body)
        self.body_text.config(state='disabled')
        
    #タグ管理
    def draw_tags(self):
        for widget in self.tag_frame.winfo_children(): widget.destroy()
        
        # 1. キーワードタグの描画 (削除ボタンあり)
        for keyword in self.master.keywords: self.create_tag(keyword, is_keyword=True)
        
        # 2. 範囲指定タグの描画 (削除ボタンなし)
        range_map = {
            'age': '年齢', 
            'price': '単価', 
            'start': '実働開始'
        }
        
        for key, label in range_map.items():
            lower = self.master.range_data[key]['lower']
            upper = self.master.range_data[key]['upper']
            
            if lower or upper: # 下限または上限のいずれかがあればタグを作成
                tag_text = f"{label}: {lower or '下限なし'}~{upper or '上限なし'}"
                self.create_tag(tag_text, is_keyword=False) 

    
    def create_tag(self, text, is_keyword):
        """タグ（キーワードまたは範囲指定）を作成する"""
        tag_container = ttk.Frame(self.tag_frame, relief='solid', borderwidth=1)
        tag_container.pack(side='left', padx=(5, 0), pady=2)
        ttk.Label(tag_container, text=text, padding=(5, 2)).pack(side='left')
        
        if is_keyword:
            ttk.Button(tag_container, text='×', width=2, command=lambda k=text: self.remove_tag(k)).pack(side='right')

    def remove_tag(self, keyword):
        """キーワードタグを削除し、フィルタリングを再実行する"""
        if keyword in self.master.keywords:
            self.master.keywords.remove(keyword)
            self.draw_tags()
            self.master.df_filtered_skills = filter_skillsheets(self.master.df_all_skills, self.master.keywords, self.master.range_data)
            self.display_search_results()

    def apply_new_keywords(self):
        new_input = [k.strip() for k in self.add_keyword_entry.get().split(',') if k.strip()]
        combined_keywords = self.master.keywords + new_input
        self.master.keywords = list(set(combined_keywords))[:5]
        
        self.draw_tags()
        self.add_keyword_entry.delete(0, 'end') 
        
        self.master.df_filtered_skills = filter_skillsheets(self.master.df_all_skills, self.master.keywords, self.master.range_data)
        self.display_search_results()
        
    #Treeviewと検索
    def setup_treeview(self):
        cols_to_display = ['受信日時','件名' ,'スキル', '年齢', '単価', '実働開始'] 
        all_columns = ['ENTRY_ID'] + cols_to_display 
        self.tree = ttk.Treeview(self, columns=all_columns, show='headings')
        
        for col in cols_to_display:
            self.tree.heading(col, text=col)
            
            if col in ['年齢', '単価']: width_val = 40
            elif col in ['実働開始']: width_val = 50
            elif col in ['スキル','件名']: width_val = 150
            else: width_val = 100
            self.tree.column(col, width=width_val, anchor='w')

        self.tree.column('ENTRY_ID', width=0, stretch=tk.NO) 
        self.tree.heading('ENTRY_ID', text='')
            
        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        
        self.tree.grid(row=6, column=0, padx=10, pady=10, sticky='nsew')
        vsb.grid(row=6, column=1, sticky='nse', padx=(0, 10), pady=10)
        
        self.tree.bind('<Double-Button-1>', self.treeview_double_click)
        self.tree.bind('<<TreeviewSelect>>', lambda event: self.check_attachment_content(self.tree.focus()))
        
    #フィルタリングされたデータをTreeviewに表示する
    def display_search_results(self):
        for item in self.tree.get_children(): self.tree.delete(item)
        for row in self.master.df_filtered_skills.itertuples(index=False):
            
            values = []
            for col in self.tree['columns']:
                val = getattr(row, col, 'N/A')
                
                if col == '年齢' or col == '単価':
                    if pd.notna(val):
                        try:
                            val = int(float(val))
                        except (ValueError, TypeError):
                            val = str(val) 

                if col == '受信日時':
                    if pd.notna(val) and str(val).strip() != '':
                        val_str = str(val).split(' ')[0] # 日付のみ
                        val = val_str
                    else:
                        val = '' # 空欄

                values.append(val)
                
            try:
                self.tree.insert('', 'end', values=values)
            except Exception as e:
                print(f"🚨 Treeview挿入エラー: 行データ {values} の挿入に失敗しました: {e}")
                
    #ID入力欄の値を使ってTreeviewをフィルタリングし直す検索ボタンの設定なのでここを変更する
    def search_by_id(self):
        search_id = self.id_entry.get().strip()
        
        if not search_id:
            # ID検索をクリアした場合、元のキーワード/範囲フィルタを再適用する
            self.master.df_filtered_skills = filter_skillsheets(self.master.df_all_skills, self.master.keywords, self.master.range_data)
        else:
            self.master.df_filtered_skills = self.master.df_all_skills[
                self.master.df_all_skills['ENTRY_ID'].astype(str).str.contains(search_id, case=False, na=False)
            ]
            
        self.display_search_results()
        
    #ダブルクリック処理 (本文表示とIDコピー)    
    def treeview_double_click(self, event):
        item_id = self.tree.identify_row(event.y)
        if not item_id: return

        self.tree.selection_set(item_id)
        
        self.copy_id_to_entry(item_id)
        self.show_email_body(item_id)

    def copy_id_to_entry(self, item_id):
        try:
            id_index = list(self.tree['columns']).index('ENTRY_ID')
            values = self.tree.item(item_id, 'values')
            if not values or id_index >= len(values): return
            
            id_value = str(values[id_index])
            
            self.master.clipboard_clear()
            self.master.clipboard_append(id_value)
            
            self.id_entry.delete(0, 'end')
            self.id_entry.insert('end', id_value)
        except ValueError:
            pass

    def show_email_body(self, item_id):
        email_body = "本文データを取得できませんでした。"
        full_text = ""
        try:
            entry_id_col_index = list(self.tree['columns']).index('ENTRY_ID')
            tree_values = self.tree.item(item_id, 'values')
            entry_id = tree_values[entry_id_col_index]
            
            body_row = self.master.df_all_skills[self.master.df_all_skills['ENTRY_ID'].astype(str) == str(entry_id)]
            if not body_row.empty and '本文' in body_row.columns:
                full_data = body_row['本文'].iloc[0]
                if pd.notna(full_data) and str(full_data).strip() != '':
                    full_text = str(full_data)
                    full_text = full_text.replace('_x000D_', '')
                    # 1000文字に制限
                    email_body = str(full_text)[:1000]
                    if len(full_text) > 1000:
                        email_body += "...\n\n[--- 1000文字以降は省略 ---]"
                else:
                   email_body = "本文のデータが空です。"

            else:
                email_body = f"ID: {entry_id} の本文データが元のリストに見つかりません。"
            
        except (ValueError, IndexError):
            pass
        
        self.body_text.config(state='normal') 
        self.body_text.delete(1.0, tk.END) 
        self.body_text.insert(tk.END, email_body)
        self.body_text.config(state='disabled')


# ==============================================================================
# 4. 実行エントリポイント
# ==============================================================================

def main():
    """アプリケーションのメイン実行関数。この関数が呼び出されるとGUIが起動する。"""
    
    # 📌 修正10: Toplevel として起動する場合、親(root) が必要
    # このファイルが直接実行された場合（テスト用）
    root = tk.Tk()
    root.withdraw() # メインのrootは隠す
    app = App(root, file_path=os.path.abspath(OUTPUT_FILENAME))
    app.mainloop()
if __name__ == "__main__":
    main()