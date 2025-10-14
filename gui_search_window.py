# gui_search_window.py
# 責務: 抽出結果（CSV/XLSX）を読み込み、Treeviewで表示し、
#       各種フィルタリング（範囲、チェックボックス、キーワード）をリアルタイムで適用する。

import os 
import tkinter as tk
from tkinter import messagebox, Frame,Scrollbar, IntVar ,Checkbutton, ttk 
import pandas as pd
import re
import unicodedata # ソート用の文字正規化に必要

# 📌 修正: 依存する関数を現在のファイル構成からインポート
from config import SCRIPT_DIR
from email_processor import OUTPUT_FILENAME # 出力ファイル名 (XLSX想定)
from evaluator_core import treeview_sort_column, safe_to_int # ソートと数値変換ロジック
from extraction_core import apply_checkbox_filter # データフィルタリングロジック


def toggle_all_checkboxes(vars_dict, select_state, update_func):
    """全てのチェックボックスの状態を切り替え、テーブルを更新する"""
    for var in vars_dict.values():
        var.set(select_state)
    update_func()


def open_search_window(root):
    output_csv_path = os.path.join(SCRIPT_DIR, OUTPUT_FILENAME)
    
    if not os.path.exists(output_csv_path):
        messagebox.showwarning("警告", f"'{OUTPUT_FILENAME}'がまだ作成されていません。\n先に「抽出を実行」を実行してください。")
        return
    try:
        # XLSXファイルを読み込む
        df = pd.read_excel(output_csv_path) 
        
        # フィルタリング用数値カラムをここで一度だけ作成する (safe_to_intに依存)
        if '年齢' in df.columns:
            df['年齢_数値'] = df['年齢'].apply(safe_to_int)
        else:
            df['年齢_数値'] = None
            
        if '単金' in df.columns:
            # 単金は円単位で格納されている前提
            df['単金_数値'] = df['単金'].apply(safe_to_int)
        else:
            df['単金_数値'] = None

    except pd.errors.EmptyDataError:
        messagebox.showwarning("警告", "ファイルが空です。処理が成功したか確認してください。")
        return
    except Exception as e:
        messagebox.showerror("エラー", f"ファイルの読み込みに失敗しました。\nエラー: {e}")
        return

    # 抽出カラム名の調整
    if 'スキルor言語' in df.columns:
        df = df.rename(columns={'スキルor言語': 'スキル_言語'})
        
    if 'スキル_言語' not in df.columns:
        messagebox.showerror("エラー", "CSVファイルに 'スキル_言語' カラムが見つかりません。")
        return
        
    BUSINESS_COLUMN = '業種' # フィルタリング対象のビジネスカラム
    OS_COLUMN = 'OS'         # フィルタリング対象のOSカラム

    has_business_filter = BUSINESS_COLUMN in df.columns
    has_os_filter = OS_COLUMN in df.columns

    def get_unique_items(df, column):
        """カラム内の重複しない項目を抽出し、出現回数順にソートして返す"""
        all_items_counts = {}
        target_col = column
        
        for items_str in df[target_col].astype(str).dropna():
            for item in re.split(r'[,/・、]', items_str):
                item = item.strip()
                if item and item != 'N/A':
                    all_items_counts[item] = all_items_counts.get(item, 0) + 1
        
        sorted_items = sorted(all_items_counts.items(), key=lambda x: x[1], reverse=True)
        return [item[0] for item in sorted_items]
        
    # ユニークアイテムのリスト生成
    sorted_skills = get_unique_items(df, 'スキル_言語')
    sorted_business = get_unique_items(df, BUSINESS_COLUMN) if has_business_filter else []
    sorted_os = get_unique_items(df, OS_COLUMN) if has_os_filter else []

    MAX_CHECKBOXES = 10 
    
    limited_skills = sorted_skills[:MAX_CHECKBOXES] 
    limited_business = sorted_business[:MAX_CHECKBOXES]
    limited_os = sorted_os[:MAX_CHECKBOXES]
        
    # --- 検索ウィンドウのUI構築 ---
    search_window = tk.Toplevel(root)
    search_window.title(f"スキルシート検索・フィルタリング")
    window_width = 1000; window_height = 700
    screen_width = root.winfo_screenwidth(); screen_height = root.winfo_screenheight()
    x = int((screen_width / 2) - (window_width / 2)); y = int((screen_height / 2) - (window_height / 2))
    search_window.geometry(f"{window_width}x{window_height}+{x}+{y}")
    
    search_window.grid_rowconfigure(0, weight=1); search_window.grid_columnconfigure(0, weight=0); search_window.grid_columnconfigure(1, weight=1)
    
    filter_frame = Frame(search_window, width=280, borderwidth=2, relief="groove")
    filter_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10); filter_frame.grid_propagate(False)
    
    canvas = tk.Canvas(filter_frame)
    v_scrollbar = Scrollbar(filter_frame, orient="vertical", command=canvas.yview)
    scrollable_frame = Frame(canvas) 
    scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=v_scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True)
    v_scrollbar.pack(side="right", fill="y")
    
    # --- フィルタリング変数とUI要素の初期化 ---
    filter_vars = {}; biz_filter_vars = {}; os_filter_vars = {} 
    lang_keyword_var = tk.StringVar(); biz_keyword_var = tk.StringVar(); os_keyword_var = tk.StringVar()
    search_salary_min_var = tk.StringVar(); search_salary_max_var = tk.StringVar()
    search_age_min_var = tk.StringVar(); search_age_max_var = tk.StringVar()
    
    # ... (単金/年齢の範囲入力欄の作成ロジックは省略)
    
    # 単金範囲UI
    tk.Label(scrollable_frame, text=" 単金 (万円) 範囲指定", font=("Arial", 10, "bold")).pack(pady=(10, 0), anchor='w', padx=5)
    salary_frame = Frame(scrollable_frame); salary_frame.pack(fill='x', padx=5)
    tk.Label(salary_frame, text="下限:").pack(side=tk.LEFT)
    tk.Entry(salary_frame, textvariable=search_salary_min_var, width=10).pack(side=tk.LEFT, padx=(2, 0)) 
    tk.Label(salary_frame, text="上限:").pack(side=tk.LEFT)  
    tk.Entry(salary_frame, textvariable=search_salary_max_var, width=10).pack(side=tk.LEFT, padx=(2, 0)) 
    
    # 年齢範囲UI
    tk.Label(scrollable_frame, text=" 年齢 (歳) 範囲指定", font=("Arial", 10, "bold")).pack(pady=(10, 0), anchor='w', padx=5)
    age_frame = Frame(scrollable_frame); age_frame.pack(fill='x', padx=5)
    tk.Label(age_frame, text="下限:").pack(side=tk.LEFT)
    tk.Entry(age_frame,  textvariable=search_age_min_var, width=10).pack(side=tk.LEFT, padx=(2, 0)) 
    tk.Label(age_frame, text="上限:").pack(side=tk.LEFT)
    tk.Entry(age_frame, textvariable=search_age_max_var, width=10).pack(side=tk.LEFT, padx=(2, 0)) 
    
    
    def toggle_all_checkboxes_internal(vars_dict, select_state):
        """全てのチェックボックスの状態を切り替え、テーブルを更新する"""
        for var in vars_dict.values():
            var.set(select_state)
        update_table()
        
    def create_checkbox_section(parent_frame, title, item_list, vars_dict, keyword_var, column_name):
        # ヘッダー
        tk.Label(parent_frame, text=f"\n {title}", font=("Arial", 10, "bold")).pack(pady=(5, 0), anchor='w', padx=5)
        # キーワード検索の説明ラベル
        tk.Label(parent_frame, text=f"キーワード検索（カンマ区切り） (AND条件)", fg="gray", font=("Arial", 9)).pack(anchor='w', padx=5)
        search_entry = tk.Entry(parent_frame, textvariable=keyword_var, width=15)
        search_entry.pack(fill='x', padx=5, pady=(2, 0))
        # 「全て解除」ボタン
        tk.Button(parent_frame, text="全て解除", font=("Arial", 8), command=lambda: toggle_all_checkboxes_internal(vars_dict, 0)).pack(anchor='w',padx=5, pady=(2, 5))
        keyword_var.trace_add("write", lambda *args: update_table())
        checkbox_container = Frame(parent_frame); checkbox_container.pack(fill='x', padx=5, pady=(0, 10))
        for item in item_list:
            var = IntVar(value=0); vars_dict[item] = var
            cb = Checkbutton(checkbox_container, text=item, variable=var, command=update_table, anchor='w')
            cb.pack(fill='x', pady=0, padx=0)
    
    # --- Treeviewの作成ロジック (Treeviewの参照を確保) ---
    result_frame = Frame(search_window, borderwidth=2, relief="groove")
    result_frame.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)
    result_frame.grid_rowconfigure(1, weight=1); result_frame.grid_columnconfigure(0, weight=1)
    status_label_result = tk.Label(result_frame, text="初期化中", font=("Arial", 10))
    status_label_result.grid(row=0, column=0, sticky="w", pady=(0, 5))
    tree_frame = Frame(result_frame); tree_frame.grid(row=1, column=0, sticky="nsew")
    tree_scroll_y = Scrollbar(tree_frame, orient="vertical"); tree_scroll_x = Scrollbar(tree_frame, orient="horizontal")

    display_columns = ['名前', '年齢', '単金', 'スキル_言語', 'OS', '業種', '信頼度スコア', 'メールURL'] 
    actual_cols = [col for col in display_columns if col in df.columns or col == 'メールURL']
    
    tree = ttk.Treeview(
        tree_frame, columns=actual_cols, show='headings', 
        yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set
    )
    
    tree_scroll_y.config(command=tree.yview); tree_scroll_x.config(command=tree.xview)
    tree_scroll_y.pack(side="right", fill="y"); tree_scroll_x.pack(side="bottom", fill="x"); tree.pack(fill="both", expand=True)

    # --- update_table 関数の定義 ---
    def update_table():
        # Treeviewの項目をクリア
        for i in tree.get_children(): tree.delete(i)
            
        filtered_df = df.copy() # 元の全データから開始

        # (フィルタリングロジックの本体)
        selected_skills = [skill for skill, var in filter_vars.items() if var.get() == 1]
        selected_business = [biz for biz, var in biz_filter_vars.items() if var.get() == 1]
        selected_os = [os_item for os_item, var in os_filter_vars.items() if var.get() == 1]
        
        lang_keywords = [k.strip() for k in lang_keyword_var.get().split(',') if k.strip()]
        biz_keywords = [k.strip() for k in biz_keyword_var.get().split(',') if k.strip()]
        os_keywords = [k.strip() for k in os_keyword_var.get().split(',') if k.strip()]
        
        min_salary = safe_to_int(search_salary_min_var.get())
        max_salary = safe_to_int(search_salary_max_var.get())
        min_age = safe_to_int(search_age_min_var.get())
        max_age = safe_to_int(search_age_max_var.get())

        # 1. 単金 範囲フィルタリングの実行 (省略)
        if '単金_数値' in filtered_df.columns and (min_salary is not None or max_salary is not None):
            salary_series = filtered_df['単金_数値']
            min_condition = (salary_series.notna()) & (salary_series >= min_salary) if min_salary is not None else True
            max_condition = (salary_series.notna()) & (salary_series <= max_salary) if max_salary is not None else True
            filtered_df = filtered_df[
                (salary_series.notna() & min_condition & max_condition) |
                salary_series.isna()
            ]

        # 2. 年齢 範囲フィルタリングの実行 (省略)
        if '年齢_数値' in filtered_df.columns and (min_age is not None or max_age is not None):
            age_series = filtered_df['年齢_数値']
            min_condition = (age_series.notna()) & (age_series >= min_age) if min_age is not None else True
            max_condition = (age_series.notna()) & (age_series <= max_age) if max_age is not None else True
            filtered_df = filtered_df[
                (age_series.notna() & min_condition & max_condition) | 
                age_series.isna()
            ]

        # 3-5. スキル, 業務, OSフィルタの適用
        filtered_df = apply_checkbox_filter(filtered_df, 'スキル_言語', selected_skills, lang_keywords)
        if has_business_filter: filtered_df = apply_checkbox_filter(filtered_df, BUSINESS_COLUMN, selected_business, biz_keywords)
        if has_os_filter: filtered_df = apply_checkbox_filter(filtered_df, OS_COLUMN, selected_os, os_keywords)
        
        # Treeviewへの挿入
        display_columns_for_insert = [col for col in display_columns if col in filtered_df.columns]
        for _, row in filtered_df.iterrows():
            row_values = []
            for col in display_columns_for_insert:
                val = row[col]
                if col in ['年齢', '単金']:
                    numeric_val = row.get(f'{col}_数値')
                    if pd.notna(numeric_val): 
                        display_val = str(int(numeric_val) // 10000) if col == '単金' else str(int(numeric_val))
                        row_values.append(display_val)
                    else:
                        row_values.append(str(val))
                else:
                    row_values.append(str(val))

            tree.insert('', 'end', values=row_values)
            
        status_label_result.config(text=f"表示件数: {len(filtered_df)} 件 (全 {len(df)} 件)")

    # フィルタリング入力欄に update_table を紐付け (省略)
    search_salary_min_var.trace_add("write", lambda *args: update_table())
    search_salary_max_var.trace_add("write", lambda *args: update_table())
    search_age_min_var.trace_add("write", lambda *args: update_table())
    search_age_max_var.trace_add("write", lambda *args: update_table())
    
    
    # スキル、業務、OSのチェックボックスセクションを作成
    create_checkbox_section(scrollable_frame, "フィルタリング条件（言語）", limited_skills, filter_vars, lang_keyword_var, 'スキル_言語')
    
    if has_business_filter and limited_business:
        create_checkbox_section(scrollable_frame, "フィルタリング条件（業務）", limited_business, biz_filter_vars, biz_keyword_var, BUSINESS_COLUMN)

    if has_os_filter and limited_os:
        create_checkbox_section(scrollable_frame, "フィルタリング条件（OS）", limited_os, os_filter_vars, os_keyword_var, OS_COLUMN)

    # Treeviewのヘッダー設定 (ソートコマンドをバインド)
    for col in actual_cols: 
        tree.heading(col, text=col)
        width = 100
        if col in ['スキル_言語', 'OS', '業種']: width = 180
        elif col == 'メールURL': width = 300
        tree.column(col, width=width, stretch=tk.YES)
        tree.heading(col, command=lambda c=col: treeview_sort_column(tree, c, False))

    # 初回テーブル更新
    update_table()