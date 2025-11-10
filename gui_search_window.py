# gui_search_window.py (バグ修正版)

import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import pandas as pd
from typing import List
import numpy as np
import os
import re
import sqlite3 # 📌 DB接続のために追加
from config import DATABASE_NAME # 📌 DB名を取得するために追加
import traceback # ← ★★★ この行を追加 ★★★
# import main_application # 循環インポート防止

# ==============================================================================
# 0. 共通ユーティリティ（データ処理ロジック）
# ==============================================================================
#キーワードの準備
def filter_skillsheets_by_keywords(df: pd.DataFrame, keywords: list) -> pd.DataFrame:
    """
    キーワードに基づいて3段階の優先順位で検索を実行し、結果を結合して返す。

    優先度1: ポジション列での完全一致 (最優先)
    優先度2: ポジション, スキル, OSでの高精度検索
    優先度3: 本文, 件名での広範囲検索
    """
    #キーワード検索の補足
    KEYWORD_MAPPING = {
    # --- 1. ポジション/ロールの略語変換 (前回の内容 + QA) ---
    'プログラマー': 'PG',
    'プログラマ': 'PG',
    'プロジェクトマネージャー': 'PM',
    'プロジェクトマネージャ': 'PM',
    'システムエンジニア': 'SE',
    'エンジニア': 'SE',
    '品質保証': 'QA',       # QA -> Quality Assurance
    'クオリティアシュアランス': 'QA',
    
    # --- 2. 技術の類義語変換 (提供リストから抽出) ---
    # Python
    'anaconda': 'Python', 'django': 'Python', 'flask': 'Python', 
    'numpy': 'Python', 'pandas': 'Python', 'scikit-learn': 'Python',
    
    # Java
    'j2ee': 'Java', 'spring': 'Java', 'hibernate': 'Java', 
    'struts': 'Java', 'mybatis': 'Java', 'jvm': 'Java',

    # JavaScript
    'typescript': 'JavaScript', 'ts': 'JavaScript', 'vue js': 'JavaScript', 
    'react js': 'JavaScript', 'angular': 'JavaScript', 'next js': 'JavaScript', 
    'nuxt js': 'JavaScript', 'jquery': 'JavaScript',

    # C#
    'dotnet': 'C#', 'aspnet': 'C#', 'wpf': 'C#', 'xamarin': 'C#',

    # C++ / C言語
    'vc++': 'C++', 'stl': 'C++', 'boost': 'C++', 
    'ansi c': 'C言語', 'embedded c': 'C言語', 'posix': 'C言語', 'C':'C言語', 'c':'C言語',

    # PHP / Ruby / Go / Mobile
    'laravel': 'PHP', 'symfony': 'PHP', 'cakephp': 'PHP',
    'rails': 'Ruby', 'rspec': 'Ruby',
    'golang': 'Go',
    'swiftui': 'Swift', 'uikit': 'Swift',
    'ktor': 'Kotlin',
    'akka': 'Scala',
    'dart': 'Flutter',
    'react native': 'ReactNative',

    # フロントエンド/DevOps/DB
    'sass': 'HTML/CSS', 'scss': 'HTML/CSS', 'less': 'HTML/CSS', 'tailwind': 'HTML/CSS', 
    'bootstrap': 'HTML/CSS',
    'jest': 'Testing_FE', 'mocha': 'Testing_FE', 'chai': 'Testing_FE', 'cypress': 'Testing_FE', 'selenium': 'Testing_FE',
    'webpack': 'Bundler', 'babel': 'Bundler', 'rollup': 'Bundler', 'vite': 'Bundler',
    'mysql': 'SQL', 'postgre': 'SQL', 'oracle': 'SQL', 'ms sql': 'SQL', 't-sql': 'SQL', 
    'pl/sql': 'SQL', 'transact-sql': 'SQL',
    'mongodb': 'NoSQL', 'redis': 'NoSQL', 'cassandra': 'NoSQL', 'dynamodb': 'NoSQL',

    # クラウド/インフラ/その他
    'amazon web services': 'AWS', 'ec2': 'AWS', 's3': 'AWS', 'lambda': 'AWS',
    'microsoft azure': 'Azure',
    'google cloud platform': 'GCP',
    '機械学習': 'ML/DL', '深層学習': 'ML/DL', 'ディープラーニング': 'ML/DL',
    'hadoop': 'BigData', 'spark': 'BigData', 'kafka': 'BigData', 'etl': 'BigData',
    'tableau': 'Data_Viz', 'power bi': 'Data_Viz', 'domo': 'Data_Viz',
    'コンテナ': 'Docker',
    'k8s': 'Kubernetes', 'kubernates': 'Kubernetes',
    'ansible': 'Ansible', 'chef': 'Ansible', 'puppet': 'Ansible',
    'ci/cd': 'CI/CD', 'jenkins': 'CI/CD', 'gitlab ci': 'CI/CD', 'github actions': 'CI/CD', 
    'circleci': 'CI/CD',
    'prometheus': 'Monitoring', 'grafana': 'Monitoring', 'zabbix': 'Monitoring', 
    'new relic': 'Monitoring',
    'vmware': 'Virtualization', 'hyper-v': 'Virtualization', 'esxi': 'Virtualization',
    'アジャイル': 'Agile/Scrum', 'スクラム': 'Agile/Scrum',
    'ウォーターフォール': 'Waterfall',
    'github': 'Git_VCS', 'gitlab': 'Git_VCS', 'bitbucket': 'Git_VCS',
    'jira': 'Task_Mgt', 'backlog': 'Task_Mgt', 'trello': 'Task_Mgt', 'redmine': 'Task_Mgt',
    }
    if df.empty or not keywords: 
        return df

    # 検索対象列の定義
    PRIORITY_COLS = ['ポジション'] 
    HIGH_PRECISION_COLS = ['ポジション', 'スキル', 'OS']
    BROAD_RANGE_COLS = ['本文(テキスト形式)', '本文(ファイル含む)', '件名']
    
    # 全てのキーワードのリストを小文字で準備
    lower_keywords = [kw.lower().strip() for kw in keywords if kw.strip()]
    if not lower_keywords: 
        return df
    processed_keywords = []
    
    for kw in keywords:
        kw_stripped = kw.strip()
        if not kw_stripped: 
            continue
            
        lower_kw = kw_stripped.lower()
        
        # 💡 マッピングをチェックし、あれば略語に置換する
        mapped_value = KEYWORD_MAPPING.get(lower_kw, None)
        
        if mapped_value:
            # 略語が見つかった場合: 略語のみをリストに追加 (元のキーワードは無視)
            processed_keywords.append(mapped_value.lower())
        else:
            # 略語が見つからなかった場合: 元のキーワードをそのまま追加
            processed_keywords.append(lower_kw)
            
    # 重複を排除し、最終的な検索キーワードリストを作成
    lower_keywords = list(set(processed_keywords))

    if not lower_keywords: 
        return df

    # --- ヘルパー関数: 指定カラムでAND検索を実行 ---
    def execute_search(current_df: pd.DataFrame, search_cols: List[str]) -> pd.DataFrame:
        available_cols = [col for col in search_cols if col in current_df.columns]
        if not available_cols:
            return pd.DataFrame()

        # 検索対象の列の値を結合し、小文字にする
        df_search_text = current_df[available_cols].astype(str).fillna(' ').agg(' '.join, axis=1).str.lower()
        
        # キーワードのAND条件を適用
        filter_condition = pd.Series([True] * len(current_df), index=current_df.index)
        for lower_keyword in lower_keywords:
            filter_condition = filter_condition & df_search_text.str.contains(lower_keyword, na=False)
                
        return current_df[filter_condition]
    
    
    # ==========================================================
    # 📌 フェーズ 1: 最優先検索 (ポジション列のみ)
    # ==========================================================
    df_phase1 = execute_search(df, PRIORITY_COLS)
    
    if not df_phase1.empty:
        # フェーズ1で結果があった場合、他のフェーズは実行せず、その結果のみを返す
        print(f"INFO: 最優先検索 (ポジション) で {len(df_phase1)} 件ヒット。他のフェーズはスキップ。")
        return df_phase1.reset_index(drop=True)

    
    # ==========================================================
    # 📌 フェーズ 2 & 3: フェーズ1が空の場合のみ実行
    # ==========================================================

    # --- フェーズ 2: 高精度検索 ---
    # 対象: ポジション, スキル, OS
    df_phase2 = execute_search(df, HIGH_PRECISION_COLS)
    print(f"INFO: 最優先検索が0件のため、高精度検索 (P,S,OS) で {len(df_phase2)} 件ヒット。")

    # --- フェーズ 3: 広範囲検索 (重複排除) ---
    df_for_phase3 = df.copy()
    if not df_phase2.empty and 'ENTRY_ID' in df.columns:
        # フェーズ2でヒットしたレコードを全体から除外
        excluded_ids = df_phase2['ENTRY_ID'].unique()
        df_for_phase3 = df[~df['ENTRY_ID'].isin(excluded_ids)].copy()

    # 対象: 本文, 件名
    df_phase3 = execute_search(df_for_phase3, BROAD_RANGE_COLS).copy() 

# これでdf_phase3は独立したDataFrameとなり、以下の代入で警告が出なくなります。
    if not df_phase3.empty and 'ポジション' in df_phase3.columns:
        df_phase3['ポジション']
    print(f"INFO: 広範囲検索 (本文,件名) で {len(df_phase3)} 件ヒット。")
    
    # --- 💡 広範囲検索でヒットしたレコードの「ポジション」列をクリア (以前の要件) ---
    if not df_phase3.empty and 'ポジション' in df_phase3.columns:
        df_phase3['ポジション'] = '' 
        
    # --- 結果の結合と表示順の決定 (フェーズ2を上位に、フェーズ3を下位に) ---
    df_final_filtered = pd.concat([df_phase2, df_phase3]).reset_index(drop=True)
    
    return df_final_filtered


def filter_skillsheets(df: pd.DataFrame, keywords: list, range_data: dict) -> pd.DataFrame:
    
    if df.empty: return df 
    
    # 📌 【修正】キーワードフィルタリングを最初に実行するよう修正
    # keywords リストが空でない場合のみ実行
    if keywords:
        df_filtered = filter_skillsheets_by_keywords(df.copy(), keywords)
    else:
        df_filtered = df.copy()
    
    if df_filtered.empty: return df_filtered
    
    for key, limits in range_data.items():
        lower = limits.get('lower')
        upper = limits.get('upper')
        
        if not lower and not upper: continue

        # 【★全範囲共通の正規化ロジック★】
        try:
            val_lower = float(re.sub(r'[^0-9.]', '', str(lower))) if lower else -float('inf')
            val_upper = float(re.sub(r'[^0-9.]', '', str(upper))) if upper else float('inf')

            if val_lower != -float('inf') and val_upper != float('inf') and val_lower > val_upper:
                limits['lower'], limits['upper'] = str(upper), str(lower)
                lower = limits.get('lower')
                upper = limits.get('upper')
                
        except (ValueError, TypeError) as e:
            print(f"🚨 範囲フィルタリングエラー (キー: {key})：入力値 '{lower}' または '{upper}' が無効な数値形式です。フィルタリングをスキップします。{e}")
            continue
        # ----------------------------------------
        
        col_name = {'age': '年齢', 'price': '単価', 'start': '実働開始'}.get(key)
        
        if col_name not in df_filtered.columns: continue

        # --- 年齢 (範囲内を優先、NaNを最後に、並び順確定) ---
        if col_name == '年齢':
            try:
                if lower and not str(lower).isdigit():
                    raise ValueError(f"'{col_name}'の下限値は純粋な数字である必要があります。入力値: {lower}")
                if upper and not str(upper).isdigit():
                    raise ValueError(f"'{col_name}'の上限値は純粋な数字である必要があります。入力値: {upper}")

                search_lower = int(lower) if lower else -float('inf')
                search_upper = int(upper) if upper else float('inf')
                
                col = df_filtered[col_name]
                col_numeric = pd.to_numeric(col, errors='coerce') 
                
                is_nan = col_numeric.isna()
                df_nan = df_filtered[is_nan]
                df_target = df_filtered[~is_nan].copy()

                if df_target.empty:
                    df_filtered = df_nan
                    continue
                
                col_numeric_target = col_numeric[~is_nan]
                
                range_condition = (col_numeric_target >= search_lower) & (col_numeric_target <= search_upper)
                
                df_filtered_target = df_target[range_condition]
                
                df_filtered = pd.concat([df_filtered_target, df_nan], ignore_index=True)
                df_filtered = df_filtered.reset_index(drop=True)
                
            except Exception as e:
                print(f"🚨 フィルタリングエラー: '{col_name}' - {e}")
                continue
                
        # --- 単価 (①完全内包[範囲] → ②完全内包[単独] → ③部分重複 → ④NaN の順で並び替え、並び順確定) ---
        elif col_name == '単価':
            try:
                col = df_filtered[col_name]
                
                Ls = int(lower) if lower and str(lower).isdigit() else -float('inf')
                Us = int(upper) if upper and str(upper).isdigit() else float('inf')

                # ... (単価解析ロジック - 変更なし) ...
                col_str = col.astype(str).str.strip()
                col_str_normalized = col_str.str.replace('〜|～|－|~', '-', regex=True)
                is_pd_nan = col.isna() 
                is_str_nan = col_str.str.lower() == 'nan'
                is_original_nan = is_pd_nan | is_str_nan
                parts = col_str_normalized.str.split('-', expand=True)
                Ld_raw = pd.to_numeric(parts[0].str.replace(r'[^0-9]', '', regex=True), errors='coerce')
                Ud_raw = pd.to_numeric(parts.get(1, pd.Series(np.nan, index=parts.index)).str.replace(r'[^0-9]', '', regex=True), errors='coerce')
                is_single_value = Ud_raw.isna() & Ld_raw.notna()
                Ud_raw = Ud_raw.fillna(Ld_raw[is_single_value])
                is_parse_error_nan = Ld_raw.isna() & Ud_raw.isna()
                is_nan_condition = is_original_nan | is_parse_error_nan
                df_group3 = df_filtered[is_nan_condition]
                df_target = df_filtered[~is_nan_condition].copy() 
                if df_target.empty:
                    df_filtered = df_group3
                    continue
                target_index = df_target.index
                Ld = Ld_raw.loc[target_index]
                Ud = Ud_raw.loc[target_index]
                Ld_filled = Ld.fillna(-float('inf'))
                Ud_filled = Ud.fillna(float('inf'))

                cond_overlap = (Ld_filled <= float(Us)) & (Ud_filled >= float(Ls))
                cond_contained = (Ld_filled >= float(Ls)) & (Ud_filled <= float(Us))
                df_target_overlap = df_target[cond_overlap]
                if df_target_overlap.empty:
                    df_filtered = df_group3
                    continue
                cond_contained_overlap = cond_contained[df_target_overlap.index]

                df_group1_original = df_target_overlap[cond_contained_overlap]
                df_group2 = df_target_overlap[~cond_contained_overlap]

                idx_g1 = df_group1_original.index
                Ld_g1 = Ld_raw.loc[idx_g1]
                Ud_g1 = Ud_raw.loc[idx_g1]
                
                cond_range = (Ld_g1 != Ud_g1)
                df_group1A = df_group1_original[cond_range] 
                cond_single = (Ld_g1 == Ud_g1)
                df_group1B = df_group1_original[cond_single] 
                
                df_filtered = pd.concat([df_group1A, df_group1B, df_group2, df_group3], ignore_index=True)
                df_filtered = df_filtered.reset_index(drop=True)

            except Exception as e:
                print(f"🚨 データ型エラー: '{col_name}'の入力値またはデータが無効です。{e}")
                continue

        # --- 実働開始 (期間内の数字 → 即日 → NaN の順に表示を固定、並び順確定) ---
        elif key == 'start' and '実働開始' in df_filtered.columns:
            
            start_col = df_filtered['実働開始']
            start_col_str = start_col.astype(str).str.strip()
            
            # ... (実働開始解析ロジック - 変更なし) ...
            is_sokujitsu = start_col_str.isin(["即日", "即"])
            is_pd_nan = start_col.isna()
            is_str_nan = start_col_str.str.lower() == 'nan'
            is_non_date_base = is_pd_nan | is_str_nan | (start_col_str == "")
            is_date_candidate = ~is_sokujitsu & ~is_non_date_base
            df_date_candidate = df_filtered[is_date_candidate].copy()
            start_col_candidate = df_date_candidate['実働開始']
            date_series = pd.to_datetime(start_col_candidate, format='%Y%m', errors='coerce')
            is_nat_in_candidate = date_series.isna()
            is_nan_condition = is_non_date_base.copy() 
            is_nan_condition.loc[is_nat_in_candidate.index] = is_nat_in_candidate
            df_nan = df_filtered[is_nan_condition]
            df_sokujitsu = df_filtered[is_sokujitsu & (~is_nan_condition)]
            is_target_condition = is_date_candidate
            df_target = df_filtered[is_target_condition].copy() 
            if df_target.empty and df_sokujitsu.empty:
                df_filtered = df_nan
                continue 

            start_col_prepared = df_target['実働開始'].astype(str)
            start_col_target_str = start_col_prepared.str.replace(r'[^0-9]', '', regex=True)
            filter_condition = pd.Series(True, index=df_target.index)
            
            if lower: 
                lower_norm = re.sub(r'[^0-9]', '', str(lower))
                filter_condition = filter_condition & (start_col_target_str >= lower_norm)
            if upper:
                upper_norm = re.sub(r'[^0-9]', '', str(upper))
                filter_condition = filter_condition & (start_col_target_str <= upper_norm)
            
            df_filtered_target = df_target[filter_condition]

            df_filtered = pd.concat([df_filtered_target, df_sokujitsu, df_nan], ignore_index=True)
            df_filtered = df_filtered.reset_index(drop=True)
            
    # すべてのフィルタリングが終わった後、最終結果の並び順を確定させて返却する
    return df_filtered.reset_index(drop=True)



# ------------------------------------------------------------------------------


# ==============================================================================
# 1. メインアプリケーション（データと画面遷移の管理）
# ==============================================================================

class App(tk.Toplevel):
    # (変更なし)
    # --- ▼▼▼【修正】引数に main_elements を追加 ▼▼▼ ---
    def __init__(self, parent, main_elements: dict, data_frame: pd.DataFrame, open_email_callback, db_has_new_data_var: tk.BooleanVar):
        super().__init__(parent) 
        self.master = parent 
        self.main_elements = main_elements # ★ main_elements を保存
        self.open_email_callback = open_email_callback
        self.db_has_new_data_var = db_has_new_data_var
        # --- ▲▲▲ 修正ここまで ▲▲▲ ---
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

# --- ▼▼▼【修正】「×」ボタンの処理を安全な終了に変更 ▼▼▼ ---
    def on_closing(self):
        """「×」ボタンが押されたときの安全な終了処理"""
        self.grab_release() 
        
        run_button = self.main_elements.get("run_button")
        stop_flag = self.main_elements.get("stop_extraction_flag")
        
        is_running = False # 処理中かどうか
        if run_button and str(run_button.cget('state')) == tk.DISABLED:
            is_running = True
            
        if is_running and stop_flag:
            # もし処理が実行中なら
            print("INFO: 検索一覧の×ボタン検知。バックグラウンド処理に停止を要求します...")
            
            # 1. 中断フラグを立てる
            stop_flag.set()
            
            # 2. シャットダウン中フラグを立てる
            self.main_elements["is_shutting_down"] = True
            
            # 3. 検索一覧ウィンドウだけを閉じる (メインウィンドウは閉じない)
            try: self.destroy()
            except tk.TclError: pass
            
        else:
            # 処理が実行中でなければ、アプリ全体を終了する
            print("INFO: 処理は実行されていません。アプリ全体を終了します。")
            try: self.master.destroy() # メインウィンドウを閉じる
            except tk.TclError: pass 
            try: self.destroy()
            except tk.TclError: pass
    # --- ▲▲▲ 修正ここまで ▲▲▲ ---
            
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

        if self.db_has_new_data_var and self.db_has_new_data_var.get():
            print("INFO: 新規データを検出。Screen2表示時に自動で一覧を更新します...")
            # 画面が描画されるのを少し待ってから自動更新を実行
            self.after(100, self.auto_refresh_on_startup)
        # ★★★ 修正ここまで ★★★
 
 
    # ★★★ このメソッドを App クラス内に「追加」 ★★★
    def auto_refresh_on_startup(self):
        """起動時の自動更新処理 (show_screen2 から呼ばれる)"""
        try:
            # 現在の画面が Screen2 であることを確認
            if self.current_frame and isinstance(self.current_frame, Screen2):
                # Screen2 の refresh_data_from_db メソッドを直接呼び出す
                print("DEBUG: auto_refresh_on_startup が refresh_data_from_db を呼び出します。")
                self.current_frame.refresh_data_from_db()
            elif self.screen2:
                # current_frame が設定されるのが遅れる場合も想定
                print("DEBUG: auto_refresh_on_startup (fallback) が refresh_data_from_db を呼び出します。")
                self.screen2.refresh_data_from_db()
            else:
                print("WARN: 自動更新を試みましたが、Screen2 が見つかりません。")
        except Exception as e:
            print(f"ERROR: 起動時の自動更新に失敗しました: {e}")
            traceback.print_exc()


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

        self.setup_treeview() 
        
        # --- ▼▼▼【修正】初期ソート順を True (降順) から False (昇順) に変更 ▼▼▼ ---
        self.sort_column = '受信日時' # 初期ソート列
        self.sort_reverse = False     # 初期ソート順 (False=昇順, True=降順)
        # --- ▲▲▲ 修正ここまで ▲▲▲ ---
        
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
        self.btn_debug_body = ttk.Button(
        button_frame, 
        text="キーワードヒット箇所表示", 
     # 呼び出すメソッドを update_display_area_with_debug に設定
         command=self.update_display_area_with_debug
        )
# column=2 に配置
        self.btn_debug_body.grid(row=0, column=2, sticky='w', padx=(10, 0)) 

# 既存の「一覧更新」ボタンの配置を column=3 に変更
        self.btn_refresh = ttk.Button(
        button_frame, 
        text="一覧更新", 
        command=self.refresh_data_from_db,
        state='disabled' 
        )
# column=3 に配置
        self.btn_refresh.grid(row=0, column=3, sticky='w', padx=(10, 0))
        
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
        #本文の文字の大きさの変更はFONTの後を変更する
        self.body_text = tk.Text(self, wrap='word', height=10, state='disabled',font=('Meiryo', 12))
        self.body_text.grid(row=8, column=0, columnspan=2, padx=10, pady=(0, 10), sticky='nsew')
        if hasattr(self, 'tree'):
            self.tree.bind('<<TreeviewSelect>>', self.on_tree_select)
        
        # 📌 追記: 初期ロード時にボタンの状態をチェック
        self.master.after(100, self._update_debug_button_state)

    def on_tree_select(self, event):
        """Treeviewの項目が選択されたときに呼び出される"""
        selected_items = self.tree.selection()
        if selected_items:
            item_id = selected_items[0]
            # 本文表示処理 (既存ロジックをここに記述)
            
            # 添付ファイルボタンの状態を更新 (既存ロジック)
            self.check_attachment_content(item_id)
        else:
            # 選択解除された場合、添付ファイルボタンを無効化
            self.btn_attachment_content.config(state='disabled')
            
        # 📌 追記: 選択状態が変更された後、デバッグボタンの状態を更新
        self._update_debug_button_state()


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

    def _update_debug_button_state(self):
        """
        キーワードリストに値があり、かつTreeviewで項目が選択されている場合のみ、
        キーワードヒット箇所表示ボタンを有効にする。
        """
        has_keywords = bool(self.master.keywords)
        is_item_selected = bool(self.tree.selection())

        if has_keywords and is_item_selected:
            # 両方の条件を満たした場合、ボタンを有効化
            self.btn_debug_body.config(state='normal')
            print("DEBUG: キーワードヒット箇所表示ボタンを有効化しました。")
        else:
            # どちらかの条件を満たさない場合、ボタンを無効化
            self.btn_debug_body.config(state='disabled')
            print(f"DEBUG: キーワードヒット箇所表示ボタンを無効化しました (Keywords: {has_keywords}, Selected: {is_item_selected})")

    def _debug_keyword_extraction(self, entry_id, col_name, text_content):
        """
        キーワードのヒット箇所を検索し、デバッグ文字列として整形して返す。
        線をすべて削除し、シンプルにする。
        """
        
        keywords = self.master.keywords
        # printデバッグ出力はそのまま残します
        print(f"🔍 DEBUG: _debug_keyword_extraction 実行中")
        print(f"🔍 参照元キーワードリスト (self.master.keywords): {keywords}")
        
        if not keywords or not text_content:
            if not keywords:
                print("🚨 警告: 参照元キーワードリストが空のため、ヒット箇所検索をスキップします。")
            
            # 📌 修正: エラー時の線と不要な改行を削除
            return f" [{col_name}] ヒット箇所検索:" \
                   f"\n  - (キーワードリストが空か、本文データがありません)"
        
        output = []
        # 📌 修正: ヘッダー部分をシンプルに
        output.append(f" [{col_name}] ヒット箇所検索:")
        
        full_text = str(text_content).replace('_x000D_', '\n')
        full_text_lower = full_text.lower()
        
        processed_keywords = [kw for kw in keywords if kw.strip()]
        
        total_hits = 0
        
        for keyword in processed_keywords:
            lower_keyword = keyword.lower()
            if not lower_keyword: continue
            current_search_pos = 0

            while True:
                start_index = full_text_lower.find(lower_keyword, current_search_pos)
                if start_index == -1: break
                
                total_hits += 1
                end_index = start_index + len(lower_keyword)
                current_search_pos = end_index
                
                start_context = max(0, start_index - 3)
                end_context = min(len(full_text), end_index + 3)
                extracted_text = full_text[start_context:end_context].replace('\n', ' ')
                
                output.append(f"  - '{keyword}' -> '{extracted_text}' ({start_index})")
        
        if total_hits == 0:
             output.append("  - ヒットしませんでした。")
             
        # 📌 修正: 末尾の線を削除
        return "\n".join(output)
    
    #デバッグ表示の機能
    def _debug_keyword_extraction(self, entry_id, col_name, text_content):
        """
        キーワードのヒット箇所を検索し、デバッグ文字列として整形して返す。
        線をすべて削除し、シンプルにする。
        """
        
        keywords = self.master.keywords
        # デバッグ出力はそのまま
        print(f"🔍 DEBUG: _debug_keyword_extraction 実行中 (EntryID: {entry_id})")
        print(f"🔍 参照元キーワードリスト (self.master.keywords): {keywords}")
        
        if not keywords or not text_content:
            if not keywords:
                print("🚨 警告: 参照元キーワードリストが空のため、ヒット箇所検索をスキップします。")
            
            # 📌 【修正箇所】ヒットしなかった、またはデータがない場合の処理を修正
            # 線のないシンプルなヘッダーとメッセージを返す
            return f" [{col_name}] ヒット箇所検索:" \
                   f"\n  - (キーワードリストが空か、本文データがありません)"
        
        output = []
        # 📌 修正: ヒットした場合のヘッダーもシンプルに
        output.append(f" [{col_name}] ヒット箇所検索:")
        
        full_text = str(text_content).replace('_x000D_', '\n')
        full_text_lower = full_text.lower()
        
        processed_keywords = [kw for kw in keywords if kw.strip()]
        
        total_hits = 0
        
        for keyword in processed_keywords:
            lower_keyword = keyword.lower()
            if not lower_keyword: continue
            current_search_pos = 0

            while True:
                start_index = full_text_lower.find(lower_keyword, current_search_pos)
                if start_index == -1: break
                
                total_hits += 1
                end_index = start_index + len(lower_keyword)
                current_search_pos = end_index
                
                start_context = max(0, start_index - 3)
                end_context = min(len(full_text), end_index + 3)
                extracted_text = full_text[start_context:end_context].replace('\n', ' ')
                
                output.append(f"  - '{keyword}' -> '{extracted_text}' ({start_index})")
        
        if total_hits == 0:
             output.append("  - ヒットしませんでした。")
             
        # 📌 修正: 末尾の線は出力しない
        return "\n".join(output)

    def update_display_area_with_debug(self):
        """
        本文(テキスト形式)と本文(ファイル含む)の両方について、
        データベースから取得したデータに対してキーワードヒット箇所デバッグ情報を表示する。
        """
        TARGET_COLUMNS = ["本文(テキスト形式)", "本文(ファイル含む)"]
        print(f"DEBUG: キーワードヒット箇所表示ボタンがクリックされました (対象: {TARGET_COLUMNS})")
        
        selected_items = self.tree.selection()
        if not selected_items:
            print("DEBUG: Treeviewで何も選択されていません。処理を中断します。")
            return

        item_id = selected_items[0]
        entry_id = ""
        
        # Textウィジェットを初期クリア
        self.body_text.config(state='normal') 
        self.body_text.delete(1.0, tk.END) 
        self.body_text.config(state='disabled')
        self.master.update_idletasks() 

        # 最終的な表示内容を保持するリスト
        final_output_parts = []
        
        # 1. Treeviewから ENTRY_ID を取得
        try:
            tree_columns = list(self.tree['columns'])
            id_index = tree_columns.index('ENTRY_ID')
            tree_values = self.tree.item(item_id, 'values')
            entry_id = str(tree_values[id_index])
        except Exception as e:
            error_msg = f"データ取得エラー: ENTRY_IDの取得に失敗しました。詳細: {e}"
            print(f"🚨 {error_msg}")
            self.body_text.config(state='normal')
            self.body_text.insert(tk.END, error_msg)
            self.body_text.config(state='disabled')
            return

        # 2. データベース接続情報
        db_path = os.path.abspath(DATABASE_NAME) 
        if not os.path.exists(db_path):
             error_msg = f"データベース {DATABASE_NAME} が見つかりません。"
             print(f"🚨 {error_msg}")
             self.body_text.config(state='normal')
             self.body_text.insert(tk.END, error_msg)
             self.body_text.config(state='disabled')
             return

        # 3. 各カラムをループしてデータを取得し、デバッグ情報を生成
        conn = None
        try:
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()

            for content_type in TARGET_COLUMNS:
                full_text_content = ""
                current_debug_output = ""
                
                try:
                    # データベースからデータを取得
                    query = f"SELECT \"{content_type}\" FROM emails WHERE \"EntryID\" = ?"
                    cursor.execute(query, (entry_id,))
                    row = cursor.fetchone()
                    
                    if row and pd.notna(row[0]) and str(row[0]).strip() != '':
                        full_text_content = str(row[0]).replace('_x000D_', '\n') 
                        
                        # キーワードヒット箇所を生成
                        current_debug_output = self._debug_keyword_extraction(entry_id, content_type, full_text_content)
                    else:
                        current_debug_output = f"🚨 警告: ENTRY_ID '{entry_id}' の [{content_type}] に本文データが見つかりませんでした。"
                    
                except Exception as col_err:
                    current_debug_output = f"🚨 DBエラー: [{content_type}] の取得中にエラーが発生しました。\n詳細: {col_err}"
                
                # デバッグ情報を最終リストに追加
                final_output_parts.append(current_debug_output)
                
        except Exception as e:
            final_output_parts.append(f"🚨 重大なデータベース接続エラー: {e}")
            
        finally:
            if conn: conn.close()
        
        # 4. Text ウィジェットに結果を連結して書き込む
        final_text = "\n\n" + ("\n\n").join(final_output_parts)
        
        self.body_text.config(state='normal') 
        self.body_text.delete(1.0, tk.END) 
        self.body_text.insert(tk.END, final_text)
        self.body_text.config(state='disabled')
        print("DEBUG: 両方の本文デバッグ情報の表示が完了しました。")

    def apply_highlights(self, keywords: List[str]):
        """
        Textウィジェットに表示された全文に対して、キーワードのハイライトを非同期で適用する
        """
        if not keywords:
            return # キーワードがなければ何もしない
        try:
            self.body_text.config(state='normal')
                        # 既存のハイライトをすべて削除            
            self.body_text.tag_remove("highlight", 1.0, tk.END)
                        # ハイライト用のタグを定義 (まだなければ)
            if "highlight" not in self.body_text.tag_names():
                self.body_text.tag_configure(                    "highlight", 
                    background="yellow", 
                    foreground="black"                )
            # print(f"INFO: ハイライト処理を開始... (キーワード: {keywords})")# ウィジェットから全文を取得 (これは軽い)            
            full_text = self.body_text.get(1.0, tk.END)
            if not full_text:
                self.body_text.config(state='disabled')
                return
            full_text_lower = full_text.lower() # 検索用に小文字化
            for kw in keywords:
                if not kw.strip():
                    continue                                
                kw_lower = kw.lower()                
                start_index = 0# ★ re.finditer ではなく、単純な .find() ループの方が高速な場合があるwhile True:
                    # .find() で次の一致箇所を探す
                start_index = full_text_lower.find(kw_lower, start_index)
                if start_index == -1:
                    break # 見つからなければループ終了
                    
                end_index = start_index + len(kw_lower)
                    
                # tk.Textのインデックス形式 (例: "1.0", "1.15") に変換
                start_tk_index = f"1.0 + {start_index} chars"
                end_tk_index = f"1.0 + {end_index} chars"# ハイライトタグを適用
                self.body_text.tag_add("highlight", start_tk_index, end_tk_index)
                    
                start_index = end_index # 次の検索開始位置を更新except Exception as e:
            
        finally:
            self.body_text.config(state='disabled')
                # print("INFO: ハイライト処理完了。")

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
                        # 以下の2行を削除または変更して、省略せずに全文を display_text に格納する
                        display_text = full_text_content
                        
                        # 元のコードにあった1000文字の切り捨てと省略メッセージを追加する処理を削除
                        # if len(full_text_content) > 1000:
                        #     display_text += "...\n\n[--- 1000文字以降は省略 ---]"
                            
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

            self.body_text.config(state='normal') 
            self.body_text.delete(1.0, tk.END) 
            self.body_text.insert(tk.END, full_text_content)
            self.body_text.config(state='disabled')
        except Exception as e:
            print(f"ERROR: テキストの挿入に失敗: {e}")
            self.body_text.config(state='disabled')
            return# ★ 2. ハイライト処理を「遅延実行」させる ★# (これにより、まず全文が表示され、一瞬遅れてハイライトが適用される)
        keywords_to_highlight = self.master.keywords
        if keywords_to_highlight: # キーワードがある場合のみハイライト処理を予約
            self.after(50, lambda: self.apply_highlights(keywords_to_highlight))
        
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

        if hasattr(self, 'reset_sort_status'):
            self.reset_sort_status()
            
        # 📌 追記: キーワードが更新された後、デバッグボタンの状態を更新
        self._update_debug_button_state()

    def _update_debug_button_state(self):
        """
        キーワードリストに値があり、かつTreeviewで項目が選択されている場合のみ、
        キーワードヒット箇所表示ボタンを有効にする。
        """
        # self.master.keywords が存在しない場合のフォールバックを追加
        has_keywords = bool(getattr(self.master, 'keywords', [])) 
        is_item_selected = bool(self.tree.selection())

        if has_keywords and is_item_selected:
            self.btn_debug_body.config(state='normal')
            print("DEBUG: キーワードヒット箇所表示ボタンを有効化しました。")
        else:
            self.btn_debug_body.config(state='disabled')
            print(f"DEBUG: キーワードヒット箇所表示ボタンを無効化しました (Keywords: {has_keywords}, Selected: {is_item_selected})")
        
    # --- ▼▼▼ 修正: バグ2対応 (setup_treeview) ▼▼▼ ---
    def setup_treeview(self):
        style = ttk.Style()
    
    # 📌 Treeview全体の文字サイズ（データ行）と行の高さを設定
    # 例: データ行のフォントサイズを12、行の高さを25pxに設定
        style.configure("Treeview", 
                    font=("Arial", 12), 
                    rowheight=30)
        # 📌 【追加】ヘッダーの文字サイズ（見出し行）を設定
    # 例: ヘッダーのフォントを太字の'Arial'でサイズ10に設定
        style.configure("Treeview.Heading", 
                    font=("Arial", 10)) # ここを変更
        
        if not self.master.df_all_skills.empty:
             cols_available = self.master.df_all_skills.columns.tolist()
             
             # 📌 修正: 'Attachments' を表示対象ベースリストに追加
             cols_to_display_base = ['受信日時', '件名', 'スキル', 'ポジション', 'OS', '年齢', '単価', '実働開始', 'Attachments']
             
             cols_to_display = [col for col in cols_to_display_base if col in cols_available]
             all_columns = ['ENTRY_ID'] + cols_to_display
        else:
             cols_to_display = []
             all_columns = ['ENTRY_ID']

        self.tree = ttk.Treeview(self, columns=all_columns, show='headings')
        
        for col in cols_to_display:
            self.tree.heading(col, text=col)
            if col in ['年齢', '単価','受信日時']: 
                self.tree.heading(col, text=col + ' ▽')
            if col in ['年齢']:     
                width_val = 50
                self.tree.heading(col, command=lambda c=col: self.sort_treeview(c))
            elif col in ['単価']:
                width_val = 70
                self.tree.heading(col, command=lambda c=col: self.sort_treeview(c)) 
            elif col in ['ポジション']: width_val = 80
            elif col in ['実働開始']: width_val = 100
            elif col in ['スキル','件名', 'OS']: width_val = 120
            elif col == '受信日時': 
                width_val = 100
                self.tree.heading(col, command=lambda c=col: self.sort_treeview(c)) 
            
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

    def sort_treeview(self, col):
        """
        Treeviewを特定の列でソートします。
        """
        # ソート列が変更されたか、同じ列がクリックされたかでソート順を反転
        # 📌 受信日時はソート順を反転しない
        if col == self.sort_column and col not in ['受信日時']: 
            self.sort_reverse = not self.sort_reverse
        # 📌 受信日時は常に降順 (ascending=False) に固定
        elif col == '受信日時':
            self.sort_column = col
            self.sort_reverse = True # 常に降順なのでTrueに固定
        else:
            self.sort_column = col
            self.sort_reverse = True  # 新しい列では昇順から開始

        if not hasattr(self.master, 'df_filtered_skills') or self.master.df_filtered_skills.empty:
            return

        df = self.master.df_filtered_skills.copy() 
        sort_key = col
        
        # --- (単価、年齢のカスタムソート処理は省略) ---
        if col == '単価':
            # ... (既存の単価ソートロジック - 変更なし) ...
            df['sort_key_value'] = df['単価'].astype(str)
            df['sort_key_value'] = df['sort_key_value'].str.split('～').str[0]
            df['sort_key_value'] = pd.to_numeric(df['sort_key_value'], errors='coerce')
            sort_key = 'sort_key_value'
            
        elif col == '年齢':
            # ... (既存の年齢ソートロジック - 変更なし) ...
            df[col] = pd.to_numeric(df[col], errors='coerce')
            sort_key = col
            
        elif col == '受信日時':
            # 📌 受信日時は datetime に変換してソートできるようにする
            df['sort_key_date'] = pd.to_datetime(df[col], errors='coerce')
            sort_key = 'sort_key_date'
            
        # DataFrameをソート
        self.master.df_filtered_skills = df.sort_values(
            by=sort_key,
            # 📌 受信日時以外は self.sort_reverse を使用。受信日時は常に降順 (False)
            ascending=False if col == '受信日時' else (not self.sort_reverse), 
            na_position='last' # NaNを最後に配置
        )
        
        # 一時的なソート列があれば削除
        if 'sort_key_value' in self.master.df_filtered_skills.columns:
            self.master.df_filtered_skills = self.master.df_filtered_skills.drop(columns=['sort_key_value'])
        if 'sort_key_date' in self.master.df_filtered_skills.columns: 
            self.master.df_filtered_skills = self.master.df_filtered_skills.drop(columns=['sort_key_date'])


        # Treeviewを再描画してソート結果を反映
        self.display_search_results()

        # ... (ヘッダーの記号表示ロジックの修正) ...
        for old_col in self.tree['columns']:
            if old_col != 'ENTRY_ID':
                clean_text = old_col.replace(' ▼', '').replace(' ▲', '').replace(' △', '')
                
                if old_col == self.sort_column:
                    self.tree.heading(old_col, text=clean_text)
                elif old_col in ['年齢', '単価','受信日時']:
                    self.tree.heading(old_col, text=clean_text + ' ▽')
                else:
                    self.tree.heading(old_col, text=clean_text)

        # 2. 現在のソート列に正しい記号 ('▲' または '▼') を追加
        if self.sort_column:
            # 📌 受信日時は常に降順(▼)なので、sort_reverseに関係なく'▼'
            if self.sort_column == '受信日時':
                 marker = ' ▼' 
            else:
                 marker = ' ▼' if self.sort_reverse else ' ▲' # 降順なら▼, 昇順なら▲
            
            current_text = self.tree.heading(self.sort_column, option="text")
            self.tree.heading(self.sort_column, text=current_text + marker)
        
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
                if val == '' and col in ['年齢', '単価']:
                # 記号 'nan' はほとんどの数値や文字列より後にソートされます
                    val = 'nan'
                values.append(val)
            try:
                self.tree.insert('', 'end', values=values)
            except Exception as e:
                print(f"🚨 Treeview挿入エラー: 行データ {values} の挿入に失敗しました: {e}")

    def reset_sort_status(self):
        """
        単価と年齢列のソート状態をリセットし、ヘッダー表示を '△' に戻す。
        """
        # 1. 内部ソート状態変数をリセット
        self.sort_column = None  # 現在ソートしている列をクリア
        self.sort_reverse = False # ソート順をクリア

        # 2. Treeviewのヘッダー表示をリセット ('△' に戻す)
        for col in self.tree['columns']:
            if col != 'ENTRY_ID':
                # まず既存の記号を削除
                clean_text = col.replace(' ▲', '').replace(' ▼', '').replace(' △', '')
                
                # 年齢と単価のみ '△' を付ける
                if col in ['年齢', '単価']:
                    self.tree.heading(col, text=clean_text + ' △')
                else:
                    # その他の列は記号なしの元のテキストに戻す
                    self.tree.heading(col, text=clean_text)

        # 3. DataFrameのソートをリセット (任意: 最新のフィルタ結果を再表示)
        # Treeviewの再描画を行うことで、データが元の読み込み順に戻ります
        # (ただし、直前のフィルタリング結果の並び順に依存します)
        self.display_search_results()
                
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
    # --- ▼▼▼【このメソッドを丸ごと置き換え】▼▼▼ ---
    def refresh_data_from_db(self):
        """
        データベースから最新の「軽量」データを再読み込みし、
        現在のフィルタを適用して Treeview を更新する。
        """
        
        # --- ▼▼▼ 1. 更新前の件数を取得 ▼▼▼ ---
        try:
            previous_item_count = len(self.tree.get_children())
        except:
            previous_item_count = 0 # エラー時は0件
        # --- ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲ ---

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
                # --- ▼▼▼【修正】SQLクエリに ORDER BY を追加 ▼▼▼ ---
                query = f"SELECT {light_columns_sql} FROM emails ORDER BY \"受信日時\" DESC"
                new_df = pd.read_sql_query(query, conn)
                # --- ▲▲▲ 修正ここまで ▲▲▲ ---
                
            finally:
                if conn: conn.close()

            # --- ▼▼▼【削除】Pandas側でのソート処理を削除 ▼▼▼ ---
            # if not new_df.empty and '受信日時' in new_df.columns:
            #     try:
            #         new_df['受信日時'] = pd.to_datetime(new_df['受信日時'], errors='coerce')
            #         new_df = new_df.sort_values(by='受信日時', ascending=False, na_position='last').reset_index(drop=True)
            #     except Exception as sort_err:
            #          print(f"警告: [Refresh] DB読み込み後のソート失敗: {sort_err}")
            # --- ▲▲▲ 削除ここまで ▲▲▲ ---
            # 5. App(self.master) のデータを更新
            self.master.df_all_skills = self.master._clean_data(new_df)
            
            # 6. 現在のフィルタ(Appが保持)を再適用
            self.master.df_filtered_skills = filter_skillsheets(
                self.master.df_all_skills, 
                self.master.keywords,
                self.master.range_data
            )
            
            # 7. Treeview を再描画
            self.display_search_results()
            
            # --- ▼▼▼ 2. 更新後の件数を取得 ▼▼▼ ---
            try:
                current_item_count = len(self.tree.get_children())
            except:
                current_item_count = 0
            # --- ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲ ---
            
            # 8. 「旗」を倒す
            if self.master.db_has_new_data_var:
                self.master.db_has_new_data_var.set(False)

            # --- ▼▼▼ 3. 表示メッセージを修正 ▼▼▼ ---
            self.body_text.config(state='normal') 
            self.body_text.delete(1.0, tk.END) 
            self.body_text.insert(tk.END, f"一覧を更新しました。\n（表示件数: {previous_item_count} 件 → {current_item_count} 件）")
            self.body_text.config(state='disabled')
            # --- ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲ ---
            
            print(f"INFO: 検索一覧をDBから更新しました。 (表示件数: {previous_item_count} -> {current_item_count})")

        except Exception as e:
            messagebox.showerror("更新エラー", f"一覧の更新中にエラーが発生しました。\n詳細: {e}")
            traceback.print_exc()
        finally:
            # 9. ボタンの状態は「旗」によって自動的に更新される
            if hasattr(self, 'btn_refresh'):
                try:
                    if self.btn_refresh.winfo_exists():
                        pass
                except tk.TclError:
                    pass
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