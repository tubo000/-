# gui_search_window.py (★ 高度な正規表現 修正版 ★)

import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import pandas as pd
from typing import List
import numpy as np
import os
import re
import sqlite3 
from config import DATABASE_NAME 
import traceback 

# --- ▼▼▼ ★★★ 修正箇所 1: configからパターンをインポート ★★★ ▼▼▼
from config import (
    SKILL_LANGUAGE_PATTERNS, POSITION_PATTERNS, OS_PATTERNS
)
# --- ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲ ---


# ==============================================================================
# 0. 共通ユーティリティ（データ処理ロジック）
# ==============================================================================
#キーワードの準備
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
    'anaconda': 'Python', 'django': 'Python', 'flask': 'Python', 
    'numpy': 'Python', 'pandas': 'Python', 'scikit-learn': 'Python',
    'j2ee': 'Java', 'spring': 'Java', 'hibernate': 'Java', 
    'struts': 'Java', 'mybatis': 'Java', 'jvm': 'Java',
    'typescript': 'JavaScript', 'ts': 'JavaScript', 'vue js': 'JavaScript', 
    'react js': 'JavaScript', 'angular': 'JavaScript', 'next js': 'JavaScript', 
    'nuxt js': 'JavaScript', 'jquery': 'JavaScript',
    'dotnet': 'C#', 'aspnet': 'C#', 'wpf': 'C#', 'xamarin': 'C#',
    'vc++': 'C++', 'stl': 'C++', 'boost': 'C++', 
    'ansi c': 'C言語', 'embedded c': 'C言語', 'posix': 'C言語', 'C':'C言語', 'c':'C言語',
    'laravel': 'PHP', 'symfony': 'PHP', 'cakephp': 'PHP',
    'rails': 'Ruby', 'rspec': 'Ruby',
    'golang': 'Go',
    'swiftui': 'Swift', 'uikit': 'Swift',
    'ktor': 'Kotlin',
    'akka': 'Scala',
    'dart': 'Flutter',
    'react native': 'ReactNative',
    'sass': 'HTML/CSS', 'scss': 'HTML/CSS', 'less': 'HTML/CSS', 'tailwind': 'HTML/CSS', 
    'bootstrap': 'HTML/CSS',
    'jest': 'Testing_FE', 'mocha': 'Testing_FE', 'chai': 'Testing_FE', 'cypress': 'Testing_FE', 'selenium': 'Testing_FE',
    'webpack': 'Bundler', 'babel': 'Bundler', 'rollup': 'Bundler', 'vite': 'Bundler',
    'mysql': 'SQL', 'postgre': 'SQL', 'oracle': 'SQL', 'ms sql': 'SQL', 't-sql': 'SQL', 
    'pl/sql': 'SQL', 'transact-sql': 'SQL',
    'mongodb': 'NoSQL', 'redis': 'NoSQL', 'cassandra': 'NoSQL', 'dynamodb': 'NoSQL',
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

# ------------------------------------------------------------------------------
# ▼▼▼【ロジック核心】キーワード検索ロジック (★ v5 ハイブリッド検索 修正版 ★) ▼▼▼
# ------------------------------------------------------------------------------
def _execute_keyword_search(df: pd.DataFrame, or_groups: List[List[str]]) -> pd.DataFrame:
    """
    キーワードに基づいて検索を実行する。
    (★ v5: 高精度/広範囲でロジックを分離 ★)
    """
    if df.empty or not or_groups: 
        return df

    ALL_PATTERNS = {}
    ALL_PATTERNS.update(SKILL_LANGUAGE_PATTERNS)
    ALL_PATTERNS.update(POSITION_PATTERNS)
    ALL_PATTERNS.update(OS_PATTERNS)
    
    # 警告(UserWarning)の原因である「キャプチャグループ」を
    # 「非キャプチャグループ (?:...)」に置換するための正規表現
    CAPTURE_GROUP_REGEX = re.compile(r'\((?![?!<])')

    # --- ヘルパー関数: 1つのORグループ (['SE', 'PL']など) で検索を実行 ---
    def execute_or_group_search(
        current_df: pd.DataFrame, 
        or_keywords: List[str], 
        search_cols: List[str], 
        use_config_regex: bool # 👈 ★ 修正: ロジック切り替えフラグ
    ) -> pd.DataFrame:
        """
        指定されたカラム (search_cols) に対し、
        or_keywords の *いずれか1つでも* 含まれている行を抽出する。
        """
        available_cols = [col for col in search_cols if col in current_df.columns]
        if not available_cols or not or_keywords:
            return current_df if use_config_regex else current_df.iloc[0:0] # 高精度の場合は0件

        # --- 1. キーワードと正規表現の準備 ---
        final_regex_patterns = []
        simple_keywords_normalized = [] # 高精度検索用
        
        for kw in or_keywords:
            kw_stripped = kw.strip()
            if not kw_stripped: continue
            
            lower_kw = kw_stripped.lower()
            mapped_value = KEYWORD_MAPPING.get(lower_kw, lower_kw)
            
            # --- ▼▼▼ ★★★ 修正箇所 (v5) ★★★ ▼▼▼
            if use_config_regex:
                # -------------------------------------------------
                # 2a. (広範囲検索用) config.py の高度なRegexを使う
                # -------------------------------------------------
                found_key = None
                if mapped_value in ALL_PATTERNS: found_key = mapped_value
                elif mapped_value.upper() in ALL_PATTERNS: found_key = mapped_value.upper()
                elif mapped_value.title() in ALL_PATTERNS: found_key = mapped_value.title()
                elif kw_stripped.upper() in ALL_PATTERNS: found_key = kw_stripped.upper()
                elif kw_stripped in ALL_PATTERNS: found_key = kw_stripped

                if found_key:
                    patterns_from_config = ALL_PATTERNS[found_key]
                    # 警告(UserWarning)対策: ( を (?: に置換
                    fixed_patterns = [CAPTURE_GROUP_REGEX.sub(r'(?:', p) for p in patterns_from_config]
                    final_regex_patterns.extend(fixed_patterns)
                else:
                    final_regex_patterns.append(r'\b' + re.escape(mapped_value) + r'\b')
            else:
                # -------------------------------------------------
                # 2b. (高精度検索用) 単純な文字列（C# や Java）を使う
                # -------------------------------------------------
                simple_keywords_normalized.append(mapped_value.lower())
            # --- ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲ ---

        # --- 3. 検索の実行 ---
        
        # (OR条件は False (ヒットなし) で初期化)
        or_filter_condition = pd.Series([False] * len(current_df), index=current_df.index)

        if use_config_regex:
            # --- (広範囲検索) ---
            if not final_regex_patterns:
                return current_df.iloc[0:0] # Regexがないなら0件

            # 改行をスペースに置換し、全列を結合
            df_search_text = current_df[available_cols].astype(str).fillna(' ').agg(
                lambda x: ' '.join(x.str.replace(r'[\n\r\t]+', ' ', regex=True)), 
                axis=1
            ).str.lower()
            
            combined_regex = "|".join(list(set(final_regex_patterns)))
            try:
                or_filter_condition = df_search_text.str.contains(
                    combined_regex, 
                    na=False, 
                    regex=True, 
                    flags=re.IGNORECASE 
                )
            except re.error as e:
                print(f"🚨 正規表現エラー: {e}")
                pass # or_filter_condition は False のまま
        else:
            # --- (高精度検索) ---
            if not simple_keywords_normalized:
                return current_df.iloc[0:0] # キーワードがないなら0件
            
            # ★ 列（スキル, OS, ポジション）を *個別に* チェック
            for col in available_cols:
                col_text_lower = current_df[col].astype(str).str.lower()
                for simple_kw in simple_keywords_normalized:
                    # ★ regex=False (単純な文字列検索)
                    # (例: "java, c#" というセルが "c#" を含むか)
                    or_filter_condition = or_filter_condition | col_text_lower.str.contains(
                        simple_kw, 
                        na=False, 
                        regex=False 
                    )
                    
        return current_df[or_filter_condition]

    # --- 検索対象カラムの定義 ---
    HIGH_PRECISION_COLS = ['ポジション', 'スキル', 'OS']
    BROAD_RANGE_COLS = ['本文(テキスト形式)', '本文(ファイル含む)', '件名']
    
    # ==========================================================
    # 📌 AND/OR 検索実行
    # ==========================================================
    
    df_phase2_result = df.copy()
    
    for or_keywords_list in or_groups:
        if not or_keywords_list: continue 
        # ★ 修正: use_config_regex=False を渡す
        df_phase2_result = execute_or_group_search(df_phase2_result, or_keywords_list, HIGH_PRECISION_COLS, use_config_regex=False)
        if df_phase2_result.empty: break 
    
    print(f"INFO: 高精度検索 (P,S,OS) で {len(df_phase2_result)} 件ヒット。")

    df_for_phase3 = df.copy()
    if not df_phase2_result.empty and 'ENTRY_ID' in df.columns:
        excluded_ids = df_phase2_result['ENTRY_ID'].unique()
        df_for_phase3 = df[~df['ENTRY_ID'].isin(excluded_ids)].copy()

    if df_for_phase3.empty:
          return df_phase2_result.reset_index(drop=True)

    df_phase3_result = df_for_phase3.copy()
    
    for or_keywords_list in or_groups:
        if not or_keywords_list: continue
        # ★ 修正: use_config_regex=True を渡す
        df_phase3_result = execute_or_group_search(df_phase3_result, or_keywords_list, BROAD_RANGE_COLS, use_config_regex=True)
        if df_phase3_result.empty: break

    if not df_phase3_result.empty and 'ポジション' in df_phase3_result.columns:
        df_phase3_result.loc[:, 'ポジション'] = '' 
        
    print(f"INFO: 広範囲検索 (本文,件名) で {len(df_phase3_result)} 件ヒット。")
    
    df_final_filtered = pd.concat([df_phase2_result, df_phase3_result]).reset_index(drop=True)
    
    return df_final_filtered
# ------------------------------------------------------------------------------
# ▲▲▲ _execute_keyword_search 修正ここまで ▲▲▲
# ------------------------------------------------------------------------------


def filter_skillsheets(df: pd.DataFrame, simple_keywords: List[str], or_groups: List[List[str]], range_data: dict) -> pd.DataFrame:
    
    if df.empty: return df 
    
    # 📌 【修正】キーワードフィルタリングを最初に実行
    
    # 1. シンプルキーワード (['C#', 'Java']) を ORグループ ( [['C#'], ['Java']] ) に変換
    simple_groups = [[kw] for kw in simple_keywords if kw.strip()]
    
    # 2. シンプルグループと高度ORグループを結合
    combined_or_groups = simple_groups + or_groups
    
    if combined_or_groups:
        # 結合したグループ (例: [['C#'], ['Java'], ['SE', 'PL']]) で検索実行
        df_filtered = _execute_keyword_search(df.copy(), combined_or_groups)
    else:
        df_filtered = df.copy()
    
    if df_filtered.empty: return df_filtered
    
    # (中略) ... 年齢・単価・実働開始の範囲フィルタリングロジック (L223〜L417) ...
    # (この部分は変更なし)
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
            
            df_filtered_target = df_target[filter_condition].copy() 

            if not df_filtered_target.empty:
                TEMP_SORT_COL = '__temp_sort_date__'
                df_filtered_target[TEMP_SORT_COL] = start_col_target_str.loc[df_filtered_target.index]
                df_filtered_target = df_filtered_target.sort_values(
                    by=TEMP_SORT_COL, 
                    ascending=True, 
                    kind='stable'
                ).drop(columns=[TEMP_SORT_COL]).reset_index(drop=True)
                
            df_filtered = pd.concat([df_sokujitsu, df_filtered_target, df_nan], ignore_index=True)
            df_filtered = df_filtered.reset_index(drop=True)
            
    return df_filtered.reset_index(drop=True)

# ------------------------------------------------------------------------------
# ==============================================================================
# 1. メインアプリケーション（データと画面遷移の管理）
# ==============================================================================

class App(tk.Toplevel):
    def __init__(self, parent, main_elements: dict, data_frame: pd.DataFrame, open_email_callback, db_has_new_data_var: tk.BooleanVar):
        super().__init__(parent) 
        self.master = parent 
        self.main_elements = main_elements 
        self.open_email_callback = open_email_callback
        self.db_has_new_data_var = db_has_new_data_var
        self.title("スキルシート検索アプリ")
        
        # --- ▼▼▼【修正】self.keywords のデータ構造を変更 ▼▼▼ ---
        self.keywords: List[str] = [] # 例: ['C#', 'Java']
        self.or_groups: List[List[str]] = [] # 例: [['SE', 'PL'], ['QA']]
        # --- ▲▲▲ 修正ここまで ▲▲▲ ---
        
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
        """ (変更なし) """
        self.grab_release() 
        run_button = self.main_elements.get("run_button")
        stop_flag = self.main_elements.get("stop_extraction_flag")
        is_running = False 
        if run_button and str(run_button.cget('state')) == tk.DISABLED:
            is_running = True
        if is_running and stop_flag:
            print("INFO: 検索一覧の×ボタン検知。バックグラウンド処理に停止を要求します...")
            stop_flag.set()
            self.main_elements["is_shutting_down"] = True
            try: self.destroy()
            except tk.TclError: pass
        else:
            print("INFO: 処理は実行されていません。アプリ全体を終了します。")
            try: self.master.destroy() 
            except tk.TclError: pass 
            try: self.destroy()
            except tk.TclError: pass
            
    def on_return_to_main(self):
        self.grab_release()
        self.master.deiconify() 
        self.destroy()

    def _clean_data(self, df: pd.DataFrame) -> pd.DataFrame:
        """ (変更なし) """
        if df.empty: return pd.DataFrame()
        try:
            df.columns = df.columns.str.strip()
            rename_map = {
                '単金': '単価', 
                'スキルor言語': 'スキル', 
                '名前': '氏名', 
                '期間_開始':'実働開始',
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
        # (Screen1側でタグとして描画されるため、Entryへの復元は不要になった)

    def _set_screen1_keywords(self, keywords_str):
        # (Screen1のUI変更に伴い、この関数はもう使われない)
        pass

    def show_screen2(self):
        # (変更なし)
        if self.current_frame: 
            if isinstance(self.current_frame, Screen1): 
                self.current_frame.save_state() # Screen1 が self.master.range_data を更新
            self.current_frame.destroy()
        
        if not self.df_all_skills.empty:
            # 📌 修正: filter_skillsheets に両方のキーワードリストを渡す
            self.df_filtered_skills = filter_skillsheets(
                self.df_all_skills, self.keywords, self.or_groups, self.range_data)
        else:
            self.df_filtered_skills = pd.DataFrame()
        
        self.screen2 = Screen2(self)
        self.current_frame = self.screen2
        self.current_frame.grid(row=0, column=0, sticky='nsew')

        if self.db_has_new_data_var and self.db_has_new_data_var.get():
            print("INFO: 新規データを検出。Screen2表示時に自動で一覧を更新します...")
            self.after(100, self.auto_refresh_on_startup)
 
    def auto_refresh_on_startup(self):
        """ (変更なし) """
        try:
            if self.current_frame and isinstance(self.current_frame, Screen2):
                print("DEBUG: auto_refresh_on_startup が refresh_data_from_db を呼び出します。")
                self.current_frame.refresh_data_from_db()
            elif self.screen2:
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

# ------------------------------------------------------------------------------
# ▼▼▼【UI核心】Screen1 を Screen2 と同様のUIに修正 ▼▼▼
# ------------------------------------------------------------------------------
class Screen1(ttk.Frame):
    
    def __init__(self, master):
        super().__init__(master)
        self.master = master
        self.lower_widgets = {} 
        self.upper_widgets = {} 
        self.columnconfigure(0, weight=1)
        # 📌 修正: column 1 は不要
        # self.columnconfigure(1, weight=1) 
        
        # --- 1. キーワード入力 (シンプルAND) ---
        ttk.Label(self, text="追加キーワード (カンマ区切り = AND):").grid(row=0, column=0, columnspan=2, padx=10, pady=(10, 0), sticky='w')
        
        kw_simple_frame = ttk.Frame(self)
        kw_simple_frame.grid(row=1, column=0, columnspan=2, padx=10, pady=(0, 10), sticky='ew')
        kw_simple_frame.columnconfigure(0, weight=1)
        
        self.keyword_entry = ttk.Entry(kw_simple_frame) 
        self.keyword_entry.grid(row=0, column=0, sticky='ew')
        
        self.apply_simple_button = ttk.Button(kw_simple_frame, text="適応 (AND)", command=self.apply_simple_keywords)
        self.apply_simple_button.grid(row=0, column=1, padx=(10, 0), sticky='e')

        # --- 2. 高度なAND/OR検索 (ポップアップ) ---
        advanced_button_frame = ttk.Frame(self)
        advanced_button_frame.grid(row=2, column=0, columnspan=2, padx=10, pady=(5, 10), sticky='ew')
        advanced_button_frame.columnconfigure(0, weight=1)
        
        self.advanced_search_button = ttk.Button(advanced_button_frame, text="高度なAND/OR検索...",
                                                  command=self.open_advanced_search_popup)
        self.advanced_search_button.grid(row=0, column=1, sticky='e') # 右寄せ
        
        # 📌 修正: カウントラベルを advanced_button_frame の中 (左) に移動
        self.keyword_count_var = tk.StringVar(value="AND条件: 0/5")
        self.keyword_count_label = ttk.Label(advanced_button_frame, textvariable=self.keyword_count_var, foreground="gray")
        self.keyword_count_label.grid(row=0, column=0, sticky='w') # Label (left)

        # --- 3. タグ表示エリア ---
        # 📌 修正: row=3 に変更 (元のrow=4から)
        self.tag_frame = ttk.Frame(self)
        self.tag_frame.grid(row=3, column=0, columnspan=2, padx=10, pady=5, sticky='w')

        # --- 4. 範囲指定 (行インデックスをずらす) ---
        
        # 📌 修正: row=4
        ttk.Label(self, text="単価 (万円) 範囲指定").grid(row=4, column=0, columnspan=2, padx=10, pady=(10, 0), sticky='w')
        # 📌 修正: row=5
        price_frame = self.create_range_input('price')
        price_frame.grid(row=5, column=0, columnspan=2, padx=10, pady=5, sticky='ew')

        # 📌 修正: row=6
        ttk.Label(self, text="年齢 (歳) 範囲指定").grid(row=6, column=0, columnspan=2, padx=10, pady=(10, 0), sticky='w')
        # 📌 修正: row=7
        age_frame = self.create_range_input('age')
        age_frame.grid(row=7, column=0, columnspan=2, padx=10, pady=5, sticky='ew')

        # 📌 修正: row=8
        ttk.Label(self, text="実働開始 範囲指定 (YYYYMM)").grid(row=8, column=0, columnspan=2, padx=10, pady=(10, 0), sticky='w')
        # 📌 修正: row=9
        start_frame = self.create_range_input('start')
        start_frame.grid(row=9, column=0, columnspan=2, padx=10, pady=5, sticky='ew')
        
        # --- 5. 伸縮スペース & 下部ボタン (行インデックスをずらす) ---
        self.rowconfigure(10, weight=1) # 📌 修正: row=10
        
        button_frame = ttk.Frame(self)
        button_frame.grid(row=11, column=0, columnspan=2, padx=10, pady=10, sticky='sew') # 📌 修正: row=11
        button_frame.columnconfigure(0, weight=0)
        button_frame.columnconfigure(1, weight=1) # 伸縮
        button_frame.columnconfigure(2, weight=0)
        button_frame.columnconfigure(3, weight=0) 

        ttk.Button(button_frame, text="抽出画面に戻る", command=self.master.on_return_to_main).grid(row=0, column=0, padx=5, sticky='sw')
        ttk.Button(button_frame, text="リセット", command=self.reset_fields).grid(row=0, column=2, padx=5, sticky='se')
        ttk.Button(button_frame, text="検索", command=master.show_screen2).grid(row=0, column=3, padx=5, sticky='se')
        
        self.rowconfigure(12, weight=1) # 📌 修正: row=12
        
        # --- 初期化 ---
        self.draw_tags() # 既存のキーワードをタグとして描画
        self._update_keyword_count_label()

    def open_advanced_search_popup(self):
        """ 高度な検索ポップアップウィンドウを開く """
        popup = AdvancedSearchPopup(self.master)
        self.master.wait_window(popup)
        # ポップアップが閉じたら、タグとカウントを再描画
        self.draw_tags()
        self._update_keyword_count_label()
        
    # ---------------------------------------------------------------------
    # --- ▼▼▼【修正】create_range_input を修正 (レイアウト崩れ対応) ▼▼▼ ---
    # ---------------------------------------------------------------------
    def create_range_input(self, key):
        """ 
        下限・上限の入力ウィジェットを含むフレームを作成し、返す。
        (Screen1 の grid に直接配置しない)
        """
        # 1. メインフレームを作成
        frame = ttk.Frame(self)
        # 📌 修正: 4列構成に変更 (ラベル + エントリ) * 2
        frame.columnconfigure(0, weight=0) # Label
        frame.columnconfigure(1, weight=1) # Entry
        frame.columnconfigure(2, weight=0) # Label
        frame.columnconfigure(3, weight=1) # Entry
        
        is_combobox = (key != 'start')
        
        # 2. 下限ウィジェット
        label_lower = ttk.Label(frame, text="下限:")
        label_lower.grid(row=0, column=0, padx=(0, 5), sticky='w') 
        
        if is_combobox:
            widget_lower = ttk.Combobox(frame, values=self.master.all_cands.get(key, []))
            widget_lower.bind('<KeyRelease>', lambda e, k=key, c=widget_lower: self.update_combobox_list(e, k, c))
        else:
            widget_lower = ttk.Entry(frame)
        
        widget_lower.grid(row=0, column=1, padx=(0, 10), pady=0, sticky='ew') 
        initial_lower_val = self.master.range_data[key]['lower']
        widget_lower.insert(0, initial_lower_val)
        self.lower_widgets[key] = widget_lower 
        
        # 3. 上限ウィジェット
        label_upper = ttk.Label(frame, text="上限:")
        label_upper.grid(row=0, column=2, padx=(10, 5), sticky='w') 
        
        if is_combobox:
            widget_upper = ttk.Combobox(frame, values=self.master.all_cands.get(key, []))
            widget_upper.bind('<KeyRelease>', lambda e, k=key, c=widget_upper: self.update_combobox_list(e, k, c))
        else:
            widget_upper = ttk.Entry(frame)
            
        widget_upper.grid(row=0, column=3, padx=(0, 0), pady=0, sticky='ew')
        initial_upper_val = self.master.range_data[key]['upper']
        widget_upper.insert(0, initial_upper_val)
        self.upper_widgets[key] = widget_upper
        
        return frame # 4. 作成したフレームを返す
    # ---------------------------------------------------------------------
    # --- ▲▲▲ create_range_input 修正ここまで ▲▲▲ ---
    # ---------------------------------------------------------------------
        
    def update_combobox_list(self, event, key, combo):
        # (変更なし)
        typed = combo.get().lower()
        all_candidates = self.master.all_cands.get(key, [])
        new_values = [item for item in all_candidates if item.lower().startswith(typed)]
        combo['values'] = new_values

    def save_state(self):
        # (変更なし - キーワードは適応ボタン/ポップアップで保存, レンジのみここで保存)
        for key in ['age', 'price', 'start']:
            if key in self.lower_widgets and self.lower_widgets[key].winfo_exists():
                 self.master.range_data[key]['lower'] = self.lower_widgets[key].get().strip()
            if key in self.upper_widgets and self.upper_widgets[key].winfo_exists():
                 self.master.range_data[key]['upper'] = self.upper_widgets[key].get().strip()
                 
    def reset_fields(self):
        # (変更なし - シンプル/高度/レンジ をすべてリセット)
        self.keyword_entry.delete(0, tk.END)
        self.master.keywords = []
        self.master.or_groups = []
        self.draw_tags() # タグ表示をクリア
        self._update_keyword_count_label() # ラベル表示もリセット

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
        self.master.range_data = {'age': {'lower': '', 'upper': ''}, 'price': {'lower': '', 'upper': ''}, 'start': {'lower': '', 'upper': ''}}
        print("INFO: 検索条件をリセットしました。") 

    # ---------------------------------------------------------------------
    # --- ▼▼▼【新設】Screen1 に タグ管理機能 を追加 (Screen2から移植) ▼▼▼ ---
    # ---------------------------------------------------------------------
    
    def apply_simple_keywords(self):
        """ (Screen1用) シンプルANDキーワードを '適応' する """
        # (変更なし - ロジックはScreen2のapply_new_keywordsと同じ)
        new_input_text = self.keyword_entry.get().strip()
        if not new_input_text:
            return
            
        new_simple_keywords = [k.strip() for k in new_input_text.split(',') if k.strip()]
        
        if not new_simple_keywords:
            return

        current_simple_keywords = self.master.keywords
        current_or_groups = self.master.or_groups
        combined_keywords = list(set(current_simple_keywords + new_simple_keywords))
        
        max_count = 5
        total_count = len(combined_keywords) + len(current_or_groups)
        
        if total_count > max_count:
             allowed_new_count = max_count - (len(current_simple_keywords) + len(current_or_groups))
             if allowed_new_count <= 0:
                 messagebox.showwarning("キーワードグループ数の制限", f"AND条件は最大 {max_count} 個までです。\n現在の合計: {total_count-len(new_simple_keywords)} 個 (新しいキーワードは追加されませんでした)。")
                 self.keyword_entry.delete(0, 'end')
                 return
             
             combined_keywords = list(set(current_simple_keywords + new_simple_keywords[:allowed_new_count]))
             messagebox.showwarning("キーワードグループ数の制限", f"AND条件は最大 {max_count} 個までです。\n{allowed_new_count} 個のキーワードのみ追加しました。")

        self.master.keywords = combined_keywords
        
        self.draw_tags() # タグ再描画
        self.keyword_entry.delete(0, 'end') 
        self._update_keyword_count_label() 

    def draw_tags(self):
        """ (Screen1用) タグを描画する """
        # (変更なし - シンプル/高度/レンジ(Screen1は非表示) に対応)
        for widget in self.tag_frame.winfo_children(): widget.destroy()
        
        for keyword in self.master.keywords:
            self.create_tag(keyword, tag_type='simple', data=keyword)
        
        for or_group_list in self.master.or_groups:
            if or_group_list:
                tag_text = f"({', '.join(or_group_list)})"
                self.create_tag(tag_text, tag_type='or_group', data=or_group_list)
        
        # (Screen1では範囲タグは表示しない)
        
    def create_tag(self, text, tag_type: str, data: any):
        """ (Screen1用) タグUIを作成する """
        # (変更なし - ロジックはScreen2のcreate_tagと同じ)
        tag_container = ttk.Frame(self.tag_frame, relief='solid', borderwidth=1)
        tag_container.pack(side='left', padx=(5, 0), pady=2)
        
        if tag_type == 'or_group':
            ttk.Label(tag_container, text=text, padding=(5, 2), foreground='blue').pack(side='left')
        else:
            ttk.Label(tag_container, text=text, padding=(5, 2)).pack(side='left')
        
        # 削除ボタン
        ttk.Button(tag_container, text='×', width=2, 
                   command=lambda t=tag_type, d=data: self.remove_tag(t, d)
        ).pack(side='right')

    def remove_tag(self, tag_type: str, data: any):
        """ (Screen1用) タグを削除する """
        # (変更なし - ロジックはScreen2のremove_tagと同じ)
        tag_removed = False
        if tag_type == 'simple' and data in self.master.keywords:
            self.master.keywords.remove(data)
            tag_removed = True
        
        elif tag_type == 'or_group' and data in self.master.or_groups:
            self.master.or_groups.remove(data)
            tag_removed = True
        
        if tag_removed:
            self.draw_tags() # タグの再描画
            self._update_keyword_count_label() 
            
    def _update_keyword_count_label(self):
        """ (Screen1用) キーワード数を更新する """
        # (変更なし - ロジックはScreen2の_update_keyword_count_labelと同じ)
        simple_keywords = getattr(self.master, 'keywords', [])
        or_groups = getattr(self.master, 'or_groups', [])
        
        current_count = len(simple_keywords) + len(or_groups)
        max_count = 5 

        if current_count > max_count: text_color = "red"
        elif current_count == max_count: text_color = "blue"
        else: text_color = "gray"

        message = f"AND条件: {current_count}/{max_count}"

        self.keyword_count_var.set(message)
        style_name = 'KeywordCount.TLabel'
        if text_color != "gray":
            style = ttk.Style()
            style.configure(style_name, foreground=text_color)
            self.keyword_count_label.config(style=style_name)
        else:
            self.keyword_count_label.config(style='TLabel')
            
    # ---------------------------------------------------------------------
    # --- ▲▲▲ Screen1 タグ管理機能 ここまで ▲▲▲ ---
    # ---------------------------------------------------------------------

# ------------------------------------------------------------------------------
# ▲▲▲ Screen1 修正ここまで ▲▲▲
# ------------------------------------------------------------------------------

# ------------------------------------------------------------------------------
# ▼▼▼【新設】高度なAND/OR検索用 ポップアップウィンドウ ▼▼▼
# ------------------------------------------------------------------------------
class AdvancedSearchPopup(tk.Toplevel):
    def __init__(self, master_app):
        """
        master_app は App クラスのインスタンス
        """
        super().__init__(master_app)
        self.master_app = master_app
        self.title("高度なAND/OR検索")
        self.geometry("600x400")
        
        # --- データ ---
        # App が保持する or_groups の *コピー* を編集する
        self.local_or_groups_data = [list(g) for g in self.master_app.or_groups]
        self.or_group_entries: List[ttk.Entry] = []

        # --- UI ---
        main_frame = ttk.Frame(self)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(2, weight=1) # Canvas (スクロール領域) を伸縮
        
        ttk.Label(main_frame, text="▼ 以下の条件を *すべて* 満たす (AND検索)").grid(row=0, column=0, sticky='w')
        ttk.Button(main_frame, text="[ + グループを追加 (AND) ]", command=self.add_keyword_group_ui).grid(row=1, column=0, pady=5, sticky='w')

        # --- スクロール可能な入力欄エリア ---
        canvas = tk.Canvas(main_frame)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)
        self.scrollable_frame.columnconfigure(1, weight=1) # Entryを伸縮

        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.grid(row=2, column=0, sticky='nsew')
        scrollbar.grid(row=2, column=1, sticky='ns')

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        # ---

        # --- 下部ボタン ---
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, columnspan=2, sticky='e', pady=(10, 0))
        
        ttk.Button(button_frame, text="キャンセル", command=self.destroy).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="適用", command=self.apply_changes).pack(side=tk.RIGHT, padx=5)

        # --- 初期化 ---
        self.restore_keyword_groups_ui()
        
        self.grab_set() # モーダルウィンドウにする

    def restore_keyword_groups_ui(self):
        """ self.local_or_groups_data からUIを復元する """
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        self.or_group_entries.clear()

        if not self.local_or_groups_data:
            self.add_keyword_group_ui(keywords_list=[])
        else:
            for or_list in self.local_or_groups_data:
                self.add_keyword_group_ui(keywords_list=or_list)

    def add_keyword_group_ui(self, keywords_list: List[str] = None):
        """ キーワード入力欄の1行 (ORグループ) をGUIに追加する """
        row_frame = ttk.Frame(self.scrollable_frame)
        row_frame.grid(sticky='ew', padx=5, pady=2)
        row_frame.columnconfigure(1, weight=1)

        del_button = ttk.Button(row_frame, text="✕", width=3,
                                command=lambda rf=row_frame: self.remove_keyword_group_ui(rf))
        del_button.grid(row=0, column=0, padx=(0, 5))

        entry = ttk.Entry(row_frame, width=60)
        entry.grid(row=0, column=1, sticky='ew')
        
        if keywords_list:
            entry.insert(0, ", ".join(keywords_list))

        label = ttk.Label(row_frame, text=" の *いずれか* を含む (OR)")
        label.grid(row=0, column=2, padx=(5, 0), sticky='w')

        self.or_group_entries.append(entry)
        row_frame.entry_widget = entry

    def remove_keyword_group_ui(self, row_frame_to_delete: ttk.Frame):
        """ 指定された行フレームと対応するEntryを削除する """
        try:
            if row_frame_to_delete.entry_widget in self.or_group_entries:
                self.or_group_entries.remove(row_frame_to_delete.entry_widget)
            row_frame_to_delete.destroy()
            if not self.or_group_entries:
                self.add_keyword_group_ui(keywords_list=[])
        except Exception as e:
            print(f"キーワードグループ削除エラー (Popup): {e}")

    def apply_changes(self):
        """ 
        現在のUIの状態を App の self.or_groups に保存し、ウィンドウを閉じる 
        """
        new_or_groups = []
        for entry in self.or_group_entries:
            if entry.winfo_exists():
                text = entry.get().strip()
                # カンマ(,)で区切り、ORリストを作成
                or_list = [k.strip() for k in text.split(',') if k.strip()]
                if or_list:
                    new_or_groups.append(or_list)
        
        # --- ▼▼▼【修正】上限チェックを追加 ▼▼▼ ---
        
        # 1. 現在の「シンプルAND」の数を取得
        current_simple_keywords = getattr(self.master_app, 'keywords', [])
        current_simple_count = len(current_simple_keywords)
        
        # 2. これから適用しようとしている「高度OR」の数を取得
        new_or_count = len(new_or_groups)
        
        # 3. 合計を計算
        max_count = 5
        total_count = current_simple_count + new_or_count
        
        if total_count > max_count:
            # 制限オーバー
            messagebox.showwarning(
                "キーワードグループ数の制限", 
                f"AND条件は最大 {max_count} 個までです。\n\n"
                f"現在のシンプルAND条件: {current_simple_count} 個\n"
                f"適用しようとしているORグループ: {new_or_count} 個\n"
                f"合計: {total_count} 個\n\n"
                "適用できませんでした。グループの数を減らしてください。",
                parent=self # 📌 ポップアップを親にする
            )
            return # 適用せずにウィンドウも閉じない
            
        # --- ▲▲▲ 修正ここまで ▲▲▲ ---
        
        # App本体のデータを更新
        self.master_app.or_groups = new_or_groups
        print(f"INFO: 高度な検索条件を適用しました。 {len(new_or_groups)} グループ")
        self.destroy() # ウィンドウを閉じる
# ------------------------------------------------------------------------------
# ▲▲▲ AdvancedSearchPopup 新設ここまで ▲▲▲
# ------------------------------------------------------------------------------


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
        
        # --- 1. シンプルAND検索 (変更なし) ---
        ttk.Label(self, text="追加キーワード (カンマ区切り = AND):").grid(row=0, column=0, columnspan=2, padx=10, pady=(10, 0), sticky='w')
        
        kw_simple_frame = ttk.Frame(self)
        kw_simple_frame.grid(row=1, column=0, columnspan=2, padx=10, pady=(0, 10), sticky='ew')
        kw_simple_frame.columnconfigure(0, weight=1)
        
        self.add_keyword_entry = ttk.Entry(kw_simple_frame) 
        self.add_keyword_entry.grid(row=0, column=0, sticky='ew')
        
        ttk.Button(kw_simple_frame, text="適応 (AND追加)", command=self.apply_new_keywords).grid(row=0, column=1, padx=(10, 0), sticky='e')

        # --- 2. 高度なAND/OR検索 (ポップアップ) ---
        # --- ▼▼▼【新設】Screen2 にも 高度な検索ボタン を追加 ▼▼▼ ---
        advanced_button_frame = ttk.Frame(self)
        advanced_button_frame.grid(row=2, column=0, columnspan=2, padx=10, pady=(5, 10), sticky='ew')
        advanced_button_frame.columnconfigure(0, weight=1)
        
        self.advanced_search_button = ttk.Button(advanced_button_frame, text="高度なAND/OR検索...",
                                                  command=self.open_advanced_search_popup)
        self.advanced_search_button.grid(row=0, column=1, sticky='e') # 右寄せ
        
        self.keyword_count_var = tk.StringVar(value="AND条件: 0/5")
        self.keyword_count_label = ttk.Label(advanced_button_frame, textvariable=self.keyword_count_var, foreground="gray")
        self.keyword_count_label.grid(row=0, column=0, sticky='w')
        # --- ▲▲▲ 新設ここまで ▲▲▲ ---

        # --- 3. タグ表示エリア ---
        self.tag_frame = ttk.Frame(self)
        self.tag_frame.grid(row=3, column=0, columnspan=2, padx=10, pady=5, sticky='w') # row=3
        
        # --- 4. ID検索 ---
        ttk.Label(self, text="IDからメールをOutlookで開く:").grid(row = 4, column=0, columnspan=2, padx=10, pady=(10, 0), sticky='w') # row=4
        id_frame = ttk.Frame(self)
        id_frame.grid(row=5, column=0, columnspan=2, padx=10, pady=5, sticky='ew') # row=5
        id_frame.columnconfigure(0, weight=1)
        
        self.id_entry = ttk.Entry(id_frame)
        self.id_entry.grid(row = 0, column=0, sticky='ew')
        ttk.Button(id_frame, text="Outlookで開く", command=self.open_email_from_entry).grid(row=0, column=1, padx=(10, 0), sticky='e')

        # --- 5. Treeview ---
        self.setup_treeview() # row=6
        
        self.sort_column = '受信日時' 
        self.sort_reverse = False     
        
        self.display_search_results()
        
        # --- 6. ボタンフレーム ---
        button_frame = ttk.Frame(self)
        button_frame.grid(row=7, column=0, columnspan=2, padx=10, pady=(10, 0), sticky='ew') # row=7
        button_frame.columnconfigure(0, weight=0) # 本文表示
        button_frame.columnconfigure(1, weight=0) # 添付ファイル
        button_frame.columnconfigure(2, weight=0) # キーワードヒット
        button_frame.columnconfigure(3, weight=0) # 一覧更新
        button_frame.columnconfigure(4, weight=1) # 伸縮する空きスペース
        button_frame.columnconfigure(5, weight=0) # 戻る

        ttk.Button(button_frame, text="本文表示", 
                   command=lambda: self.update_display_area("本文(テキスト形式)")
        ).grid(row=0, column=0, sticky='w', padx=(0, 10))
        
        self.btn_attachment_content = ttk.Button(
            button_frame, text="添付ファイル内容表示", 
            command=lambda: self.update_display_area("本文(ファイル含む)"), state='disabled'
        )
        self.btn_attachment_content.grid(row=0, column=1, sticky='w')
        
        self.btn_debug_body = ttk.Button(
            button_frame, 
            text="キーワードヒット箇所表示", 
             command=self.update_display_area_with_debug
        )
        self.btn_debug_body.grid(row=0, column=2, sticky='w', padx=(10, 0)) 

        self.btn_refresh = ttk.Button(
            button_frame, 
            text="一覧更新", 
            command=self.refresh_data_from_db,
            state='disabled' 
        )
        self.btn_refresh.grid(row=0, column=3, sticky='w', padx=(10, 0))
        
        if self.master.db_has_new_data_var:
            def update_refresh_button_state(*args):
                try:
                    if self.master.db_has_new_data_var.get():
                        self.btn_refresh.config(state=tk.NORMAL) 
                    else:
                        self.btn_refresh.config(state=tk.DISABLED) 
                except tk.TclError:
                    pass 
            
            self.master.db_has_new_data_var.trace_add("write", update_refresh_button_state)
            update_refresh_button_state() 
        
        ttk.Button(button_frame, text="戻る (検索条件へ)", command=master.show_screen1
        ).grid(row=0, column=5, sticky='e', padx=10)
        
        # --- 7. テキストエリア ---
        self.body_text = tk.Text(self, wrap='word', height=10, state='disabled',font=('Meiryo', 12))
        self.body_text.grid(row=8, column=0, columnspan=2, padx=10, pady=(0, 10), sticky='nsew') # row=8
        
        # --- 初期化 ---
        if hasattr(self, 'tree'):
            self.tree.bind('<<TreeviewSelect>>', self.on_tree_select)
        
        self.draw_tags() # 既存のタグ描画を呼び出し
        self._update_keyword_count_label() # 初期ロード時にラベルを更新
        self.master.after(100, self._update_debug_button_state)

    def on_tree_select(self, event):
        """Treeviewの項目が選択されたときに呼び出される"""
        selected_items = self.tree.selection()
        if selected_items:
            item_id = selected_items[0]
            self.check_attachment_content(item_id)
        else:
            self.btn_attachment_content.config(state='disabled')
            
        self._update_debug_button_state()


    def open_email_from_entry(self):
        # (変更なし)
        entry_id = self.id_entry.get().strip()
        if hasattr(self.master, 'open_email_callback') and callable(self.master.open_email_callback):
            self.master.open_email_callback(entry_id)
        else:
             print("エラー: open_email_callback が設定されていません。")
             messagebox.showerror("内部エラー", "Outlookを開く機能が正しく設定されていません。")

    def check_attachment_content(self, item_id):
        # (変更なし)
        if not item_id:
            self.btn_attachment_content.config(state='disabled')
            return
        
        is_content_available = False
        try:
            tree_columns = list(self.tree['columns'])
            
            if 'Attachments' not in tree_columns:
                 self.btn_attachment_content.config(state='disabled')
                 return 
                 
            attachments_col_index = tree_columns.index('Attachments')
            tree_values = self.tree.item(item_id, 'values')
            
            if len(tree_values) <= attachments_col_index: return
            
            attachments_data = tree_values[attachments_col_index] 
            
            if attachments_data and str(attachments_data).strip() not in ['', 'N/A']:
                is_content_available = True
                
        except (ValueError, IndexError, KeyError) as e: 
             print(f"check_attachment_content でエラー: {e}")
             pass 
             
        if is_content_available:
            self.btn_attachment_content.config(state='normal') 
        else:
            self.btn_attachment_content.config(state='disabled') 

    def _update_debug_button_state(self):
        # (変更なし - ロジックはシンプル+高度に対応済み)
        has_simple_keywords = bool(getattr(self.master, 'keywords', []))
        has_or_groups = bool(getattr(self.master, 'or_groups', []))
        has_keywords = has_simple_keywords or has_or_groups
        
        is_item_selected = bool(self.tree.selection())

        if has_keywords and is_item_selected:
            self.btn_debug_body.config(state='normal')
        else:
            self.btn_debug_body.config(state='disabled')

    def _debug_keyword_extraction(self, entry_id, col_name, text_content):
        # (変更なし - ロジックはシンプル+高度に対応済み)
        simple_keywords = getattr(self.master, 'keywords', [])
        or_groups = getattr(self.master, 'or_groups', [])
        flat_or_keywords = [item for sublist in or_groups for item in sublist]
        keywords = list(set(simple_keywords + flat_or_keywords))
        
        print(f"🔍 DEBUG: _debug_keyword_extraction 実行中")
        print(f"🔍 参照元キーワードリスト (フラット化): {keywords}")
        
        if not keywords or not text_content:
            if not keywords:
                print("🚨 警告: 参照元キーワードリストが空のため、ヒット箇所検索をスキップします。")
            
            return f" [{col_name}] ヒット箇所検索:" \
                   f"\n  - (キーワードリストが空か、本文データがありません)"
        
        output = []
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
             
        return "\n".join(output)
    
    # (重複する _debug_keyword_extraction を削除)

    def update_display_area_with_debug(self):
        # (変更なし - ロジックはシンプル+高度に対応済み)
        TARGET_COLUMNS = ["本文(テキスト形式)", "本文(ファイル含む)"]
        print(f"DEBUG: キーワードヒット箇所表示ボタンがクリックされました (対象: {TARGET_COLUMNS})")
        
        selected_items = self.tree.selection()
        if not selected_items:
            print("DEBUG: Treeviewで何も選択されていません。処理を中断します。")
            return

        item_id = selected_items[0]
        entry_id = ""
        
        self.body_text.config(state='normal') 
        self.body_text.delete(1.0, tk.END) 
        self.body_text.config(state='disabled')
        self.master.update_idletasks() 

        final_output_parts = []
        
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

        db_path = os.path.abspath(DATABASE_NAME) 
        if not os.path.exists(db_path):
             error_msg = f"データベース {DATABASE_NAME} が見つかりません。"
             print(f"🚨 {error_msg}")
             self.body_text.config(state='normal')
             self.body_text.insert(tk.END, error_msg)
             self.body_text.config(state='disabled')
             return

        conn = None
        try:
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()

            for content_type in TARGET_COLUMNS:
                full_text_content = ""
                current_debug_output = ""
                
                try:
                    query = f"SELECT \"{content_type}\" FROM emails WHERE \"EntryID\" = ?"
                    cursor.execute(query, (entry_id,))
                    row = cursor.fetchone()
                    
                    if row and pd.notna(row[0]) and str(row[0]).strip() != '':
                        full_text_content = str(row[0]).replace('_x000D_', '\n') 
                        current_debug_output = self._debug_keyword_extraction(entry_id, content_type, full_text_content)
                    else :
                        current_debug_output = self._debug_keyword_extraction(entry_id, content_type, "")
                    
                except Exception as col_err:
                    current_debug_output = f"🚨 DBエラー: [{content_type}] の取得中にエラーが発生しました。\n詳細: {col_err}"
                
                final_output_parts.append(current_debug_output)
                
        except Exception as e:
            final_output_parts.append(f"🚨 重大なデータベース接続エラー: {e}")
            
        finally:
            if conn: conn.close()
        
        final_text = "\n\n" + ("\n\n").join(final_output_parts)
        
        self.body_text.config(state='normal') 
        self.body_text.delete(1.0, tk.END) 
        self.body_text.insert(tk.END, final_text)
        self.body_text.config(state='disabled')
        print("DEBUG: 両方の本文デバッグ情報の表示が完了しました。")

# gui_search_window.py (Screen2 クラス内)

    def apply_highlights(self, keywords: List[str]):
        """
        Textウィジェットに表示された全文に対して、
        ★「最初の100件まで」★ のキーワードハイライトを非同期で適用する
        (★ v3w 色が付かないバグ修正版 ★)
        """
        
        HIGHLIGHT_LIMIT = 100 
        
        if not keywords:
            return 
            
        try:
            self.body_text.config(state='normal')
            
            # 既存のハイライトをすべて削除
            self.body_text.tag_remove("highlight", 1.0, tk.END)
            
            # --- ▼▼▼ ★★★ 修正箇所 ★★★ ▼▼▼
            # if "highlight" not in self.body_text.tag_names(): 
            # 
            # 👆 この if チェックを「削除」し、毎回必ずタグを(再)定義する
            # (これにより、タグ定義が失われても色が必ず付くようになる)
            self.body_text.tag_configure(
                "highlight", 
                background="yellow", 
                foreground="black"
            )
            # --- ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲ ---

            full_text = self.body_text.get(1.0, tk.END)
            if not full_text:
                self.body_text.config(state='disabled')
                return

            full_text_lower = full_text.lower() 
            
            total_hits = 0 

            for kw in keywords:
                if not kw.strip():
                    continue
                
                if total_hits >= HIGHLIGHT_LIMIT: 
                    break 

                kw_lower = kw.lower()
                start_index = 0
                
                while True:
                    if total_hits >= HIGHLIGHT_LIMIT:
                        print(f"INFO: ハイライトは {HIGHLIGHT_LIMIT} 件で打ち切りました。")
                        break 

                    start_index = full_text_lower.find(kw_lower, start_index)
                    if start_index == -1:
                        break 
                    
                    end_index = start_index + len(kw_lower)
                    
                    start_tk_index = f"1.0 + {start_index} chars"
                    end_tk_index = f"1.0 + {end_index} chars"
                    self.body_text.tag_add("highlight", start_tk_index, end_tk_index)
                    
                    start_index = end_index 
                    total_hits += 1 
            
            print(f"INFO: ハイライト処理完了。合計 {total_hits} 件をタグ付けしました。")

        except Exception as e:
            print(f"ERROR: ハイライト処理中にエラー: {e}")
            traceback.print_exc()
        finally:
            self.body_text.config(state='disabled')

# gui_search_window.py (Screen2 クラス内)

    def update_display_area(self, content_type: str):
        # (DBオンデマンド読み込み)
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

        full_text_content = "" # 👈 ★ ハイライト処理で使うため、tryの外で定義

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
            try:
                conn = sqlite3.connect(db_path)
                cursor = conn.cursor()
                
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
                        
                        # --- ▼▼▼ ★★★ 修正 (v20のバグ修正) ★★★ ▼▼▼
                        # 1. 1回目の置換
                        target_keywords = r'(\s*(最寄り駅|勤務地|スキル|単価|業務内容|必須スキル|歓迎スキル))' 
                        formatted_content = re.sub(
                            r'([^\n]|^)' + target_keywords, 
                            r'\1\n\2', 
                            full_text_content
                        )
                        # 2. 2回目の置換 (full_text_content ではなく formatted_content を使う)
                        target_chars = r'[■【━―─=最氏所単稼]' 
                        formatted_content = re.sub(
                            r'([^\n]|^)(' + target_chars + '+)', 
                            r'\1\n\2', 
                            formatted_content # 👈 ★ 修正
                        )
                        # --- ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲ ---
                        
                        cleaned_content = formatted_content.strip()
                        cleaned_content = re.sub(r'\n{3,}', '\n\n', cleaned_content)
                        lines = cleaned_content.split('\n')
                        final_formatted_content = []
                        for line in lines:
                            final_formatted_content.append(line.lstrip())
                        
                        display_text = '\n'.join(final_formatted_content)
                        
                    else:
                        display_text = f"{content_type} のデータが空です。"
                else:
                    display_text = f"データベースで EntryID '{entry_id}' が見つかりません。"
            except Exception as db_err:
                 print(f"DB読み込みエラー (update_display_area): {db_err}")
                 display_text = f"データベースからのテキスト取得中にエラーが発生しました。\n詳細: {db_err}"
            finally:
                if conn: conn.close()

            # 最終的な表示処理
            self.body_text.config(state='normal') 
            self.body_text.delete(1.0, tk.END) 
            
            # --- ▼▼▼ ★★★ 修正 (v20のバグ修正) ★★★ ▼▼▼
            # (挿入するのは整形後の display_text)
            self.body_text.insert(tk.END, display_text) 
            # --- ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲ ---
            
            self.body_text.config(state='disabled')
        except Exception as e:
            print(f"ERROR: テキストの挿入に失敗: {e}")
            self.body_text.config(state='disabled')
            return
        
        # --- ▼▼▼ ★★★ ハイライト適応コード ★★★ ▼▼▼
        # (v20のコード をv3sのロジック に合わせて修正)
        
        # 1. シンプル(AND)と高度(OR)の両方のキーワードをフラット化（平坦化）
        simple_keywords = getattr(self.master, 'keywords', [])
        or_groups = getattr(self.master, 'or_groups', [])
        flat_or_keywords = [item for sublist in or_groups for item in sublist]
        keywords_to_highlight = list(set(simple_keywords + flat_or_keywords))

        # 2. キーワードが1つでもあれば、50ミリ秒後にハイライト処理を実行
        if keywords_to_highlight: 
            self.after(50, lambda: self.apply_highlights(keywords_to_highlight))
        # --- ▲▲▲ ★★★ 適応コードここまで ★★★ ▲▲▲
    # ------------------------------------------------------------------------------
    # ▼▼▼【UI核心】Screen2 の タグ/キーワード管理機能 (Screen1とほぼ同等) ▼▼▼
    # ------------------------------------------------------------------------------
    
    def open_advanced_search_popup(self):
        """ (Screen2用) 高度な検索ポップアップウィンドウを開く """
        popup = AdvancedSearchPopup(self.master)
        self.master.wait_window(popup)
        # ポップアップが閉じたら、タグとカウントを再描画
        self.draw_tags()
        self._update_keyword_count_label()
        # 📌 Screen2 では、変更を即座にフィルタリングに反映
        self.refresh_data_locally()
        
    def draw_tags(self):
        # (変更なし - ロジックはシンプル+高度に対応済み)
        for widget in self.tag_frame.winfo_children(): widget.destroy()
        
        # 1. シンプルANDキーワード (['A', 'B']) を描画
        for keyword in self.master.keywords:
            self.create_tag(keyword, tag_type='simple', data=keyword)
        
        # 2. 高度ORグループ ( [['SE', 'PL']] ) を描画
        for or_group_list in self.master.or_groups:
            if or_group_list:
                tag_text = f"({', '.join(or_group_list)})"
                self.create_tag(tag_text, tag_type='or_group', data=or_group_list)
        
        # 3. 範囲指定タグ (Screen2では範囲指定タグも表示)
        range_map = {'age': '年齢', 'price': '単価', 'start': '実働開始'}
        for key, label in range_map.items():
            lower = self.master.range_data[key]['lower']
            upper = self.master.range_data[key]['upper']
            if lower or upper: 
                self.create_tag(f"{label}: {lower or '下限なし'}~{upper or '上限なし'}", tag_type='range', data=None) 

    def create_tag(self, text, tag_type: str, data: any):
        # (変更なし - ロジックはシンプル+高度に対応済み)
        tag_container = ttk.Frame(self.tag_frame, relief='solid', borderwidth=1)
        tag_container.pack(side='left', padx=(5, 0), pady=2)
        
        if tag_type == 'or_group':
            ttk.Label(tag_container, text=text, padding=(5, 2), foreground='blue').pack(side='left')
        else:
            ttk.Label(tag_container, text=text, padding=(5, 2)).pack(side='left')
        
        if tag_type == 'simple' or tag_type == 'or_group':
            ttk.Button(tag_container, text='×', width=2, 
                       command=lambda t=tag_type, d=data: self.remove_tag(t, d)
            ).pack(side='right')

    def remove_tag(self, tag_type: str, data: any):
        # (変更なし - ロジックはシンプル+高度に対応済み)
        tag_removed = False
        if tag_type == 'simple' and data in self.master.keywords:
            self.master.keywords.remove(data)
            tag_removed = True
        
        elif tag_type == 'or_group' and data in self.master.or_groups:
            self.master.or_groups.remove(data)
            tag_removed = True
        
        if tag_removed:
            self.draw_tags() # タグの再描画
            self.refresh_data_locally() # フィルタリング再実行
            self._update_keyword_count_label() 
            self._update_debug_button_state()

    def apply_new_keywords(self):
        # (変更なし - ロジックはシンプル+高度に対応済み)
        
        # 1. 入力欄からテキストを取得し、カンマで分割 (それぞれがAND条件)
        new_input_text = self.add_keyword_entry.get().strip()
        if not new_input_text:
            return
            
        new_simple_keywords = [k.strip() for k in new_input_text.split(',') if k.strip()]
        
        if not new_simple_keywords:
            return

        # 2. 既存のキーワードリストと結合
        current_simple_keywords = self.master.keywords
        current_or_groups = self.master.or_groups
        combined_keywords = list(set(current_simple_keywords + new_simple_keywords))
        
        # 3. グループ数の制限チェック
        max_count = 5
        total_count = len(combined_keywords) + len(current_or_groups)
        
        if total_count > max_count:
             allowed_new_count = max_count - (len(current_simple_keywords) + len(current_or_groups))
             if allowed_new_count <= 0:
                 messagebox.showwarning("キーワードグループ数の制限", f"AND条件は最大 {max_count} 個までです。\n現在の合計: {total_count-len(new_simple_keywords)} 個 (新しいキーワードは追加されませんでした)。")
                 self.add_keyword_entry.delete(0, 'end')
                 return
             
             combined_keywords = list(set(current_simple_keywords + new_simple_keywords[:allowed_new_count]))
             messagebox.showwarning("キーワードグループ数の制限", f"AND条件は最大 {max_count} 個までです。\n{allowed_new_count} 個のキーワードのみ追加しました。")
        
        # 4. 新しいシンプルANDキーワードを追加
        self.master.keywords = combined_keywords
        
        self.draw_tags() # タグ再描画
        self.add_keyword_entry.delete(0, 'end') 
        
        # 5. フィルタリング再実行
        self.refresh_data_locally()

        if hasattr(self, 'reset_sort_status'):
            self.reset_sort_status()
            
        # 6. ラベルとボタンの状態を更新
        self._update_keyword_count_label() 
        self._update_debug_button_state()
        
    def _update_keyword_count_label(self):
        # (変更なし - ロジックはシンプル+高度に対応済み)
        simple_keywords = getattr(self.master, 'keywords', [])
        or_groups = getattr(self.master, 'or_groups', [])
        
        current_count = len(simple_keywords) + len(or_groups)
        max_count = 5 

        if current_count > max_count: text_color = "red"
        elif current_count == max_count: text_color = "blue"
        else: text_color = "gray"

        message = f"AND条件: {current_count}/{max_count}"

        self.keyword_count_var.set(message)
        style_name = 'KeywordCount.TLabel'
        if text_color != "gray":
            style = ttk.Style()
            style.configure(style_name, foreground=text_color)
            self.keyword_count_label.config(style=style_name)
        else:
            self.keyword_count_label.config(style='TLabel')
            
    def refresh_data_locally(self):
        """ (Screen2用) DBを叩かず、メモリ上のデータでフィルタを再実行する """
        if not self.master.df_all_skills.empty:
            self.master.df_filtered_skills = filter_skillsheets(
                self.master.df_all_skills, 
                self.master.keywords, 
                self.master.or_groups, 
                self.master.range_data
            )
        else:
            self.master.df_filtered_skills = pd.DataFrame()
            
        self.display_search_results()
        
    # ------------------------------------------------------------------------------
    # ▲▲▲ Screen2 タグ管理機能 ここまで ▲▲▲
    # ------------------------------------------------------------------------------

    def setup_treeview(self):
        # (変更なし)
        style = ttk.Style()
        style.configure("Treeview", 
                    font=("Arial", 12), 
                    rowheight=30)
        style.configure("Treeview.Heading", 
                    font=("Arial", 10))
        
        if not self.master.df_all_skills.empty:
             cols_available = self.master.df_all_skills.columns.tolist()
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
            elif col == 'Attachments': width_val = 0 
            
            self.tree.column(col, width=width_val, anchor='w', stretch=(col != 'Attachments'))
            if col == 'Attachments':
                 self.tree.column(col, stretch=tk.NO)
                 
        self.tree.column('ENTRY_ID', width=0, stretch=tk.NO) 
        self.tree.heading('ENTRY_ID', text='')
            
        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.grid(row=6, column=0, columnspan=2, padx=10, pady=10, sticky='nsew')
        vsb.grid(row=6, column=1, sticky='nse', padx=(0, 10), pady=10)
        self.tree.bind('<Double-Button-1>', self.treeview_double_click)
        self.tree.bind('<<TreeviewSelect>>', lambda event: self.check_attachment_content(self.tree.focus()))

    def sort_treeview(self, col):
        # (変更なし)
        if col == self.sort_column and col not in ['受信日時']: 
            self.sort_reverse = not self.sort_reverse
        elif col == '受信日時':
            self.sort_column = col
            self.sort_reverse = True
        else:
            self.sort_column = col
            self.sort_reverse = True 

        if not hasattr(self.master, 'df_filtered_skills') or self.master.df_filtered_skills.empty:
            return

        df = self.master.df_filtered_skills.copy() 
        sort_key = col
        
        if col == '単価':
            df['sort_key_value'] = df['単価'].astype(str)
            df['sort_key_value'] = df['sort_key_value'].str.split('～').str[0]
            df['sort_key_value'] = pd.to_numeric(df['sort_key_value'], errors='coerce')
            sort_key = 'sort_key_value'
        elif col == '年齢':
            df[col] = pd.to_numeric(df[col], errors='coerce')
            sort_key = col
        elif col == '受信日時':
            df['sort_key_date'] = pd.to_datetime(df[col], errors='coerce')
            sort_key = 'sort_key_date'
            
        self.master.df_filtered_skills = df.sort_values(
            by=sort_key,
            ascending=not self.sort_reverse, 
            na_position='last'
        )
        
        if 'sort_key_value' in self.master.df_filtered_skills.columns:
            self.master.df_filtered_skills = self.master.df_filtered_skills.drop(columns=['sort_key_value'])
        if 'sort_key_date' in self.master.df_filtered_skills.columns: 
            self.master.df_filtered_skills = self.master.df_filtered_skills.drop(columns=['sort_key_date'])

        self.display_search_results()

        for old_col in self.tree['columns']:
            if old_col != 'ENTRY_ID':
                clean_text = old_col.replace(' ▼', '').replace(' ▲', '').replace(' ▽', '')
                
                if old_col == self.sort_column:
                    self.tree.heading(old_col, text=clean_text)
                elif old_col in ['年齢', '単価','受信日時']:
                    self.tree.heading(old_col, text=clean_text + ' ▽')
                else:
                    self.tree.heading(old_col, text=clean_text)

        if self.sort_column:
            marker = ' ▼' if self.sort_reverse else ' ▲'
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
                    val = 'nan'
                values.append(val)
            try:
                self.tree.insert('', 'end', values=values)
            except Exception as e:
                print(f"🚨 Treeview挿入エラー: 行データ {values} の挿入に失敗しました: {e}")

    def reset_sort_status(self):
        # (変更なし)
        self.sort_column = None
        self.sort_reverse = False 

        for col in self.tree['columns']:
            if col != 'ENTRY_ID':
                clean_text = col.replace(' ▲', '').replace(' ▼', '').replace(' ▽', '')
                if col in ['年齢', '単価', '受信日時']:
                    self.tree.heading(col, text=clean_text + ' ▽')
                else:
                    self.tree.heading(col, text=clean_text)
        self.display_search_results()
                
    def search_by_id(self):
        # (変更なし)
        search_id = self.id_entry.get().strip()
        if not self.master.df_all_skills.empty and 'ENTRY_ID' in self.master.df_all_skills.columns:
            if not search_id:
                self.master.df_filtered_skills = filter_skillsheets(
                    self.master.df_all_skills, 
                    self.master.keywords, 
                    self.master.or_groups, 
                    self.master.range_data
                )
            else:
                self.master.df_filtered_skills = self.master.df_all_skills[
                    self.master.df_all_skills['ENTRY_ID'].astype(str).str.contains(search_id, case=False, na=False)
                ]
        else:
             self.master.df_filtered_skills = pd.DataFrame()
        self.display_search_results()

    def treeview_double_click(self, event):
        # (変更なし)
        item_id = self.tree.identify_row(event.y)
        if not item_id: return
        self.tree.selection_set(item_id)
        self.copy_id_to_entry(item_id)
        self.update_display_area('本文(テキスト形式)') 

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

    def refresh_data_from_db(self):
        # (変更なし)
        try:
            previous_item_count = len(self.tree.get_children())
        except:
            previous_item_count = 0 

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
                query = f"SELECT {light_columns_sql} FROM emails ORDER BY \"受信日時\" DESC"
                new_df = pd.read_sql_query(query, conn)
                
            finally:
                if conn: conn.close()

            self.master.df_all_skills = self.master._clean_data(new_df)
            
            # 📌 修正: フィルタリングも両方のリストを渡す
            self.master.df_filtered_skills = filter_skillsheets(
                self.master.df_all_skills, 
                self.master.keywords,
                self.master.or_groups,
                self.master.range_data
            )
            
            self.display_search_results()
            
            try:
                current_item_count = len(self.tree.get_children())
            except:
                current_item_count = 0
            
            if self.master.db_has_new_data_var:
                self.master.db_has_new_data_var.set(False)

            self.body_text.config(state='normal') 
            self.body_text.delete(1.0, tk.END) 
            self.body_text.insert(tk.END, f"一覧を更新しました。\n（表示件数: {previous_item_count} 件 → {current_item_count} 件）")
            self.body_text.config(state='disabled')
            
            print(f"INFO: 検索一覧をDBから更新しました。 (表示件数: {previous_item_count} -> {current_item_count})")

        except Exception as e:
            messagebox.showerror("更新エラー", f"一覧の更新中にエラーが発生しました。\n詳細: {e}")
            traceback.print_exc()
        finally:
            if hasattr(self, 'btn_refresh'):
                try:
                    if self.btn_refresh.winfo_exists():
                        pass
                except tk.TclError:
                    pass

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
         '年齢': [30, None],
         '単価': [60, 70],
         '実働開始': ['202501', ''],
         'Attachments': ['file1.xlsx', '']
    })
    
    def dummy_open_email_callback(entry_id):
        print(f"--- [TEST CALLBACK] Outlookでメールを開きます: {entry_id} ---")
        messagebox.showinfo("テストコールバック", f"Outlookを開く関数が呼ばれました。\nID: {entry_id}")

    dummy_main_elements = {}
    dummy_db_flag = tk.BooleanVar(value=False)
        
    app = App(
        root, 
        main_elements=dummy_main_elements, 
        data_frame=df_dummy, 
        open_email_callback=dummy_open_email_callback,
        db_has_new_data_var=dummy_db_flag 
    ) 
    app.mainloop()

if __name__ == "__main__":
    main()