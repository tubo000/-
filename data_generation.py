# data_generation.py
# 責務: 難易度の高いランダムな試験用ダミーデータと、その正解データを作成・TSV形式で出力する。

import pandas as pd
import random
import os
import csv
import re
# configファイルから、データ生成に必要な定数（辞書、パス、件数）をインポート
from config import NUM_RECORDS, NAMES, LANGUAGES, INDUSTRIES, SALARY_UNITS, NOISE, INPUT_QUESTION_CSV, MASTER_ANSWERS_PATH


def add_noise_to_name(name_tuple):
    """
    設定された名前リストから選ばれた名前に、ランダムなノイズ（スペース、記号、括弧書きの備考）を追加する。
    これにより、抽出コードの氏名クリーンアップ機能をテストする。
    """
    first, last, lang_type = name_tuple
    
    # 20%の確率でノイズを挿入
    if random.random() < 0.2:
        if lang_type == 0: # 日本語名の場合
            if random.random() < 0.5:
                # スペース、全角スペース、中黒（・）を挿入
                return f"{first}{random.choice([' ', '　', '・'])}{last}"
            else:
                # 括弧書きのノイズ（除去されるべき情報）を追加
                return f"{first}{last}（現職：株式会社A）"
        else: # 英語名の場合
            if random.random() < 0.5:
                # 括弧書きのノイズを追加
                return f"{first} ({random.choice(['JAPAN', 'ENG'])}) {last}"
            else:
                # アンダースコア（_）ノイズを追加
                return f"{first}_{last}"
    
    # ノイズを挿入しない場合、元の形式を維持
    return f"{first}{last}" if lang_type == 0 else f"{first} {last}"


def generate_raw_data(num_records=NUM_RECORDS):
    """
    指定された件数分の試験用データレコードを生成する。
    ノイズの挿入、欠落のランダム化、順序のシャッフルを行う。
    """
    raw_records = []
    
    for i in range(1, num_records + 1):
        # 1. マスターデータのランダム生成
        name_tuple = random.choice(NAMES)
        name = add_noise_to_name(name_tuple) # ノイズ付き氏名を生成
        
        # 難易度向上: 項目を欠落させるかをランダムに決定
        is_age_missing = random.random() < 0.05       # 5%で年齢を欠落
        is_salary_missing = random.random() < 0.03    # 3%で単金を欠落
        is_skill_noise = random.random() < 0.1          # 10%でスキル表記にノイズを追加
        
        age = random.randint(20, 65)
        salary_man = random.randint(40, 150) # 万円単位の数値
        salary_unit = random.choice(SALARY_UNITS)
        skill_set = random.sample(LANGUAGES, random.randint(1, 5))
        industry = random.choice(INDUSTRIES)
        
        # 2. 抽出対象となる「本文」データの生成（フィールドのランダム化）
        field_parts = []
        
        # 氏名と性別
        field_parts.append(f"名 前: {name} ({random.choice(['男性', '女性'])})")
        field_parts.append(random.choice(NOISE)) # ランダムノイズを挿入

        # 年齢
        if not is_age_missing:
            field_parts.append(f"年 齢: {age} 歳")
        
        # 単金（表記ゆれとマスター値の決定）
        if not is_salary_missing:
            # カンマ区切り表記をランダムに適用
            salary_text = f"{salary_man:,}" if random.random() < 0.3 else str(salary_man)
            
            if '円' in salary_unit or '000k' in salary_unit:
                 # 本文に円単位で表記する場合
                 field_parts.append(f"単 金: {salary_man * 10000}{salary_unit}")
                 salary_master = str(salary_man * 10000) # マスターは円単位
            else:
                 # 本文に万円単位で表記する場合
                 field_parts.append(f"単 金: {salary_text} {salary_unit}")
                 salary_master = str(salary_man * 10000) # マスターは円単位
        else:
            salary_master = 'N/A'
            
        field_parts.append(random.choice(NOISE)) # ランダムノイズを挿入
            
        # スキル（表記ノイズをランダムに適用）
        skill_text = ', '.join(skill_set)
        if is_skill_noise:
            skill_text = f"【言 語】{skill_text} (必須)" # 抽出を妨害するノイズを追加
            
        field_parts.append(f"スキル:{skill_text}")
        field_parts.append(f"【業 務】{industry}システム")
        
        # 難易度向上: フィールドの順序をランダムに入れ替える
        random.shuffle(field_parts)
        
        # 生の本文文字列を構築
        body_text_raw = f"""
Subject: {random.choice(['【人材情報】', 'スキルシート', '経歴書'])} - {name}
-------------------------------------------------
{' '.join(field_parts)}
-------------------------------------------------
備考: 経験年数 : {random.randint(5, 15)}年。
        """
        
        # 3. 最終的な本文クリーンアップ（本文を一列に収めるための処理）
        clean_body_text = body_text_raw.strip()
        clean_body_text = re.sub(r'[\r\n\t"]', ' ', clean_body_text) # 改行、タブ、引用符をスペースに置換
        clean_body_text = re.sub(r'\s+', ' ', clean_body_text)       # 連続するスペースを一つに統一
        
        # 4. レコードの追加（問題データと正解マスターの両方を記録）
        raw_records.append({
            'EntryID': f'ID_{i:03d}',
            '件名': f'スキルシート送付 ({i})',
            '本文(テキスト形式)': clean_body_text,
            '宛先メール': 'sender@test.com',
            
            # 正解データ（マスター）の列を記録
            '名前_M': name,
            '年齢_M': str(age) if not is_age_missing else 'N/A', 
            '単金_M': salary_master,
            'スキルor言語_M': ','.join(skill_set),
            '業種_M': industry
        })
        
    # DataFrameの構築とクリーニング
    df_raw = pd.DataFrame(raw_records)
    df_raw = df_raw.fillna('N/A').astype(str)
    
    # カラム名を統一 ('業務_業種_M' -> '業種_M')
    return df_raw.rename(columns={'業務_業種_M': '業種_M'})


def export_dataframes_to_tsv(df_raw: pd.DataFrame): 
    """
    生成されたDataFrameから、問題ファイルと正解ファイルに分割し、タブ区切りCSV (TSV) として出力する。
    """
    # 評価に必要な列のリスト (正解データ用)
    MASTER_COLUMNS_EVAL_FOR_EXPORT = ['EntryID', '名前', '年齢', '単金', 'スキルor言語', '業種'] 
    
    # 1. 問題ファイル (question_source.csv) の作成: 抽出に必要な情報のみ
    df_question = df_raw[['EntryID', '件名', '本文(テキスト形式)', '宛先メール']].copy()
    
    # 2. 正解ファイル (master_answers.csv) の作成: 評価に必要な正解データのみ
    df_answer_cols = ['EntryID'] + [f'{col}_M' for col in MASTER_COLUMNS_EVAL_FOR_EXPORT if col != 'EntryID']
    df_answer = df_raw[df_answer_cols].copy()
    df_answer.columns = MASTER_COLUMNS_EVAL_FOR_EXPORT # カラム名を 'XXX_M' から 'XXX' に戻す
    
    # --- TSV保存ヘルパー関数 ---
    def save_tsv(df, path):
        try:
            # TSV形式で保存 (UTF-8 BOM付き, タブ区切り, 全て引用符で囲む)
            df.to_csv(
                path, 
                index=False, 
                encoding='utf-8-sig', 
                sep='\t', 
                quoting=csv.QUOTE_ALL
            )
            print(f"🎉 成功: ファイル '{path}' を出力しました。")
            return True
        except Exception as e:
            print(f"\n❌ 最終ファイル出力エラー: '{path}' の書き込みに失敗しました。詳細: {e}")
            return False
            
    # ファイル保存の実行
    success_q = save_tsv(df_question, INPUT_QUESTION_CSV)
    success_a = save_tsv(df_answer, MASTER_ANSWERS_PATH)

    # 処理結果の表示
    if success_q and success_a:
        print("\n========================================================")
        print(f"🎉 処理完了: {NUM_RECORDS}件の試験用データが正常に出力されました。")
    else:
        print("\n⚠️ 処理中断: ファイルの出力に失敗しました。")