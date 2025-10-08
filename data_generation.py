# data_generation.py

import pandas as pd
import random
import os
import csv
import re
from config import NUM_RECORDS, NAMES, LANGUAGES, INDUSTRIES, SALARY_UNITS, NOISE, INPUT_QUESTION_CSV, MASTER_ANSWERS_PATH


def add_noise_to_name(name_tuple):
    """名前にランダムなノイズ（スペース、特殊文字）を追加する"""
    first, last, lang_type = name_tuple
    
    if random.random() < 0.2:
        if lang_type == 0: 
            if random.random() < 0.5:
                return f"{first}{random.choice([' ', '　', '・'])}{last}"
            else:
                return f"{first}{last}（現職：株式会社A）"
        else:
             if random.random() < 0.5:
                return f"{first} ({random.choice(['JAPAN', 'ENG'])}) {last}"
             else:
                return f"{first}_{last}"
    return f"{first}{last}" if lang_type == 0 else f"{first} {last}"


def generate_raw_data(num_records=NUM_RECORDS):
    """難易度の高い試験用データを生成する。"""
    raw_records = []
    
    for i in range(1, num_records + 1):
        name_tuple = random.choice(NAMES)
        name = add_noise_to_name(name_tuple)
        
        is_age_missing = random.random() < 0.05
        is_salary_missing = random.random() < 0.03
        is_skill_noise = random.random() < 0.1
        
        age = random.randint(20, 65)
        salary_man = random.randint(40, 150)
        salary_unit = random.choice(SALARY_UNITS)
        skill_set = random.sample(LANGUAGES, random.randint(1, 5))
        industry = random.choice(INDUSTRIES)
        
        field_parts = []
        field_parts.append(f"名 前: {name} ({random.choice(['男性', '女性'])})")
        field_parts.append(random.choice(NOISE))

        if not is_age_missing:
            field_parts.append(f"年 齢: {age} 歳")
        
        if not is_salary_missing:
            salary_text = f"{salary_man:,}" if random.random() < 0.3 else str(salary_man)
            
            if '円' in salary_unit or '000k' in salary_unit:
                 field_parts.append(f"単 金: {salary_man * 10000}{salary_unit}")
                 salary_master = str(salary_man * 10000)
            else:
                 field_parts.append(f"単 金: {salary_text} {salary_unit}")
                 salary_master = str(salary_man * 10000)
        else:
            salary_master = 'N/A'
            
        field_parts.append(random.choice(NOISE))
            
        skill_text = ', '.join(skill_set)
        if is_skill_noise:
            skill_text = f"【言 語】{skill_text} (必須)"
            
        field_parts.append(f"スキル:{skill_text}")
        field_parts.append(f"【業 務】{industry}システム")
        
        random.shuffle(field_parts)
        
        body_text_raw = f"""
Subject: {random.choice(['【人材情報】', 'スキルシート', '経歴書'])} - {name}
-------------------------------------------------
{' '.join(field_parts)}
-------------------------------------------------
備考: 経験年数 : {random.randint(5, 15)}年。
        """
        
        clean_body_text = body_text_raw.strip()
        clean_body_text = re.sub(r'[\r\n\t"]', ' ', clean_body_text)
        clean_body_text = re.sub(r'\s+', ' ', clean_body_text)
        
        raw_records.append({
            'EntryID': f'ID_{i:03d}',
            '件名': f'スキルシート送付 ({i})',
            '本文(テキスト形式)': clean_body_text,
            '宛先メール': 'sender@test.com',
            
            '名前_M': name,
            '年齢_M': str(age) if not is_age_missing else 'N/A', 
            '単金_M': salary_master,
            'スキルor言語_M': ','.join(skill_set),
            '業種_M': industry
        })
        
    df_raw = pd.DataFrame(raw_records)
    df_raw = df_raw.fillna('N/A').astype(str)
    
    return df_raw.rename(columns={'業務_業種_M': '業種_M'})


def export_dataframes_to_tsv(df_raw: pd.DataFrame): 
    """データフレームをタブ区切りファイル（TSV）として出力する。"""
    MASTER_COLUMNS_EVAL_FOR_EXPORT = ['EntryID', '名前', '年齢', '単金', 'スキルor言語', '業種'] 
    
    df_question = df_raw[['EntryID', '件名', '本文(テキスト形式)', '宛先メール']].copy()
    
    df_answer_cols = ['EntryID'] + [f'{col}_M' for col in MASTER_COLUMNS_EVAL_FOR_EXPORT if col != 'EntryID']
    df_answer = df_raw[df_answer_cols].copy()
    df_answer.columns = MASTER_COLUMNS_EVAL_FOR_EXPORT 
    
    def save_tsv(df, path):
        try:
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
            
    success_q = save_tsv(df_question, INPUT_QUESTION_CSV)
    success_a = save_tsv(df_answer, MASTER_ANSWERS_PATH)

    if success_q and success_a:
        print("\n========================================================")
        print(f"🎉 処理完了: {NUM_RECORDS}件の試験用データが正常に出力されました。")
    else:
        print("\n⚠️ 処理中断: ファイルの出力に失敗しました。")