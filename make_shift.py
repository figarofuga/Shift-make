#%%
import pandas as pd
import re
from pulp import LpProblem, LpVariable, lpSum, LpMinimize, LpStatus, value

#%%
# 列名を変更する関数
def extract_date(column_name):
    match = re.search(r'\d+月\d+日', column_name)
    if match:
        return match.group()  # 日付部分のみを返す
    return column_name  # マッチしない場合は元の列名を返す

#%%
# ファイルを読み込みます
duty_schedule_df = pd.read_excel('2024年7月当直表_test.xlsx')
duty_preferences_df_pre = pd.read_excel('2024年7月_res_0527.xlsx')

# "お名前"列でGroupbyしてタイムスタンプで最も最近の行を取得します
# "お名前", "あなたの立場は？", r'日直・当直希望.*\d{1,2}月'の列のみ選択します

duty_preferences_df = (duty_preferences_df_pre
                       .assign(タイムスタンプ=lambda x: pd.to_datetime(x['タイムスタンプ']))
                       .sort_values(by='タイムスタンプ', ascending=True)
                       .groupby('お名前')
                       .last()
                       .reset_index()
                       .filter(regex=r'お名前|あなたの立場は？|日直・当直希望.*\d{1,2}月')  
                       .rename(columns=extract_date))

# レジデントとスタッフを分けます
residents = duty_preferences_df[duty_preferences_df['あなたの立場は？'] == 'レジデント']
staffs = duty_preferences_df[duty_preferences_df['あなたの立場は？'] == 'スタッフ']

#%%
# スケジューリング問題を定義します
prob = LpProblem("DutyScheduling", LpMinimize)

# 当直表のサイズを取得します
rows, cols = duty_schedule_df.shape

# 各セルごとの変数を作成します
duty_vars = {}
for row in range(rows):
    for col in range(cols):
        if duty_schedule_df.iloc[row, col] in [1, 2, 3]:
            for person in duty_preferences_df['お名前']:
                duty_vars[(row, col, person)] = LpVariable(f"duty_{row}_{col}_{person}", 0, 1, cat='Binary')

# 列名を取得します
preference_columns = duty_preferences_df.columns[2:].tolist()

# 目的関数（希望日の当直回数を最大化）を定義します
objective_terms = []
for person in duty_preferences_df['お名前']:
    for row in range(rows):
        for col in range(cols):
            if duty_schedule_df.iloc[row, col] in [1, 2, 3]:
                col_name = preference_columns[col + 1]  # Adjusting for the 0-index and assuming the first column is 'お名前'
                if duty_preferences_df[duty_preferences_df['お名前'] == person].iloc[0][col_name] == '希望日':
                    objective_terms.append(duty_vars[(row, col, person)])

prob += lpSum(objective_terms)

# 各セルに1人の担当者を割り当てます
for row in range(rows):
    for col in range(cols):
        if duty_schedule_df.iloc[row, col] in [1, 2, 3]:
            prob += lpSum(duty_vars[(row, col, person)] for person in duty_preferences_df['お名前']) == 1

# 当直の種類ごとに担当者を制限します
for row in range(rows):
    for col in range(cols):
        if duty_schedule_df.iloc[row, col] == 1:
            prob += lpSum(duty_vars[(row, col, person)] for person in residents['お名前']) == 1
        elif duty_schedule_df.iloc[row, col] == 2:
            prob += lpSum(duty_vars[(row, col, person)] for person in staffs['お名前']) == 1
        elif duty_schedule_df.iloc[row, col] == 3:
            prob += lpSum(duty_vars[(row, col, person)] for person in duty_preferences_df['お名前']) == 1

# 不可日に当直を割り当てないようにします
for person in duty_preferences_df['お名前']:
    for row in range(rows):
        for col in range(1, cols):
            col_name = duty_schedule_df.columns[col]
            if col_name in duty_preferences_df.columns and duty_schedule_df.iloc[row, col] in [1, 2, 3] and duty_preferences_df[duty_preferences_df['お名前'] == person].iloc[0][col_name] == '不可日':
                prob += duty_vars[(row, col, person)] == 0

# 各人の当直回数を均等にします
total_duties_per_person = {person: lpSum(duty_vars[(row, col, person)] for row in range(rows) for col in range(1, cols) if duty_schedule_df.iloc[row, col] in [1, 2, 3]) for person in duty_preferences_df['お名前']}
average_duties = lpSum(total_duties_per_person[person] for person in total_duties_per_person) / len(duty_preferences_df['お名前'])
for person in total_duties_per_person:
    prob += total_duties_per_person[person] <= average_duties + 1
    prob += total_duties_per_person[person] >= average_duties - 1

#%%
# 同一人物の当直は少なくとも3日間あける制約
# 同一人物が同じ日に複数の当直を持たないようにする制約
for person in duty_preferences_df['お名前']:
    for row in range(rows):
        prob += lpSum(duty_vars[(row, col, person)] for col in range(cols) if duty_schedule_df.iloc[row, col] in [1, 2, 3]) <= 1

# 同一人物の当直は少なくとも3日間あける制約
for person in duty_preferences_df['お名前']:
    for row in range(rows - 2):  # -2 to ensure we have enough days to check
        prob += lpSum(
            duty_vars[(row + offset, col, person)]
            for offset in range(3)
            for col in range(cols)
            if (row + offset) < rows and duty_schedule_df.iloc[row + offset, col] in [1, 2, 3]
        ) <= 1

#%%
# 問題を解きます
prob.solve()

# 結果を表示します
assigned_duties = {}
for person in duty_preferences_df['お名前']:
    assigned_duties[person] = []
    for row in range(rows):
        for col in range(cols):
            if duty_schedule_df.iloc[row, col] in [1, 2, 3] and value(duty_vars[(row, col, person)]) == 1:
                assigned_duties[person].append((preference_columns[col + 1], row, col))

# 結果を出力します
output = {}
for person in assigned_duties:
    matching_wishes = 0
    conflicting_days = 0
    for duty, row, col in assigned_duties[person]:
        col_name = preference_columns[col + 1]
        if duty_preferences_df[duty_preferences_df['お名前'] == person].iloc[0][col_name] == '希望日':
            matching_wishes += 1
        if duty_preferences_df[duty_preferences_df['お名前'] == person].iloc[0][col_name] == '×':
            conflicting_days += 1
    output[person] = {
        "duties": assigned_duties[person],
        "matching_wishes": matching_wishes,
        "conflicting_days": conflicting_days
    }

for person in output:
    print(f"{person}: {output[person]['duties']}, {output[person]['matching_wishes']}, {output[person]['conflicting_days']}")
# %%
# 辞書をDataFrameに変換
df_output = pd.DataFrame(output)

# CSVファイルに書き出し
df_output.to_excel('output.xlsx')
# %%
