#%%
import pandas as pd
import re
from pulp import LpProblem, LpMinimize, LpVariable, lpSum, LpStatus

#%%
# Load the Excel files
june_schedule_df = pd.read_excel('2024年6月当直表.xlsx')
preferences_df = pd.read_excel('当直希望データ.xlsx')

#%%
# Extract residents and staff
residents_df = preferences_df[preferences_df['あなたの立場は？'] == 'レジデント']
staff_df = preferences_df[preferences_df['あなたの立場は？'] == 'スタッフ']


# Extract necessary columns for processing
# 曜日部分を削除
june_schedule_df['Cleaned_Date'] = june_schedule_df['2024年6月当直表'].apply(lambda x: re.sub(r'\([^)]*\)', '', x).strip())
# 年を追加して日付を解析
june_schedule_df['Date'] = pd.to_datetime('2024年' + june_schedule_df['Cleaned_Date'], format='%Y年%m月%d日')

# Convert preference data to a suitable format
preferences_data = preferences_df.set_index('お名前').T.drop(['あなたの立場は？', 'あなたは医師何年目ですか？'])
# 曜日部分を削除
preferences_data.index = preferences_data.index.to_series().apply(lambda x: re.sub(r'\([^)]*\)', '', x))
# 年を追加して日付を解析
preferences_data.index = pd.to_datetime('2024年' + preferences_data.index, format='%Y年%m月%d日')

#%%
# Define the problem
problem = LpProblem("OnCallScheduling", LpMinimize)

# Extract unique dates and duty types from the schedule
dates = june_schedule_df['Date'].unique()
duty_types = ['E当直', 'F当直', 'A当直', 'B当直', '病棟当直', '教育当直']

# Define decision variables
on_call = LpVariable.dicts("on_call", 
                           ((person, date, duty) for person in preferences_df['お名前'] for date in dates for duty in duty_types), 
                           cat='Binary')

# Add constraints
# No duties on forbidden days
for person in preferences_df['お名前']:
    for date in dates:
        if date in preferences_data.index and preferences_data.loc[date, person] == '不可日':
            for duty in duty_types:
                problem += on_call[(person, date, duty)] == 0

# Balanced distribution of duties
for person in preferences_df['お名前']:
    problem += lpSum(on_call[(person, date, duty)] for date in dates for duty in duty_types) <= 10
    problem += lpSum(on_call[(person, date, duty)] for date in dates for duty in duty_types) >= 8

# Minimum 3-day gap between same person's duties
for person in preferences_df['お名前']:
    for i in range(len(dates) - 3):
        for duty in duty_types:
            problem += lpSum(on_call[(person, dates[j], duty)] for j in range(i, i + 4)) <= 1

# Prioritize preferred days
problem += lpSum(on_call[(person, date, duty)] for person in preferences_df['お名前'] for date in dates if date in preferences_data.index and preferences_data.at[date, person] == '希望日' for duty in duty_types)

# Solve the problem
problem.solve()

# Extract and format the results
assigned_duties = {}
for v in problem.variables():
    if v.varValue == 1:
        person, date, duty = v.name.split('_')[2:]
        date = pd.to_datetime(date)
        if person not in assigned_duties:
            assigned_duties[person] = []
        assigned_duties[person].append(f"{duty}: {date.strftime('%Y-%m-%d')}")

# Display the output
output = []
for person, duties in assigned_duties.items():
    output.append(f"{person}: {', '.join(duties)}")

print("\n".join(output))
# %%
