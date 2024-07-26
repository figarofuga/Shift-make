#%%
import re
import pandas as pd
from ortools.linear_solver import pywraplp

# データの読み込みと前処理
month = 9
dat_proto = pd.read_excel(f"rawdata/{month}m/2024年{month}月の日当直・ICU勤務・一次救急希望（回答）.xlsx")

dat = (dat_proto
       .assign(name=lambda x: x['お名前'].str.replace('[　 ]', '', regex=True))
       .assign(doctoryear=lambda x: x['あなたは医師何年目ですか？'].replace(r'年.*', '', regex=True).astype(int))
       .rename(columns={'C日直以外の、日当直の該当者ですか？': 'tochoku_yn'})
)

def extract_date(column_name):
    match = re.search(r'\d+月\d+日', column_name)
    if match:
        return match.group()  # 日付部分のみを返す
    return column_name  # マッチしない場合は元の列名を返す

dat_tochoku = (dat
        .query('tochoku_yn == "はい"')
        .filter(regex=r'名前|タイムスタンプ|立場|日直・当直希望.*\d{1,2}月')
        .assign(タイムスタンプ=lambda x: pd.to_datetime(x['タイムスタンプ']))
        .sort_values(by='タイムスタンプ', ascending=True)
        .groupby('お名前')
        .last()
        .drop(['タイムスタンプ'], axis=1)
        .reset_index()
        .rename(columns={'お名前': 'name', 'あなたの立場は？': 'position'})
        .assign(name=lambda x: x['name'].str.replace('[　 ]', '', regex=True))
        .rename(columns=lambda x: extract_date(x))
)

def create_availability_dict(df):
    result = {}
    date_pattern = re.compile(fr"{month}月\d+日")
    for index, row in df.iterrows():
        name = row["name"]
        availability = {"希望日": [], "不可日": []}
        for col in df.columns:
            if date_pattern.match(col):
                day = int(col.replace(f"{month}月", "").replace("日", ""))
                if row[col] == "希望日":
                    availability["希望日"].append(day)
                elif row[col] == "不可日":
                    availability["不可日"].append(day)
        availability["希望日"] = [day - 1 for day in availability["希望日"]]
        availability["不可日"] = [day - 1 for day in availability["不可日"]]
        result[name] = availability
    return result

availability_dict = create_availability_dict(dat_tochoku)

all_members = dat_tochoku['name'].unique().tolist() + ["東絛誠也"]
staffs = dat_tochoku.query('position == "スタッフ"')['name'].unique().tolist()
residents = dat_tochoku.query('position == "レジデント"')['name'].unique().tolist()
days_in_month = dat_tochoku.filter(regex=r'\d月\d{1,2}日').shape[1]
holidays = [5, 6, 12, 13, 19, 20, 26, 27]

# Google OR-Tools の Solver の作成
solver = pywraplp.Solver.CreateSolver('SCIP')

# 変数の定義
x = {}
for emp in all_members:
    for day in range(days_in_month):
        x[(emp, day)] = solver.IntVar(0, 1, f'shift_{emp}_{day}')

# 各担当者のシフト回数をカウントする変数
shift_count = {}
for emp in all_members:
    shift_count[emp] = solver.Sum(x[(emp, day)] for day in range(days_in_month))

# 最小および最大シフト回数の変数
min_shifts = solver.IntVar(0, days_in_month, 'min_shifts')
max_shifts = solver.IntVar(0, days_in_month, 'max_shifts')

# 目的関数: シフト回数の差を最小化する
objective = solver.Objective()
objective.SetCoefficient(max_shifts, 1)
objective.SetCoefficient(min_shifts, -1)
objective.SetMinimization()

# 各曜日のシフト人数の制約
for day in range(days_in_month):
    if day + 1 in holidays:  # holiday
        solver.Add(solver.Sum(x[(emp, day)] for emp in all_members) == 3)
        solver.Add(solver.Sum(x[(emp, day)] for emp in staffs) >= 1)
        solver.Add(solver.Sum(x[(emp, day)] for emp in residents) >= 1)
    else:  # weekday
        solver.Add(solver.Sum(x[(emp, day)] for emp in all_members) == 4)
        solver.Add(solver.Sum(x[(emp, day)] for emp in staffs) >= 1)
        solver.Add(solver.Sum(x[(emp, day)] for emp in residents) >= 2)

# 不可日の制約
for emp in all_members:
    for day in availability_dict[emp]['不可日']:
        solver.Add(x[(emp, day)] == 0)

# 担当者の間隔制約
for emp in all_members:
    for day in range(days_in_month):
        if day < 4:
            solver.Add(solver.Sum(x[(emp, d)] for d in range(0, day + 5)) <= 1)
        elif day > 4:
            solver.Add(solver.Sum(x[(emp, d)] for d in range(day - 4, min(7, day + 5))) <= 1)
        else:
            solver.Add(solver.Sum(x[(emp, d)] for d in range(day - 4, 7)) <= 1)

# 最小および最大シフト回数の制約
for emp in all_members:
    solver.Add(shift_count[emp] >= min_shifts)
    solver.Add(shift_count[emp] <= max_shifts)

# 最大シフト回数と最小シフト回数の差が2以下
solver.Add(max_shifts - min_shifts <= 2)

# 問題を解く
status = solver.Solve()

# 結果の表示
if status == pywraplp.Solver.OPTIMAL:
    for day in range(days_in_month):
        assigned = [emp for emp in all_members if x[(emp, day)].solution_value() == 1]
        print(f'Day {day + 1}: {", ".join(assigned)}')

    print("\n希望日と不可日:")
    for emp in all_members:
        print(f'{emp} - 希望日: {availability_dict[emp]["希望日"]}, 不可日: {availability_dict[emp]["不可日"]}')

    print("\n各担当者のシフト回数:")
    for emp in all_members:
        print(f'{emp}: {shift_count[emp].solution_value()}回')
else:
    print("Optimal solution not found.")
