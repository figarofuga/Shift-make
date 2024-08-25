
# yearmonth = "202409"
# add = [1, 2, 3, 4, 5]
# remove = [1, 2, 3, 4, 5]
# no_answered_residents = ["東條", "吉村"]
# no_answered_staffs = ["松本", "松田"]

# %%
import re
import datetime
import jpholiday
from ortools.linear_solver import pywraplp
import pandas as pd

def makeholidays(yearmonth):
    # 年と月を取得
    year = int(yearmonth[:4])
    month = int(yearmonth[4:])

    # 月の最初と最後の日を取得
    start_date = datetime.date(year, month, 1)
    if month == 12:
        end_date = datetime.date(year + 1, 1, 1) - datetime.timedelta(days=1)
    else:
        end_date = datetime.date(year, month + 1, 1) - datetime.timedelta(days=1)
    
    # 日付のリストを作成
    holidays = []
    current_date = start_date
    while current_date <= end_date:
        if current_date.weekday() >= 5 or jpholiday.is_holiday(current_date):
            holidays.append(current_date.day)
        current_date += datetime.timedelta(days=1)
    
    holidays.sort()  # 日付を昇順にソート
    return holidays

# October: [5, 6, 12, 13, 14, 19, 20, 26, 27]
def makeshift(yearmonth, holidays, no_answered_residents=None, no_answered_staffs=None):
    # 引数の初期化
    if no_answered_residents is None:
        no_answered_residents = []
    if no_answered_staffs is None:
        no_answered_staffs = []

    # 年と月を取得
    year = int(yearmonth[:4])
    month = int(yearmonth[4:])
    
    # データの読み込みと前処理
    path = f"rawdata/{month}m/2024年{month}月の日当直・ICU勤務・一次救急希望（回答）.xlsx"
    dat_proto = pd.read_excel(path)
    
    
    # データフレームの前処理
    dat = (dat_proto
           .assign(name=lambda x: x['お名前'].str.replace('[　 ]', '', regex=True))
           .assign(doctoryear=lambda x: x['あなたは医師何年目ですか？'].replace(r'年.*', '', regex=True).astype(int))
           .rename(columns={'C日直以外の日・当直の該当者ですか？': 'tochoku_yn'})
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

    answered = dat_tochoku['name'].values.tolist()

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

    answered_staffs = dat_tochoku.query('position == "スタッフ"')['name'].unique().tolist()
    staffs = answered_staffs + no_answered_staffs

    answered_residents = dat_tochoku.query('position == "レジデント/フェロー"')['name'].unique().tolist()
    residents = answered_residents + no_answered_residents

    all_members = staffs + residents

    days_in_month = dat_tochoku.filter(regex=r'\d月\d{1,2}日').shape[1]

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
    for emp in all_members:
        for day in range(days_in_month):
            if day in availability_dict.get(emp, {}).get('希望日', []):
                objective.SetCoefficient(x[(emp, day)], 1)
            if day in availability_dict.get(emp, {}).get('不可日', []):
                objective.SetCoefficient(x[(emp, day)], -500)
    objective.SetMaximization()

    # 各曜日のシフト人数の制約
    for day in range(days_in_month):
        if day + 1 in holidays:  # holiday
            solver.Add(solver.Sum(x[(emp, day)] for emp in all_members) == 4)
            solver.Add(solver.Sum(x[(emp, day)] for emp in staffs) >= 1)
            solver.Add(solver.Sum(x[(emp, day)] for emp in residents) >= 2)
        else:  # weekday
            solver.Add(solver.Sum(x[(emp, day)] for emp in all_members) == 3)
            solver.Add(solver.Sum(x[(emp, day)] for emp in staffs) >= 1)
            solver.Add(solver.Sum(x[(emp, day)] for emp in residents) >= 1)

    # 不可日の制約
    for emp in all_members:
        # 不可日のリストを取得。存在しない場合は空のリストを返す。
        unavailable_days = availability_dict.get(emp, {}).get('不可日', [])
        for day in unavailable_days:
            solver.Add(x[(emp, day)] == 0)

    # 担当者の間隔制約
    for emp in all_members:
        for day in range(days_in_month):
            if day >= 5:
                # シフトが5日以上開く制約を追加
                solver.Add(solver.Sum(x[(emp, d)] for d in range(day - 5, day)) <= 1)


    # 最小および最大シフト回数の制約
    for emp in all_members:
        solver.Add(shift_count[emp] >= min_shifts)
        solver.Add(shift_count[emp] <= max_shifts)

    # 最大シフト回数と最小シフト回数の差が2以下
    solver.Add(max_shifts - min_shifts <= 2)

    # 問題を解く
    status = solver.Solve()
    # 結果の表示と保存

    # 結果を格納するための空のリストを作成
    shift_assignments = []
    availability_summary = []
    shift_counts = []
    penalty_details = []
    total_penalty = 0

    if status == pywraplp.Solver.OPTIMAL or status == pywraplp.Solver.FEASIBLE:
        for day in range(days_in_month):
            assigned = [emp for emp in all_members if x[(emp, day)].solution_value() == 1]
            shift_assignments.append({'Day': day + 1, 'Assigned': ", ".join(assigned)})

            # ペナルティの計算
            for emp in all_members:
                if x[(emp, day)].solution_value() == 1:
                    if day in availability_dict.get(emp, {}).get('不可日', []):
                        penalty_details.append({'Name': emp, 'Day': day + 1, 'Penalty': -500})
                        total_penalty += -500
                    elif day in availability_dict.get(emp, {}).get('希望日', []):
                        penalty_details.append({'Name': emp, 'Day': day + 1, 'Penalty': 1})
                        total_penalty += 1

        shift_assignments_df = pd.DataFrame(shift_assignments)
    
        for emp in all_members:
            希望日 = availability_dict.get(emp, {}).get('希望日', [])
            不可日 = availability_dict.get(emp, {}).get('不可日', [])
            availability_summary.append({'Name': emp, '希望日': 希望日, '不可日': 不可日})

        availability_summary_df = pd.DataFrame(availability_summary)
    
        for emp in all_members:
            shift_counts.append({'Name': emp, 'Shift Count': shift_count[emp].solution_value()})
        
        shift_counts_df = pd.DataFrame(shift_counts)
    
        print("Shift Assignments:")
        print(shift_assignments_df)

        print("\nAvailability Summary:")
        print(availability_summary_df)

        print("\nShift Counts:")
        print(shift_counts_df)
    
    # 不可日に勤務が割り当てられた担当者の名前とその日付を出力
        penalty_details_df = pd.DataFrame(penalty_details)
        print("\nPenalty Details:")
        print(penalty_details_df)

        print(f"\nTotal Penalty: {total_penalty}")

    else:
        print("No feasible solution found.")

    return {"shift_assignment": shift_assignments_df, "shift_count": shift_counts_df, "pelaity_details": penalty_details_df}

# %%
res_10m = makeshift(yearmonth="202410", holidays=[5, 6, 12, 13, 14, 19, 20, 26, 27], no_answered_residents=["小宮"])
# %%
# Excel writerを使用してファイルに書き込む
with pd.ExcelWriter('res_10m.xlsx') as writer:
    for key, df in res_10m.items():
        df.to_excel(writer, sheet_name=key, index=False)
        
        
# %%
