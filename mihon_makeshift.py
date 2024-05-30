#%%
import pulp
import random

# 担当者リスト
employees = ['社員1', '社員2', '社員3', '社員4', '社員5', '社員6',
             'アルバイト1', 'アルバイト2', 'アルバイト3', 'アルバイト4', 'アルバイト5', 'アルバイト6']

# 社員とアルバイトの区別
full_timers = employees[:6]
part_timers = employees[6:]

# 希望日と不可日をランダムに設定
availability = {emp: {'希望日': random.sample(range(7), 2), '不可日': random.sample(range(7), 2)} for emp in employees}

# Pulpの問題を定義
prob = pulp.LpProblem('ShiftScheduling', pulp.LpMaximize)

# 変数の定義
x = pulp.LpVariable.dicts('shift', ((emp, day) for emp in employees for day in range(7)), cat='Binary')

# 各担当者のシフト回数をカウントする変数
shift_count = {emp: pulp.lpSum([x[(emp, day)] for day in range(7)]) for emp in employees}

# 最小および最大シフト回数の変数
min_shifts = pulp.LpVariable('min_shifts', lowBound=0, cat='Integer')
max_shifts = pulp.LpVariable('max_shifts', lowBound=0, cat='Integer')

# 目的関数
prob += pulp.lpSum([x[(emp, day)] for emp in employees for day in availability[emp]['希望日']])

# 各曜日のシフト人数の制約
for day in range(7):
    if day < 5:  # 平日
        prob += (pulp.lpSum([x[(emp, day)] for emp in employees]) == 3)
        prob += (pulp.lpSum([x[(emp, day)] for emp in full_timers]) >= 1)
        prob += (pulp.lpSum([x[(emp, day)] for emp in part_timers]) >= 1)
    else:  # 土日
        prob += (pulp.lpSum([x[(emp, day)] for emp in employees]) == 4)
        prob += (pulp.lpSum([x[(emp, day)] for emp in full_timers]) >= 1)
        prob += (pulp.lpSum([x[(emp, day)] for emp in part_timers]) >= 2)

# 不可日の制約
for emp in employees:
    for day in availability[emp]['不可日']:
        prob += (x[(emp, day)] == 0)

# 担当者の間隔制約
for emp in employees:
    for day in range(7):
        if day <= 2:
            prob += (pulp.lpSum([x[(emp, d)] for d in range(day, day + 4)]) <= 1)
        elif day >= 4:
            prob += (pulp.lpSum([x[(emp, d)] for d in range(day - 3, day + 1)]) <= 1)

# 最小および最大シフト回数の制約
for emp in employees:
    prob += (shift_count[emp] >= min_shifts)
    prob += (shift_count[emp] <= max_shifts)

# 最大シフト回数と最小シフト回数の差が2以下
prob += (max_shifts - min_shifts <= 2)

# 問題を解く
prob.solve()

# 結果の表示
for day in range(7):
    assigned = [emp for emp in employees if pulp.value(x[(emp, day)]) == 1]
    print(f'Day {day + 1}: {", ".join(assigned)}')

# 希望日と不可日を表示
print("\n希望日と不可日:")
for emp in employees:
    print(f'{emp} - 希望日: {availability[emp]["希望日"]}, 不可日: {availability[emp]["不可日"]}')
# %%
