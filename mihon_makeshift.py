#%%
import re
import pandas as pd
#%%
month = 7

#%%

tochoku_member = ["福原誠一郎","太良史郎","林智史","持丸貴生","大重達寛","雪野満","吉田心慈","久冨木原健二","松永崇宏","新美望","篠﨑太郎","青木康浩","脇坂悠介","茅島敦人","津山頌章","川瀬咲","門松賢",
                  "東條誠也","莇舜平","荒金直美","佐川偲","藤澤まり",
                  "勝俣敬寛","松井貴裕","伊藤国秋","平井智大",
                #   "中瀬大","山本晴二郎","宗大輔","伊藤瑛基","中枝建郎","花岡孝行","安藤昂志",
                  "熊木聡美","織部峻太郎","臼坂優希","深沢夏海","松原龍輔","佐久川佳怜","吉田博道","佐久間一也","野上創生","髙木菜々美","谷岡友則","細川善弘","小宮健太郎","高梨航輔","福井梓穂","石井真央","岩崎文美","岡村真伊","黒崎颯","鈴木徹志郎","吉村梨沙","先﨑光","星貴文"
]

# read excel data
dat = pd.read_excel(f"rawdata/{month}m/2024_{month}answer.xlsx")
#%%
def extract_date(column_name):
    match = re.search(r'\d+月\d+日', column_name)
    if match:
        return match.group()  # 日付部分のみを返す
    return column_name  # マッチしない場合は元の列名を返す

# 当直のデータを整形
dat_tochoku = (dat
        .filter(regex=r'名前|タイムスタンプ|立場|日直・当直希望.*\d{1,2}月')
        .assign(タイムスタンプ=lambda x: pd.to_datetime(x['タイムスタンプ']))
        .sort_values(by='タイムスタンプ', ascending=True)
        .groupby('お名前')
        .last()
        .drop(['タイムスタンプ'], axis=1)
        .reset_index()
        .rename(columns={'お名前': 'name', 
                         'あなたの立場は？': 'position'})
        .assign(name = lambda x: x['name'].str.replace('[　 ]', '', regex=True))
        .rename(columns=lambda x: extract_date(x))
)
#%% 
# 各人の名前ごとに希望日と不可日のデータを整形
# ｛"川瀬咲": {"希望日": [1, 2, 3], "不可日": [4, 5, 6]}｝のようにする
# 必要なデータを抽出して辞書に整形する関数
def create_availability_dict(df):
    result = {}
    date_pattern = re.compile(r"7月\d+日")
    for index, row in df.iterrows():
        name = row["name"]
        availability = {"希望日": [], "不可日": []}
        for col in df.columns:
            if date_pattern.match(col):
                day = int(col.replace("7月", "").replace("日", ""))
                if row[col] == "希望日":
                    availability["希望日"].append(day)
                elif row[col] == "不可日":
                    availability["不可日"].append(day)
        # 各値を-1する
        availability["希望日"] = [day - 1 for day in availability["希望日"]]
        availability["不可日"] = [day - 1 for day in availability["不可日"]]
        result[name] = availability
    return result


# 辞書データの作成
availability_dict = create_availability_dict(dat_tochoku)

#%%
all_members = (dat_tochoku
 .filter(items=['name'])
 .drop_duplicates()
 .values.flatten().tolist()
)

staffs = (dat_tochoku
          .query('position == "スタッフ"')
          .filter(items=['name'])
          .drop_duplicates()
          .values.flatten().tolist()
)

residents = (dat_tochoku
            .query('position == "レジデント"')
            .filter(items=['name'])
            .drop_duplicates()
            .values.flatten().tolist()
    )

days_in_month = (dat_tochoku
        .filter(regex=r'\d月\d{1,2}日')
        .shape[1])

holidays = [5, 6, 12, 13, 19, 20, 26, 27]

#%%
import pulp
import random
#%%
# 担当者リスト
# all_members = ['社員1', '社員2', '社員3', '社員4', '社員5', '社員6',
#              'アルバイト1', 'アルバイト2', 'アルバイト3', 'アルバイト4', 'アルバイト5', 'アルバイト6']

# # 社員とアルバイトの区別
# staffs = all_members[:6]
# part_timers = all_members[6:]

# 希望日と不可日をランダムに設定する関数
# def generate_availability():
#     days = list(range(7))
#     hope_days = random.sample(days, 2)
#     remaining_days = [day for day in days if day not in hope_days]
#     no_days = random.sample(remaining_days, 2)
#     return hope_days, no_days

# availability = {emp: {'希望日': [], '不可日': []} for emp in all_members}
# for emp in all_members:
#     availability[emp]['希望日'], availability[emp]['不可日'] = generate_availability()

# Pulpの問題を定義
prob = pulp.LpProblem('ShiftScheduling', pulp.LpMinimize)

# 変数の定義
x = pulp.LpVariable.dicts('shift', ((emp, day) for emp in all_members for day in range(days_in_month)), cat='Binary')

# 各担当者のシフト回数をカウントする変数
shift_count = {emp: pulp.lpSum([x[(emp, day)] for day in range(days_in_month)]) for emp in all_members}

# 最小および最大シフト回数の変数
min_shifts = pulp.LpVariable('min_shifts', lowBound=0, cat='Integer')
max_shifts = pulp.LpVariable('max_shifts', lowBound=0, cat='Integer')

# 目的関数: シフト回数の差を最小化する
prob += max_shifts - min_shifts

# 各曜日のシフト人数の制約
for day in range(days_in_month):
    if day in holidays:  # holiday
        prob += (pulp.lpSum([x[(emp, day)] for emp in all_members]) == 3)
        prob += (pulp.lpSum([x[(emp, day)] for emp in staffs]) >= 1)
        prob += (pulp.lpSum([x[(emp, day)] for emp in residents]) >= 1)
    else:  # weekday
        prob += (pulp.lpSum([x[(emp, day)] for emp in all_members]) == 4)
        prob += (pulp.lpSum([x[(emp, day)] for emp in staffs]) >= 1)
        prob += (pulp.lpSum([x[(emp, day)] for emp in residents]) >= 2)

#%%
# 不可日の制約
for emp in all_members:
    for day in availability_dict[emp]['不可日']:
        prob += (x[(emp, day)] == 0)

# 担当者の間隔制約
# Todo ここを直す
for emp in all_members:
    for day in range(days_in_month):
        if day < 4:
            prob += pulp.lpSum(x[(emp, d)] for d in range(0, day+5)) <= 1
        elif day > 4:
            prob += pulp.lpSum(x[(emp, d)] for d in range(day-4, min(7, day+5))) <= 1
        else:
            prob += pulp.lpSum(x[(emp, d)] for d in range(day-4, 7)) <= 1

# 最小および最大シフト回数の制約
for emp in all_members:
    prob += (shift_count[emp] >= min_shifts)
    prob += (shift_count[emp] <= max_shifts)

# 最大シフト回数と最小シフト回数の差が2以下
prob += (max_shifts - min_shifts <= 2)

#%%
# 問題を解く
prob.solve()

# 結果の表示
for day in range(days_in_month):
    assigned = [emp for emp in all_members if pulp.value(x[(emp, day)]) == 1]
    print(f'Day {day + 1}: {", ".join(assigned)}')

# 希望日と不可日を表示
print("\n希望日と不可日:")
for emp in all_members:
    print(f'{emp} - 希望日: {availability_dict[emp]["希望日"]}, 不可日: {availability_dict[emp]["不可日"]}')

# 各担当者のシフト回数の表示
print("\n各担当者のシフト回数:")
for emp in all_members:
    print(f'{emp}: {pulp.value(shift_count[emp])}回')

# %%
