
#%%
import numpy as np
import pandas as pd
import seaborn as sns
from datetime import date
import datetime
import re

#%%
# read excel data
dat = pd.read_excel("rawdata/2024_06answer.xlsx")

dat_notes = pd.read_excel("rawdata/2024_06notes_data.xlsx")

#%%

comment_dat = (dat.filter(['タイムスタンプ', 'お名前', '日直・当直希望についての備考', '1次救急希望についての備考',
       'ICU勤務希望についての備考', 'このアンケート対する意見があればお願いします。'])
       .assign(タイムスタンプ=lambda x: pd.to_datetime(x['タイムスタンプ']))
                     .sort_values(by='タイムスタンプ', ascending=True)
                     .groupby('お名前')
                     .last()
                     .drop(['タイムスタンプ'], axis=1)
                     .reset_index()
                     .rename(columns={'お名前': 'name'})
)

notes_comment_dat = (dat_notes.filter(['人', '日当直備考', '一次救急備考', 'ICU勤務備考'])
             .rename(columns={"人": "name", 
                               "日当直備考": "日直・当直希望についての備考",
                               "一次救急備考": "1次救急希望についての備考",
                               "ICU勤務備考": "ICU勤務希望についての備考"})
)

# 列名を変更する関数
def extract_date(column_name):
    match = re.search(r'\d+月\d+日', column_name)
    if match:
        return match.group()  # 日付部分のみを返す
    return column_name  # マッチしない場合は元の列名を返す

# base data
base_data = (dat
             .filter(['タイムスタンプ', 'お名前', 'あなたの専門はなんですか?', 'あなたは医師何年目ですか?'])
             )

# subset of toschoku data

tochoku_data = (dat
             .filter(regex = r'日直・当直希望.*\d月\d日')
             .rename(columns=extract_date)
             )

ICU_data = (dat
            .filter(regex = r'ICU勤務.*\d月\d日')
            .rename(columns=extract_date)
)

ichijikyu_data = (dat
                  .filter(regex = r'1次救急.*\d月\d日')
                  .rename(columns=extract_date)
)

#%%
tmp = []

spplist = ['tochoku', 'ichijikyu', 'ICU']
for i in spplist:
    if i == 'tochoku':
        regex = r'日直・当直希望.*\d月\d日'
    elif i == 'ichijikyu':
        regex = r'1次救急.*\d月\d日'
    elif i == 'ICU':
        regex = r'ICU勤務.*\d月\d日'
    data = (dat
             .filter(regex = regex)
             .rename(columns=extract_date)
             )
    prep_data = (pd.concat([base_data, data], axis=1)
                     .assign(タイムスタンプ=lambda x: pd.to_datetime(x['タイムスタンプ']))
                     .sort_values(by='タイムスタンプ', ascending=True)
                     .groupby('お名前')
                     .last()
                     .drop(['タイムスタンプ'], axis=1)
                     .reset_index()
                     .rename(columns={'お名前': 'name'})
                     .melt(id_vars=['name'], 
                            var_name='date', 
                            value_name='request', 
                            ignore_index=False)
                     )
    
    # Notesデータと結合

    if i == 'tochoku':
        colname = '日直・当直'
    elif i == 'ichijikyu':
        colname = '一次救急'
    elif i == 'ICU':
        colname = 'ICU勤務'

    notes_data = (dat_notes
             .filter(['人', '日付', colname])
             .rename(columns={"人": "name", 
                              "日付": "date",
                              colname: "request"}))
    combined_data = (pd.concat([prep_data, notes_data], axis=0)
                    .applymap(lambda x: x.strip() if isinstance(x, str) else x)   
                    .assign(request=lambda x: np.where(x['request'] == '', None, x['request']))   
                         )
    data_wide = (combined_data
             .groupby(['name', 'request'])['date']
             .apply(lambda x: ' ,'.join(x))
             .reset_index()
             .pivot(index='name', columns='request', values='date')    
    )
    columns_ok = data_wide.filter(items = ['○', '◯', '希望日']).columns
    if not columns_ok.empty:
        # 該当する列のデータをコンマ区切りで結合
        data_wide['accept'] = data_wide[columns_ok].apply(lambda x: ', '.join(x.dropna().astype(str)), axis=1)
    
    columns_bad = data_wide.filter(items=['×', '✕', '不可日']).columns
    if not columns_bad.empty:
        # 該当する列のデータをコンマ区切りで結合
        data_wide['reject'] = data_wide[columns_bad].apply(lambda x: ', '.join(x.dropna().astype(str)), axis=1)

    data_wide = data_wide.drop(columns=columns_ok).drop(columns=columns_bad)

    data_all = (data_wide
                .merge(comment_dat, on='name', how='left')
                .merge(notes_comment_dat, on='name', how='left')
                .assign(**{'日直・当直希望についての備考': lambda x: x['日直・当直希望についての備考_x'].combine_first(x['日直・当直希望についての備考_x'])})
                .assign(**{'1次救急希望についての備考': lambda x: x['1次救急希望についての備考_x'].combine_first(x['1次救急希望についての備考_x'])})
                .assign(ICU勤務希望についての備考=lambda x: x['ICU勤務希望についての備考_x'].combine_first(x['ICU勤務希望についての備考_x']))
                .drop(['日直・当直希望についての備考_x', '日直・当直希望についての備考_y', '1次救急希望についての備考_x', '1次救急希望についての備考_y', 'ICU勤務希望についての備考_x', 'ICU勤務希望についての備考_y'], axis = 1)
                .assign(species = i))
    
    data_all.to_excel(f"{i}_6.xlsx")

    tmp.append(data_all)

# %%

pd.concat(tmp).to_excel("tochoku_all_6.xlsx")

# %%
