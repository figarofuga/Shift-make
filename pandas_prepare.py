
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

#%%

all_columns = dat.columns


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
    tmp.append(data_wide)

                                    
    

# %%

tochoku_wide.merge(comment_dat, on='name', how='left').to_excel("tochoku_wide.xlsx")

#%%
ichijikyu_wide.merge(comment_dat, on='name', how='left').to_excel("ichijikyu_wide.xlsx")


# %%
