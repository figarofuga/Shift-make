
#%%
import numpy as np
import pandas as pd
import seaborn as sns
from datetime import date
import datetime
import re

#%%
# read excel data
dat = pd.read_excel("rawdata/2024_05answer.xlsx")
#%%

comment_dat = (dat.filter(['タイムスタンプ', 'お名前', '日直・当直希望についての備考', '1次救急希望についての備考',
       'このアンケート対する意見があればお願いします。'])
       .assign(タイムスタンプ=lambda x: pd.to_datetime(x['タイムスタンプ']))
                     .sort_values(by='タイムスタンプ', ascending=True)
                     .groupby('お名前')
                     .last()
                     .drop(['タイムスタンプ'], axis=1)
                     .reset_index()
                     .rename(columns={'お名前': 'name'})
)

#%%

name_column = 'お名前'
start_column = 'タイムスタンプ'
tochoku_stop_column = '日直・当直希望についての備考'
icijikyu_stop_column = '1次救急希望についての備考'
all_columns = dat.columns

name_index = all_columns.get_loc(name_column)
start_index = all_columns.get_loc(start_column)
tochoku_stop_index = all_columns.get_loc(tochoku_stop_column)
ichijikyu_stop_index = all_columns.get_loc(icijikyu_stop_column)

# 列名を変更する関数
def extract_date(column_name):
    match = re.search(r"\[(\d+/\d+\(.\))\]", column_name)
    if match:
        return match.group(1)  # 日付部分のみを返す
    return column_name  # マッチしない場合は元の列名を返す

# base data
base_data = (dat
             .filter(all_columns[start_index:(name_index + 1)])
             )

# subset of toschoku data
tochoku_data = (dat
             .filter(all_columns[(name_index+1):tochoku_stop_index])
             .rename(columns=extract_date)
             )

ichijikyu_data = (dat
                  .filter(all_columns[(tochoku_stop_index+1):ichijikyu_stop_index])
                  .rename(columns=extract_date)
)
#%%

# subset of ichijikyu data

tochoku_data_prep = (pd.concat([base_data, tochoku_data], axis=1)
                     .assign(タイムスタンプ=lambda x: pd.to_datetime(x['タイムスタンプ']))
                     .sort_values(by='タイムスタンプ', ascending=True)
                     .groupby('お名前')
                     .last()
                     .drop(['タイムスタンプ', 'メールアドレス', '現在在籍・研修中の科でお願いします'], axis=1)
                     .reset_index()
                     .rename(columns={'お名前': 'name'})
                     .melt(id_vars=['name'], 
                            var_name='date', 
                            value_name='request', 
                            ignore_index=False)
                     )


ichijikyu_data_prep = (pd.concat([base_data, ichijikyu_data], axis=1)
                     .assign(タイムスタンプ=lambda x: pd.to_datetime(x['タイムスタンプ']))
                     .sort_values(by='タイムスタンプ', ascending=True)
                     .groupby('お名前')
                     .last()
                     .drop(['タイムスタンプ', 'メールアドレス', '現在在籍・研修中の科でお願いします'], axis=1)
                     .reset_index()
                     .rename(columns={'お名前': 'name'})
                     .melt(id_vars=['name'], 
                            var_name='date', 
                            value_name='request', 
                            ignore_index=False)
                     )
 
# %%

dat_notes = pd.read_excel("rawdata/notes_data.xlsx")

# %%

notes_tochoku_data_prep = (dat_notes
                           .filter(['人', '\u3000日付', '日直・当直'])
                           .rename(columns={"人": "name", 
                                    "\u3000日付": "date",
                                    "日直・当直": "request"})
                               )
                   
                                    
                                    
 #%%                                   

tochoku_data = (pd.concat([tochoku_data_prep, notes_tochoku_data_prep], 
                         axis = 0)
                  .applymap(lambda x: x.strip() if isinstance(x, str) else x)   
                  .assign(request=lambda x: np.where(x['request'] == '', None, x['request']))   
                         )


# %%

notes_ichijikyu_data_prep = (dat_notes
                           .filter(['人', '\u3000日付', '一次救急'])
                           .rename(columns={"人": "name", 
                                    "\u3000日付": "date",
                                    "一次救急": "request"})
                           .applymap(lambda x: x.strip() if isinstance(x, str) else x)  
                           .assign( request=lambda x: np.where(x['request'] == '', None, x['request']))   
                         
                                    )



# notes_ichijikyu_data_prep['request'].unique().to_list()

ichijikyu_data = (pd.concat([ichijikyu_data_prep, notes_ichijikyu_data_prep], 
                         axis=0)
)


ichijikyu_wide = (ichijikyu_data
       .groupby(['name', 'request'])['date']
       .apply(lambda x: ' ,'.join(x))
       .reset_index()
       .pivot(index='name', columns='request', values='date'))

# %%

tochoku_wide = (tochoku_data
       .groupby(['name', 'request'])['date']
       .apply(lambda x: ' ,'.join(x))
       .reset_index()
       .pivot(index='name', columns='request', values='date')
       )

# %%

tochoku_wide.merge(comment_dat, on='name', how='left').to_excel("tochoku_wide.xlsx")

#%%
ichijikyu_wide.merge(comment_dat, on='name', how='left').to_excel("ichijikyu_wide.xlsx")


# %%
