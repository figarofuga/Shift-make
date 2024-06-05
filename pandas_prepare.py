
#%%
import numpy as np
import pandas as pd
import datetime
import re
import pickle
#%% 
month = 7
ok_sign = r'◯|〇|◯１|希望日'
ng_sign = r'×|✕|不可日'
#%%
# read excel data
dat = pd.read_excel(f"rawdata/{month}m/2024_{month}answer.xlsx")

dat_notes = (pd.read_excel(f"rawdata/{month}m/2024_{month}notes_data.xlsx", header=1)
             .dropna(how='all', axis=1)
)

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
#%%

kibou_df = (dat.filter(regex=r'タイムスタンプ|お名前|日直・当直希望.*\d{1,2}月')
       .rename(columns=extract_date)
       .melt(id_vars=['タイムスタンプ', 'お名前'], 
             var_name='date', value_name='request')
       .dropna(axis=0, subset=['request'])
       .query('request=="希望日"')
       .groupby('date')['お名前']
       .agg(' ,'.join)
       .reset_index()
)

fukabi_df = (dat.filter(regex=r'タイムスタンプ|お名前|日直・当直希望.*\d{1,2}月')
       .rename(columns=extract_date)
       .melt(id_vars=['タイムスタンプ', 'お名前'], 
             var_name='date', value_name='request')
       .dropna(axis=0, subset=['request'])
       .query('request=="不可日"')
       .groupby('date')['お名前']
       .agg(' ,'.join)
       .reset_index()
)

#%%
# 当直のデータを整形
data_tochoku = (dat
        .filter(regex=r'日直・当直希望.*\d{1,2}月')
        .rename(columns=extract_date)
)
data_concat_tochoku = (pd.concat([base_data, data_tochoku], axis=1)
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
# notes dataと結合

notes_data_tochoku = (dat_notes
                .filter(regex=r'人|日付|日直・当直.*')
                .rename(columns=lambda x: 'name' if '人' in x else x)
                .rename(columns=lambda x: 'date' if '日付' in x else x)
                .rename(columns=lambda x: 'request' if '日直・当直希望' in x and '備考' not in x else x)
)

combined_data_tochoku = (pd.concat([data_concat_tochoku, notes_data_tochoku], axis=0)
                            .applymap(lambda x: x.strip() if isinstance(x, str) else x)   
                            .assign(request=lambda x: np.where(x['request'] == '', None, x['request']))   
                            .assign(name = lambda x: x['name'].str.replace('[　 ]', '', regex=True))
                            )

data_wide_tochoku_pre = (combined_data_tochoku
             .groupby(['name', 'request'])['date']
             .apply(lambda x: ' ,'.join(x))
             .pivot(index='name', columns='request', values='date')
             .reset_index()
    )
columns_ok = data_wide_tochoku_pre.filter(regex=ok_sign).columns
if not columns_ok.empty:
    # 該当する列のデータをコンマ区切りで結合
    data_wide_tochoku_pre['accept'] = data_wide_tochoku_pre[columns_ok].apply(lambda x: ', '.join(x.dropna().astype(str)), axis=1)
    
columns_bad = data_wide_tochoku_pre.filter(regex=ng_sign).columns
if not columns_bad.empty:
    # 該当する列のデータをコンマ区切りで結合
    data_wide_tochoku_pre['reject'] = data_wide_tochoku_pre[columns_bad].apply(lambda x: ', '.join(x.dropna().astype(str)), axis=1)

data_wide_tochoku = (data_wide_tochoku_pre
                 .drop(columns=columns_ok)
                 .drop(columns=columns_bad)
                 .assign(species = 'tochoku')
    )

comment_tochoku = comment_dat.filter(regex=r'name|日直・当直')

data_wide_tochoku_comment = (data_wide_tochoku
                 .merge(comment_tochoku, on='name', how='left')
                 .assign(comment = lambda df: df['日直・当直希望'].combine_first(df['日直・当直希望についての備考']))
            )

#%%
# ICU勤務のデータを整形
data_icu = (dat
        .filter(regex=r'ICU勤務.*\d{1,2}月')
        .rename(columns=extract_date)
)
data_concat_icu = (pd.concat([base_data, data_icu], axis=1)
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
# notes dataと結合

notes_data_icu = (dat_notes
                .filter(regex=r'人|日付|ICU.*')
                .rename(columns=lambda x: 'name' if '人' in x else x)
                .rename(columns=lambda x: 'date' if '日付' in x else x)
                .rename(columns=lambda x: 'request' if 'ICU' in x and '備考' not in x else x)
)

combined_data_icu = (pd.concat([data_concat_icu, notes_data_icu], axis=0)
                            .applymap(lambda x: x.strip() if isinstance(x, str) else x)   
                            .assign(request=lambda x: np.where(x['request'] == '', None, x['request']))   
                            .assign(name = lambda x: x['name'].str.replace('[　 ]', '', regex=True))
                            )

data_wide_icu_pre = (combined_data_icu
             .groupby(['name', 'request'])['date']
             .apply(lambda x: ' ,'.join(x))
             .pivot(index='name', columns='request', values='date')  
             .reset_index()  
    )
columns_ok = data_wide_icu_pre.filter(regex=ok_sign).columns
if not columns_ok.empty:
    # 該当する列のデータをコンマ区切りで結合
    data_wide_icu_pre['accept'] = data_wide_icu_pre[columns_ok].apply(lambda x: ', '.join(x.dropna().astype(str)), axis=1)
    
columns_bad = data_wide_icu_pre.filter(regex=ng_sign).columns
if not columns_bad.empty:
    # 該当する列のデータをコンマ区切りで結合
    data_wide_icu_pre['reject'] = data_wide_icu_pre[columns_bad].apply(lambda x: ', '.join(x.dropna().astype(str)), axis=1)

data_wide_icu = (data_wide_icu_pre
                 .drop(columns=columns_ok)
                 .drop(columns=columns_bad)
                 .assign(species = 'icu')
    )

comment_icu = comment_dat.filter(regex=r'name|ICU勤務')

data_wide_icu_comment = (data_wide_icu_pre
                 .merge(comment_icu, on='name', how='left')
                 .assign(comment = lambda df: df['ICU勤務'].combine_first(df['ICU勤務希望についての備考']))
            )
#%%
# 1次救急のデータを整形
data_ichijikyu = (dat
        .filter(regex=r'1次救急.*\d{1,2}月')
        .rename(columns=extract_date)
)
data_concat_ichijikyu = (pd.concat([base_data, data_ichijikyu], axis=1)
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
# notes dataと結合

notes_data_ichijikyu = (dat_notes
                .filter(regex=r'人|日付|1次救急.*')
                .rename(columns=lambda x: 'name' if '人' in x else x)
                .rename(columns=lambda x: 'date' if '日付' in x else x)
                .rename(columns=lambda x: 'request' if '1次救急' in x and '備考' not in x else x)
)
combined_data_ichijikyu = (pd.concat([data_concat_ichijikyu, notes_data_ichijikyu], axis=0)
                            .applymap(lambda x: x.strip() if isinstance(x, str) else x)   
                            .assign(request=lambda x: np.where(x['request'] == '', None, x['request']))   
                            .assign(name = lambda x: x['name'].str.replace('[　 ]', '', regex=True))
                            )

data_wide_ichijikyu_pre = (combined_data_ichijikyu
             .groupby(['name', 'request'])['date']
             .apply(lambda x: ' ,'.join(x))
             .pivot(index='name', columns='request', values='date')
             .reset_index()
    )
columns_ok = data_wide_ichijikyu_pre.filter(regex=ok_sign).columns
if not columns_ok.empty:
    # 該当する列のデータをコンマ区切りで結合
    data_wide_ichijikyu_pre['accept'] = data_wide_ichijikyu_pre[columns_ok].apply(lambda x: ', '.join(x.dropna().astype(str)), axis=1)
    
columns_bad = data_wide_ichijikyu_pre.filter(regex=ng_sign).columns
if not columns_bad.empty:
    # 該当する列のデータをコンマ区切りで結合
    data_wide_ichijikyu_pre['reject'] = data_wide_ichijikyu_pre[columns_bad].apply(lambda x: ', '.join(x.dropna().astype(str)), axis=1)

data_wide_ichijikyu = (data_wide_ichijikyu_pre
                 .drop(columns=columns_ok)
                 .drop(columns=columns_bad)
                 .assign(species = 'ichijikyu')
    )

comment_ichijikyu = comment_dat.filter(regex=r'name|1次救急')

data_wide_ichijikyu_comment = (data_wide_ichijikyu_pre
                 .merge(comment_ichijikyu, on='name', how='left')
                 .assign(comment = lambda df: df['1次救急希望'].combine_first(df['1次救急希望についての備考']))
            )

# %%

pd.concat(list(wide_dict.values())).to_excel(f"prepdata/{month}m/tochoku_all_{month}.xlsx")

# %%
with open(f"prepdata/{month}m/long_dict_{month}.pkl", 'wb') as f:
    pickle.dump(long_dict, f, protocol=pickle.HIGHEST_PROTOCOL)
# %%
