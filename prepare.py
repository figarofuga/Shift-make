#%%
import numpy as np
import pandas as pd
import polars as pl
import seaborn as sns
from datetime import date
import datetime
import re

#%%
# read excel data
dat = pl.read_excel("rawdata/2024_05answer.xlsx")
#%%

name_column = 'お名前'
start_column = 'タイムスタンプ'
tochoku_stop_column = '日直・当直希望についての備考'
icijikyu_stop_column = '1次救急希望についての備考'
all_columns = dat.columns

name_index = all_columns.index(name_column)
start_index = all_columns.index(start_column)
tochoku_stop_index = all_columns.index(tochoku_stop_column)
ichijikyu_stop_index = all_columns.index(icijikyu_stop_column)

# 列名を変更する関数を定義
def extract_date(col_name):
    match = re.search(r'\[(\d{1,2}/\d{1,2}\([日月火水木金土祝]\))\]', col_name)
    return match.group(1) if match else col_name

# base data
base_data = (dat
             .select(all_columns[start_index:(name_index + 1)])
             )

# subset of toschoku data
tochoku_data = (dat
             .select(all_columns[(name_index+2):tochoku_stop_index])
             .rename(lambda col_name: extract_date(col_name))
             )

ichijikyu_data = (dat
                  .select(all_columns[(tochoku_stop_index+1):ichijikyu_stop_index])
                  .rename(lambda col_name: extract_date(col_name))
)
# subset of ichijikyu data

tochoku_data_prep = (pl.concat([base_data, tochoku_data], how="horizontal")
          .with_columns([pl.col("タイムスタンプ").str.strptime(pl.Datetime, "%m/%d/%Y %H:%M:%S")])
          .sort("タイムスタンプ")
          .group_by("お名前", maintain_order=True).last()
          .drop(['タイムスタンプ', 'メールアドレス'])
          .rename({'お名前': 'name', '現在在籍・研修中の科でお願いします': 'department'})
          .melt(id_vars=['name', 'department'], 
                variable_name='date', 
                value_name='request')
          .drop('department')
)

ichijikyu_data_prep = (pl.concat([base_data, ichijikyu_data], how="horizontal")
          .with_columns([pl.col("タイムスタンプ").str.strptime(pl.Datetime, "%m/%d/%Y %H:%M:%S")])
          .sort("タイムスタンプ")
          .group_by("お名前", maintain_order=True).last()
          .drop(['タイムスタンプ', 'メールアドレス'])
          .rename({'お名前': 'name', '現在在籍・研修中の科でお願いします': 'department'})
          .melt(id_vars=['name', 'department'], 
                variable_name='date', 
                value_name='request')
          .drop('department')
)
 
# %%

dat_notes = pl.read_excel("rawdata/notes_data.xlsx")

# %%

notes_tochoku_data_prep = (dat_notes
                           .select(['人', '\u3000日付', '日直・当直'])
                           .rename({"人": "name", 
                                    "\u3000日付": "date",
                                    "日直・当直": "request"})
                           .with_columns([
                               pl.when(pl.col("request").is_in(["　", " "]))
                               .then(None)
                               .when(pl.col("request") == '\u3000○１')
                               .then(pl.lit("第一希望日"))
                               .when(pl.col("request") == '\u3000○２')
                               .then(pl.lit("第二希望日"))
                               .when(pl.col("request") == '\u3000○３')
                               .then(pl.lit("第三希望日"))
                               .when(pl.col("request") == '\u3000○４')
                               .then(pl.lit("第四希望日"))
                               .when(pl.col("request") == '\u3000○')
                               .then(pl.lit("希望日"))
                               .when(pl.col("request") == '\u3000')
                               .then(None).otherwise(pl.lit("不可日")).alias("request")
                               ])
)                   
                                    
                                    
                                    

tochoku_data = pl.concat([tochoku_data_prep, notes_tochoku_data_prep], 
                         how="vertical")
# %%

notes_ichijikyu_data_prep = (dat_notes
                           .select(['人', '\u3000日付', '一次救急'])
                           .rename({"人": "name", 
                                    "\u3000日付": "date",
                                    "一次救急": "request"})
                           .with_columns([
                               pl.when(pl.col("request").is_in(["　", " "]))
                               .then(None)
                               .when(pl.col("request") == '\u3000○')
                               .then(pl.lit("希望日"))
                               .otherwise(pl.lit("不可日")).alias("request")
                               ]    
                                    ))



# notes_ichijikyu_data_prep['request'].unique().to_list()

ichijikyu_data = pl.concat([ichijikyu_data_prep, notes_ichijikyu_data_prep], 
                         how="vertical")


# %%
notes_ichijikyu_data_prep
dat_transformed = (notes_ichijikyu_data_prep.filter(pl.col("request") != "")
                    .group_by(["name", "request"])
                    .agg(pl.col("date").list().alias("dates"))
                    .pivot(pivot_column="request", values_column="dates")
                    .fill_none("")
                    .with_columns([pl.col("◯").apply(lambda x: ','.join(x) if x else ""),
                                   pl.col("x").apply(lambda x: ','.join(x) if x else ""),
                                   pl.col("X").apply(lambda x: ','.join(x) if x else ""),
                                   pl.col("◯1").apply(lambda x: ','.join(x) if x else "")])
                 )
# %%
# グループ化とカスタム集約を定義
def aggregate_dates(series: pl.Series) -> str:
    return ",".join(series.to_list())

tmp = (notes_tochoku_data_prep 
    .groupby("name") 
    .agg(pl.col("date").apply(aggregate_dates).alias("不可日")) 
)

# %%
