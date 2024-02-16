#%%
import numpy as np
import pandas as pd
import polars as pl
import seaborn as sns
import pulp
from datetime import date
import jpholiday
import datetime
import re

#%%
# read csv data
dat=pl.read_excel("2024_03data.xlsx")

#%%

start_column='タイムスタンプ'
stop_column='日直・当直希望についての備考'
all_columns=dat.columns

start_index=all_columns.index(start_column)
stop_index=all_columns.index(stop_column)


dat2 = (dat
        .select(all_columns[(start_index+4):stop_index])
        .rename(lambda c: re.search(r"\d{1,2}(.*)\)", c).group())
)

dat3 = (pl.concat([dat.select(all_columns[start_index:start_index+4]), 
                 dat2], how="horizontal")
          .with_columns([pl.col("タイムスタンプ").str.strptime(pl.Datetime, "%m/%d/%Y %H:%M:%S")])
          .sort("タイムスタンプ")
          .group_by("お名前", maintain_order=True).last()

)


# `holiday` column will be of type Boolean: True if the date is a holiday, False otherwise
# %%
