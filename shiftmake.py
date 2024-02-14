#%%
import numpy as np
import pandas as pd
import polars as pl
import seaborn as sns
import pulp
from datetime import date
import jpholiday
import datetime

#%%
# read csv data
dat=pl.read_excel("2024_03data.xlsx")

#%%

start_column='タイムスタンプ'
stop_column='日直・当直希望についての備考'
all_columns=dat.columns
start_index=all_columns.index(start_column)
stop_index=all_columns.index(stop_column)

dat2=dat.select(all_columns[start_index:stop_index])

# make test data
# to make columns of A to Z that is random choice of "no", "soso", "good"

# dat2= pl.DataFrame({
#     "date": pl.date_range(
#     date(2024, 5, 1), 
#     date(2024, 5, 31), 
#     "1d", 
#     eager=True),
#     **{chr(65+i): np.random.choice(["no", "soso", "good"], replace=True, p=[0.5, 0.3, 0.2], size=31) for i in range(26)}
# })

# convert "date" column to date type column 
# #  def is_holiday(date):
# #     return jpholiday.is_holiday(date)

# # Define a function that checks if a date is a holiday

# dat2=(dat.with_columns(
#     [pl.col("date")
#         .str.strptime(pl.Date,
#                       strict=False)
# # make date column to date type to concordant with jpholiday package
#         .apply(lambda x: datetime.date(x.year, x.month, x.day))
#         .apply(is_holiday).alias("holiday")])
                  
# )



# `holiday` column will be of type Boolean: True if the date is a holiday, False otherwise
# %%
