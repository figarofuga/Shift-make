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
dat=pl.read_csv("testshiftdata.csv")

#%%
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

dat2=(dat.with_columns(
    [pl.col("date")
        .str.strptime(pl.Date,
                      strict=False)])
)

# %%
