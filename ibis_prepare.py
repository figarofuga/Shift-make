#%%
import pandas as pd
import ibis
from ibis import _
import ibis.selectors as s
ibis.options.interactive = True

ibis.set_backend("pandas")
# %%
# read excel data
dat = pd.read_excel("rawdata/2024_05answer.xlsx")
# %%
t = ibis.memtable(dat, name="t")
# %%
