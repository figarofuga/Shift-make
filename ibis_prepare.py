#%%
import ibis

con = ibis.pandas.connect()

# %%
# read excel data
dat = con.read_excel("rawdata/2024_05answer.xlsx")
# %%
