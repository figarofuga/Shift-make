#%%
import pandas as pd

#%%
# read excel data
res_dat = pd.read_excel("rawdata/2024_5_res.xlsx")

#%%
prep_dat = (res_dat
            .melt(id_vars = "2024年5月当直表", value_name='name', var_name='tochoku')
            .dropna()
            .query("tochoku.isin(['D日直', 'E当直', 'F当直', 'A当直', 'B当直', '教育当直', '病棟当直'])")
            .groupby(["name", "tochoku"])["2024年5月当直表"]
            .apply(lambda x: ' ,'.join(x))
)
# %%
