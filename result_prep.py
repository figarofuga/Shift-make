#%%
import pandas as pd
import pickle

#%%
# read excel data
res_dat_5 = pd.read_excel("rawdata/2024_5_res.xlsx")
res_dat_6 = pd.read_excel("rawdata/2024_6_res.xlsx")

# read pickle of desired data
with open('long_dict_6.pkl', 'rb') as handle:
    loaded_data = pickle.load(handle)
    

#%%
prep_dat_5 = (res_dat_5
            .melt(id_vars = "2024年5月当直表", value_name='name', var_name='tochoku')
            .dropna()
            .query("tochoku.isin(['D日直', 'E当直', 'F当直', 'A当直', 'B当直', '教育当直', '病棟当直'])")
            .groupby(["name", "tochoku"])["2024年5月当直表"]
            .apply(lambda x: ' ,'.join(x))
)

prep_dat_6 = (res_dat_6
            .melt(id_vars = "2024年6月当直表", value_name='name', var_name='tochoku')
            .dropna()
            .query("tochoku.isin(['D日直', 'E当直', 'F当直', 'A当直', 'B当直', '教育当直', '病棟当直'])")
            .groupby(["name", "tochoku"])["2024年6月当直表"]
            .apply(lambda x: ' ,'.join(x))
)
# %%
