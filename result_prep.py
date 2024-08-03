#%%
import numpy as np
import pandas as pd
import pickle
import re
#%%
month = 8
#%%
named_list = {
 '久冨木原': '久冨木原健二', 
 '伊藤(国)': '伊藤国秋', 
 '佐久川': '佐久川桂怜', 
 '佐久間': '佐久間一也', 
 '佐川': '佐川偲', 
 '先崎': '先崎光', 
 '勝俣': '勝俣敬寛', 
 '吉村': '吉村梨沙', 
 '吉永': '吉永薫',
 '吉田(博)': '吉田博道', 
 '吉田(心)': '吉田心慈', 
 '大重': '大重達寛', 
 '太良': '太良史郎', 
 '小宮': '小宮健太郎', 
 '小林(洋)': '小林洋太', 
 '星': '星貴文', 
 '山田': '山田康博', 
 '岡村': '岡村真伊', 
 '岩崎': '岩崎文美',
 '川瀬': '川瀬咲', 
 '平井': '平井智大', 
 '持丸': '持丸貴生', 
 '新美': '新美望', 
 '松井': '松井貴裕', 
 '東條': '東條誠也',
 '松原': '松原龍輔', 
 '松永': '松永崇宏', 
 '林(智)': '林智史', 
 '津山': '津山頌章', 
 '福原': '福原誠一郎',
 '深沢': '深沢夏海', 
 '熊木': '熊木聡美',
 '石井': '石井真央', 
 '福井': '福井梓穂', 
 '篠崎': '篠﨑太郎', 
 '細川': '細川善弘', 
 '織部': '織部峻太郎', 
 '脇坂': '脇坂悠介', 
 '臼坂': '臼坂優希', 
 '茅島': '茅島敦人', 
 '荒金': '荒金直美', 
 '莇': '莇舜平',
 '藤澤': '藤澤まり', 
 '谷岡': '谷岡友則', 
 '野上': '野上創生', 
 '鈴木(徹)': '鈴木徹志朗', 
 '門松': '腎内門松', 
 '雪野': '雪野満', 
 '青木': '青木康浩', 
 '高木': '髙木菜々美', 
 '高梨': '髙梨航輔', 
 '黒崎': '黒崎颯' 
 
 }


#%%
# read excel data
res_dat_before = pd.read_excel(f"rawdata/{month-1}m/2024_{month-1}_res.xlsx")
res_dat_check = pd.read_excel(f"rawdata/{month}m/2024_{month}_res.xlsx")

# read pickle of desired data
with open(f"prepdata/{month}m/long_dict_{month}.pkl", 'rb') as handle:
    request_data_check = pickle.load(handle)
    

#%%
prep_dat_before = (res_dat_before
            .assign(ICU勤務 = lambda x: x['ICU勤務'].str.replace(r'^(AM|PM): ', '', regex=True))
            .melt(id_vars = f'2024年{month-1}月当直表', value_name='name', var_name='tochoku')
            .dropna()
            .query("tochoku.isin(['D日直', 'E当直', 'F当直', 'A当直', 'B当直', '教育当直', '病棟当直'])")
            .groupby(["name", "tochoku"])[ f'2024年{month-1}月当直表']
            .apply(lambda x: ' ,'.join(x))
            .reset_index()
            .assign(name = lambda x: x['name'].str.replace('[　 ]', '', regex=True))
            .assign(tochoku_date = lambda x: x[ f'2024年{month-1}月当直表'].str.extract(r'(\d{1,2}月\d{1,2}日)'))
)


prep_dat_check = (res_dat_check
                     .assign(ICU勤務 = lambda x: x['ICU勤務'].str.replace(r'^(AM|PM): ', '', regex=True))
                     .melt(id_vars = f'2024年{month}月当直表', value_name='name', var_name='tochoku')
                     .dropna()
                     .query("tochoku.isin(['D日直', 'E当直', 'F当直', '内科1(旧教育/B)', '内科2(旧病棟/A)', 'ICU勤務'])")
                     .groupby(["name", "tochoku"])[f'2024年{month}月当直表']
                     .apply(lambda x: ' ,'.join(x))
                     .reset_index()
                     .assign(name = lambda x: x['name'].str.replace('[　 ]', '', regex=True))
                     .assign(tochoku_date = lambda x: x[f'2024年{month}月当直表'].str.extract(r'(\d{1,2}月\d{1,2}日)'))
)


# %%
tochoku_request = (request_data_check['tochoku']
                   .assign(request2 = lambda x: np.select(
    [
        x['request'].isin(['○１', '○', '◯', '希望日']),
        x['request'].isin(['×', '✕', '不可日'])
    ], 
    [
        1, 
        -1, 

    ],
    default = 0  # 上記のいずれにも該当しない場合
))

)
tmp = (
    prep_dat_check
    .assign(name = lambda x:x['name'].replace(named_list))
    .assign(tochoku_date2 = lambda x: pd.to_datetime('2024年' + x['tochoku_date'], format='%Y年%m月%d日'))
    .assign(shift_day = lambda dat: dat.sort_values(['name', 'tochoku_date2']).groupby('name')['tochoku_date2'].shift(-1))        
    .assign(diffday = lambda x: (x['shift_day'] - x['tochoku_date2']).dt.days)    
)

# %%
check_request_dt = pd.merge(tochoku_request, tmp, on='name', how='right')
