#%%
import pandas as pd
import pickle
import re

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
            .reset_index()
            .assign(name = lambda x: x['name'].str.replace('[　 ]', '', regex=True))
)
prep_dat_6 = (res_dat_6
                     .melt(id_vars = "2024年6月当直表", value_name='name', var_name='tochoku')
                     .dropna()
                     .query("tochoku.isin(['D日直', 'E当直', 'F当直', 'A当直', 'B当直', '教育当直', '病棟当直'])")
                     .groupby(["name", "tochoku"])["2024年6月当直表"]
                     .apply(lambda x: ' ,'.join(x))
                     .reset_index()
                     .assign(name = lambda x: x['name'].str.replace('[　 ]', '', regex=True))
)
# %%

{'久冨木原': '久冨木原健二', 
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
 '津山', 
 '深沢', 
 '熊木',
 '石井', '福井', '福原', '篠崎', '細川', '織部', '脇坂', '臼坂', '茅島', '荒金', '莇',
       '藤澤', '谷岡', '野上', '鈴木(徹)', '門松', '雪野', '青木', '高木', '高梨', '黒崎'}

 , , ,
       '津山\u3000頌章', '消化器内科\u3000\u3000福原\u3000誠一郎',
       '消化器内科\u3000渡邉\u3000多代', '清水\u3000隆之', '熊木聡美', '片山充哉',
       '石井\u3000真央', '福井梓穂', '福島龍貴', '篠﨑太郎', '籠尾壽哉', '糖尿病科岩瀬', '細川善弘',
       '織部\u3000峻太郎', '腎内\u3000門松', '腎臓内科\u3000松浦', '臼坂優希', '荒金直美', '莇舜平',
       '藤村 慶子', '藤澤まり', '谷岡友則', '里見\u3000良輔', '野上\u3000創生', '鈴木勝也',
       '長谷川華子', '雪野\u3000満', '髙木菜々美']


# %%
