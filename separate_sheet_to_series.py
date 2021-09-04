import pandas as pd

index = pd.date_range(start='2021-8-9 01:00:00', end='2021-8-28 23:00:00', freq='1h0min')
a = 5
t_18_30 = pd.date_range(start='2021-8-9 18:30:00', end='2021-8-28 18:30:00', freq='1D')
t_19_30 = pd.date_range(start='2021-8-9 19:30:00', end='2021-8-28 19:30:00', freq='1D')
index = index.append(t_18_30).append(t_19_30)

df = pd.DataFrame(index=index).sort_index()
df.to_excel('D:\\My Drive\\NWPGCL_Khl225MW_gen_9to28_8_2021.xlsx')
b = 5


