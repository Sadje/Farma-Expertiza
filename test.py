import pandas as pd


path = r'Raw_2025\\'

df = pd.read_excel(path + 'Anketa-KPM-0244.xlsx', sheet_name='Р1_НИОКТР_2', header=5, skipfooter=1)
df = df.drop(df.columns[[1, 2, 4]], axis=1)
df = df.rename(columns={df.columns[0]: "Результаты"})
df = df.rename(columns={df.columns[1]: "Булевые"})
df.loc[df['Наличие'] == True, 'Наличие'] = 'Да'
df.loc[df['Наличие'] == False, 'Наличие'] = 'Нет'
print(df)
