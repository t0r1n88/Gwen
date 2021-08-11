import pandas as pd
import openpyxl

# Основной список
df_main = pd.read_excel('resources/selection.xlsx')
# Список доп статусы
df_slave = pd.read_excel('resources/dop_status.xlsx',usecols=['Абитуриент','Доп. статус'])

all_df = pd.merge(df_main,df_slave,left_on='ФИО',right_on='Абитуриент')

out_df = all_df[['ФИО','Доп. статус']]

out_df.to_excel('Orphans_report.xlsx',index=False)
