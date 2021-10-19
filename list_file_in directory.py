import os
import pandas as pd
import openpyxl
path = 'c:/Users/1\YandexDisk/БРИТ/Тандем/Развертывание/Импорт/для ТАНДЕМ/'

lst_name = []
for file in os.listdir(path):
    name_file = file.split('.')[0]
    lst_name.append(name_file)

df = pd.DataFrame(lst_name,index=range(len(lst_name)),columns=['Название группы'])
df.to_excel('Список групп.xlsx',index=False)