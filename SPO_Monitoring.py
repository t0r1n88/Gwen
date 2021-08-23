import pandas as pd
import openpyxl
import os

# Создаем базовый  датафрейм
base_df = pd.read_excel('temp/columns.xlsx')

# Получаем список файлов
example_dir = 'C:/Users/1/PycharmProjects/Gwen/resources'
with os.scandir(example_dir) as files:
    lst_files = [file.name for file in files if file.is_file() and file.name.endswith('xlsx')]
# Перебираем файлы
for file in lst_files:
    df = pd.read_excel(f'resources/{file}')
    # Удаляем столбец с ФИО, так как при группировке он будет мешать
    df = df.drop(['ФИО'],axis=1)

    # добавляем параметр numeric_only чтобы текстовые значения также суммировались
    temp_df = df.groupby(['Специальность','Группа'],).sum(numeric_only=False).reset_index()
    # Добавляем данные в итоговый датафрейм
    base_df = base_df.append(temp_df,ignore_index=True)

    # temp_df.to_excel(f'{file}',index=False)
# Сохраняем промежуточный  результат для того чтобы легче было отслеживать корректность данных
base_df.to_excel('Промежуточный результат.xlsx',index=False)

# Проводим итоговую группировку
# Нужно провести подсчет причин в графе Прочие.
itog_df = base_df.groupby(['Специальность']).sum(numeric_only=False).reset_index()
itog_df.to_excel('Итоговый результат.xlsx',index=False)
