"""
Скрипт для генерации из шаблона, приказов о отчислении
"""
from docxtpl import DocxTemplate
import csv


# Загружаем данные
reader = csv.DictReader(open('Данные для приказа о отчислении.csv'), delimiter=';')

# Список для номеров приказов
lst_order = []
# Перебираем превращенный в список словарей документ csv и
for row in list(reader):
    lst_order.append(row['number_order'])

# Получаем уникальные значения
unique_order = list(set(lst_order))
# Создаем словарь, где номер приказа будет являтся ключом, а значением пустой список, куда будут добавляется строки с
# совпадающим номером приказа
data_finish = {number_order: [] for number_order in unique_order}
data_not_finish = {number_order: [] for number_order in unique_order}
data_all = {number_order: [] for number_order in unique_order}

# Заполняем словарь ФИО студентов
# Заполняем словарь со всеми данными
# Пересоздаем объект DictReader, потому что при первом использовании он исчерпался
reader = csv.DictReader(open('Данные для приказа о отчислении.csv'), delimiter=';')
for row in list(reader):
    if row['finish'].lower() == 'закончил':
        data_finish[row['number_order']].append(row['FIO'])
    else:
        data_not_finish[row['number_order']].append(row['FIO'])
    data_all[row['number_order']].append(row)

# Объединяем словари
data = {key: value[0] for key, value in data_all.items()}
# Добавляем для каждого ключа соответсвущий список студентов
for key, value in data.items():
    value['students'] = data_finish[key]
    value['not_finish'] = data_not_finish[key]

for number_order, value in data.items():
    # Создаем документ
    doc = DocxTemplate('template/Шаблон Приказ о отчислении.docx')
    # Заполняем словарь контекста
    context = {'date_order': value['date_order'], 'number_order': value['number_order'],
               'reason': value['reason'],
               'all_par': value['all_par'],
               'director':value['director'],'students':value['students'],'not_finish':value['not_finish'],
               'fio_do':value['fio_do'],'phone_do':value['phone_do'],'break_losers':value['break_losers']}
    # Генерируем документ
    doc.render(context)
    # Сохраняем документ
    doc.save(f'Приказ о отчислении {value["number_order"]}.docx')
