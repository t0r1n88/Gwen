"""
Скрипт для генерации из шаблона, приказов о зачислении
"""
from docxtpl import DocxTemplate
import csv

# Загружаем данные
reader = csv.DictReader(open('data.csv'), delimiter=';')

# Создаем словарь, где номер приказа будет являтся ключом, а остальные данные списком словарей.
# Для того чтобы записывать несколько студентов  в  один приказ

# Список для номеров приказов
lst_order = []
# Перебираем превращенный в список словарей документ csv и
for row in list(reader):
    lst_order.append(row['number_order'])

# Получаем уникальные значения
unique_order = list(set(lst_order))
# Создаем словарь, где номер приказа будет являтся ключом, а значением пустой список, куда будут добавляется строки с
# совпадающим номером приказа
data_fio = {number_order: [] for number_order in unique_order}
data_all = {number_order: [] for number_order in unique_order}

# Заполняем словарь ФИО студентов
# Заполняем словарь со всеми данными
# Пересоздаем объект DictReader, потому что при первом использовании он исчерпался
reader = csv.DictReader(open('data.csv'), delimiter=';')
for row in list(reader):
    data_fio[row['number_order']].append(row['FIO'])
    data_all[row['number_order']].append(row)
"""
Получаем словари вида:
Номер приказа: список студентов с таким же номером приказа
Номер приказа: все строки с таким же номеров приказа
После чего мы можем спокойно генерировать документы не боясь что значения перепутаются
"""
# Объединяем словари
data = {key: value[0] for key, value in data_all.items()}
# Добавляем для каждого ключа соответсвущий список студентов
for key, value in data.items():
    value['students'] = data_fio[key]

for number_order, value in data.items():
    # Создаем документ
    doc = DocxTemplate('template/template.docx')
    # Заполняем словарь контекста
    context = {'date_order': value['date_order'], 'number_order': value['number_order'],
               'type_prog': value['type_prog'],
               'name_prog': value['name_prog'], 'volume': value['volume'], 'date_begin': value['date_begin'],
               'date_end': value['date_end'],
               'year': value['year'],'students':value['students'],'prog_prep':value['prog_prep'],'control':value['control'],
               'director':value['director']}
    # Генерируем документ
    doc.render(context)
    # Сохраняем документ
    doc.save(f'Приказ о зачислении {value["number_order"]}.docx')
