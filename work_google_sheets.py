import gspread

# Указываем путь к json key
gc = gspread.service_account(filename='c:/Users/1\PycharmProjects/Key google Sheets.json')
# Открываем тестовую таблицу
sh = gc.open('Тест распределения доступа')
sh.share('bator1qaz@gmail.com',perm_type='user',role='writer')

sh =gc.create('Новый лист для проверки')







# print(sh.sheet1.get('B1'))
#
# # Получение рабочего листа
# worksheet = sh.sheet1
# # печать строки
# print(worksheet.col_values(1))
#
# # Получение всех значений из рабочего листа в виде списка словарей
# print(worksheet.get_all_records())

# Создание нового листа

