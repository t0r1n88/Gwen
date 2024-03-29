import pandas as pd
import openpyxl
import os


def processing_column_other(text: str):
    """
    Функция для подсчета причин указанных в колонке Прочее
    :param text: Текст разделенный запятыми
    :return: Текст разделенный знаками переноса вида:
     семейные обстоятельства -3
     переезд -4
    """
    if text == 'Нет':
        return ''
    else:
        # Создаем словарь для подсчета
        word_frequency = dict()
        # Очищаем текст от возможных знаков переноса

        text = text.replace('', '')
        text = text.replace("\n", " ")
        # Создаем список
        word_lst = text.split(',')

        # Итерируемся по списку слов и подсчитываем количество
        for word in word_lst:
            # о костыль.Но странно почему replace не срабатывает
            if word == '':
                continue
            if word in word_frequency:
                # Создаем ключ с заглавной буквы, чтобы потом не тратить на это время
                word_frequency[word] += 1
            else:
                word_frequency[word] = 1
        # Превращаем словарь в список, чтобы метод join правильно отработал
        output_lst = []
        for reason in word_frequency.items():
            output_lst.append(f'{reason[0]} - {reason[1]}')
        return ',\n'.join(output_lst)


def processing_type_df(df):
    """
    Функция для обработки датафрейма. Обработка nan, конвертирование типов колонок
    :param df: датафрейм
    :return: обработанный датафрейм
    """
    # Заменяем None в колонках для того чтобы избежать в последующем проблем с проверкой данных

    df.fillna(value={'Трудоустроен': 0, 'ИП': 0, 'Самозанятые': 0, 'Призваны в ВС': 0, 'Продолжают обучение': 0,
                     'Находятся в отпуске по уходу за ребенком': 0,
                     'Находящиеся под риском нетрудоустройства': 0, 'Состоят на учете в центрах занятости': 0,
                     'Не определились': 0, 'План.Трудоустроен': 0
        , 'План.ИП': 0, 'План.Самозанятые': 0, 'План.Призваны в ВС': 0, 'План.Продолжают обучение': 0,
                     'План.Находятся в отпуске по уходу за ребенком': 0,
                     'План.Находящиеся под риском нетрудоустройства': 0,
                     'План.Состоят на учете в центрах занятости': 0, 'План.Не определились': 0, 'Всего': 0,
                     'Лица с ограниченными возможностями здоровья': 0, 'Инвалиды и дети-инвалиды': 0,
                     'Инвалиды и дети-инвалиды (кроме учтенных в предыдущем столбце)': 0,
                     'Имеют договор о целевом обучении': 0,
                     'Лица с ограниченными возможностями здоровья (имеющие договор о целевом обучении)': 0,
                     'Инвалиды и дети-инвалиды (имеющие договор о целевом обучении)': 0,
                     'Инвалиды и дети-инвалиды (кроме учтенных в предыдущем столбце) (имеющие договор о целевом обучении)': 0},
              inplace=True)
    df.fillna(value={'Прочее': 'Нет,', 'План.Прочее': 'Нет,'}, inplace=True)
    # Применяем ко всему датафрейму конвертирование в инт, текстовые типы будут пропущены

    df[['Трудоустроен', 'ИП', 'Самозанятые', 'Призваны в ВС', 'Продолжают обучение',
        'Находятся в отпуске по уходу за ребенком',
        'Находящиеся под риском нетрудоустройства', 'Состоят на учете в центрах занятости', 'Не определились',
        'План.Трудоустроен'
        , 'План.ИП', 'План.Самозанятые', 'План.Призваны в ВС', 'План.Продолжают обучение',
        'План.Находятся в отпуске по уходу за ребенком', 'План.Находящиеся под риском нетрудоустройства',
        'План.Состоят на учете в центрах занятости', 'План.Не определились', 'Всего',
        'Лица с ограниченными возможностями здоровья', 'Инвалиды и дети-инвалиды',
        'Инвалиды и дети-инвалиды (кроме учтенных в предыдущем столбце)',
        'Имеют договор о целевом обучении',
        'Лица с ограниченными возможностями здоровья (имеющие договор о целевом обучении)',
        'Инвалиды и дети-инвалиды (имеющие договор о целевом обучении)',
        'Инвалиды и дети-инвалиды (кроме учтенных в предыдущем столбце) (имеющие договор о целевом обучении)']] \
        = df[['Трудоустроен', 'ИП', 'Самозанятые', 'Призваны в ВС', 'Продолжают обучение',
              'Находятся в отпуске по уходу за ребенком',
              'Находящиеся под риском нетрудоустройства', 'Состоят на учете в центрах занятости', 'Не определились',
              'План.Трудоустроен'
        , 'План.ИП', 'План.Самозанятые', 'План.Призваны в ВС', 'План.Продолжают обучение',
              'План.Находятся в отпуске по уходу за ребенком', 'План.Находящиеся под риском нетрудоустройства',
              'План.Состоят на учете в центрах занятости', 'План.Не определились', 'Всего',
              'Лица с ограниченными возможностями здоровья', 'Инвалиды и дети-инвалиды',
              'Инвалиды и дети-инвалиды (кроме учтенных в предыдущем столбце)',
              'Имеют договор о целевом обучении',
              'Лица с ограниченными возможностями здоровья (имеющие договор о целевом обучении)',
              'Инвалиды и дети-инвалиды (имеющие договор о целевом обучении)',
              'Инвалиды и дети-инвалиды (кроме учтенных в предыдущем столбце) (имеющие договор о целевом обучении)']].astype(
        'int64')
    # Устанавливаем для колонки Статус тип поля текстовый
    df['Статус'] = df['Статус'].astype('str')
    df['Статус'] = ''
    return df


def check_quantity_row(df):
    """
    Функция для проверки количества строк в сгруппированном датафрейме. На случай если в списке группы окажутся другие
    специальности или кураторы, ну или просто преподаватель опечатается
    :param df: сгруппированный датафрейм
    :return: статус проверки
    """
    # В корректном файле должно быть 1 строка и 32 колонки
    if df.shape[0] != 1 or df.shape[1] != 32:
        return None
    else:
        return True


def check_columns_name(base_df, checked_df):
    """
    Функция для проверки корректности названий колонок
    :param base_df: эталонный датафрейм с колонками
    :param checked_df: проверяемый датафрейм
    :return: результат проверки
    """

    if (len(base_df.columns) == len(checked_df.columns)) and (all((base_df.columns == checked_df.columns))):
        return True
    else:
        return None


def check_correct_data(df, name_file):
    """
    Функция для проверки итогов по разделам:фактическое, планируемое, инвалиды
    :param df: датафрейм со списком студентов
    :param name_file: имя проверяемого файла
    :return: датафрейм с дополнительным столбцом содержащим в себе итог проверки.
    """
    # Создаем датафрейм для записи итогов проверки и сохранения в excel
    output_check_df = pd.DataFrame(columns=['ФИО', 'Факт', 'План'])
    # Заменяем в колонках Прочее и План.прочее текст на числа, для удобства подсчетов

    df['Прочее'] = df['Прочее'].apply(lambda x: 0 if x == 'Нет,' else 1)
    df['План.Прочее'] = df['План.Прочее'].apply(lambda x: 0 if x == 'Нет' else 1)
    # Итерируемся по датафрейму, дада это не очень хорошо но файлы маленькие
    for row in df.itertuples(index=True, name='Студент'):
        fact_status = 'Данные корректны'
        plan_status = 'Данные корректны'
        real_sum = row[5] + row[6] + row[7] + row[8] + row[9] + row[10] + row[11] + row[12] + row[13] + row[14]
        if real_sum != 1:
            fact_status = f'Фактические показатели {row[1]} введены некорректно, студент посчитан несколько раз или не посчитан.'

        plan_sum = row[15] + row[16] + row[17] + row[18] + row[19] + row[20] + row[21] + row[22] + row[23] + row[24]
        if plan_sum != 1:
            plan_status = f'Планируемые показатели {row[1]} введены некорректно, студент посчитан несколько раз или не посчитан.'

        output_check_df.loc[row[0], 'ФИО'] = row[1]
        output_check_df.loc[row[0], 'Факт'] = fact_status
        output_check_df.loc[row[0], 'План'] = plan_status

    output_check_df.to_excel(f'Итог проверки {name_file}.xlsx', index=False)




# Создаем базовый  датафрейм

base_df = pd.DataFrame(columns=['ФИО','Группа','Специальность','Куратор','Трудоустроен', 'ИП', 'Самозанятые', 'Призваны в ВС', 'Продолжают обучение',
        'Находятся в отпуске по уходу за ребенком',
        'Находящиеся под риском нетрудоустройства', 'Состоят на учете в центрах занятости','Прочее', 'Не определились',
        'План.Трудоустроен'
        , 'План.ИП', 'План.Самозанятые', 'План.Призваны в ВС', 'План.Продолжают обучение',
        'План.Находятся в отпуске по уходу за ребенком', 'План.Находящиеся под риском нетрудоустройства',
        'План.Состоят на учете в центрах занятости','План.Прочее', 'План.Не определились', 'Всего',
        'Лица с ограниченными возможностями здоровья', 'Инвалиды и дети-инвалиды',
        'Инвалиды и дети-инвалиды (кроме учтенных в предыдущем столбце)',
        'Имеют договор о целевом обучении',
        'Лица с ограниченными возможностями здоровья (имеющие договор о целевом обучении)',
        'Инвалиды и дети-инвалиды (имеющие договор о целевом обучении)',
        'Инвалиды и дети-инвалиды (кроме учтенных в предыдущем столбце) (имеющие договор о целевом обучении)','Статус'])
# print(len(base_df.columns))
# base_df = pd.read_excel('temp/columns.xlsx')

# Получаем список файлов
example_dir = 'C:/Users/1/PycharmProjects/Gwen/resources'
with os.scandir(example_dir) as files:
    lst_files = [file.name for file in files if file.is_file() and file.name.endswith('xlsx')]
# Перебираем файлы
for file in lst_files:

    df = pd.read_excel(f'resources/{file}')
    # Проверяем корректность названий колонок
    if not check_columns_name(base_df, df):
        print('Некорректное количество колонок или название колонок')
        continue
    # Обрабатываем колонки, приводим к типу инт, заполняем пропуски
    df = processing_type_df(df)

    # Проверяем на правильность заполнения, совпадают ли суммы
    checked_df = df.copy()

    check_correct_data(checked_df, file)
    # Удаляем столбец с ФИО, так как при группировке он будет мешать
    df = df.drop(['ФИО'], axis=1)

    # Проверяем корректность заполнения

    # добавляем параметр numeric_only чтобы текстовые значения также суммировались
    temp_df = df.groupby(['Специальность', 'Группа', 'Куратор'], ).sum(numeric_only=False).reset_index()
    if not check_quantity_row(temp_df):
        print('В файле присутствуют другие специальности,группы,кураторы или есть ошибки в написании')
        continue



    base_df = base_df.append(temp_df, ignore_index=True)

# Сохраняем промежуточный  результат для того чтобы легче было отслеживать корректность данных
base_df = base_df.drop('ФИО', axis=1)
base_df.to_excel('Промежуточный результат.xlsx', index=False)

# Проводим итоговую группировку
# Нужно провести подсчет причин в графе Прочие.
itog_df = base_df.groupby(['Специальность']).sum(numeric_only=False).reset_index()

# Обрабатываем колонку Прочее

itog_df['Прочее'] = itog_df['Прочее'].apply(processing_column_other)
itog_df['План.Прочее'] = itog_df['План.Прочее'].apply(processing_column_other)

itog_df.to_excel('Итоговый результат.xlsx', index=False)
