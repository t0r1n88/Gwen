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
        if word == '' or word == 'Отсутствуют':
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


def check_data(df, quantity):
    """
    Функция для проверки полученного датафрейма на корректность данных.
    Проверка фактических и планируемых показателей, проверка количества целевиков и инвалидов.
    :param df: однострочный датафрейм полученный после суммирования
    :quantity: количество студентов
    :return: строка со статусом проверки
    """
    # Создаем счетчики для подсчета причин из колонок Прочее и План.Прочее
    real_count_reason = 0
    plan_count_reason = 0
    # Сохраняем данные из колонок Прочее и План.Прочее
    print( df['Прочее'].values)
    # real_text_reasons = df['Прочее'].values

    # # считаем количество реальных причин
    # for real_reason in real_text_reasons.split(','):
    #     # Условие, чтобы не учитывать пустую строку которая образуется
    #     if not real_reason == '':
    #         real_count_reason += 1


    # plan_text_reasons = df['План.Прочее'].values
    # # считаем количество планируемых причин
    # for plan_reason in plan_text_reasons.split(','):
    #     if not plan_reason == '':
    #         plan_count_reason += 1

    # print(real_count_reason)
    # print(plan_count_reason)



    # Удаляем колонки с причинами
    df.drop(['Прочее', 'План.Прочее'], axis=1, inplace=True)
    # Получаем строку с итогами по фактическому распределению выпускников

    d = df.iloc[0, 3:12].sum()


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
    df = df.drop(['ФИО'], axis=1)

    # добавляем параметр numeric_only чтобы текстовые значения также суммировались
    temp_df = df.groupby(['Специальность', 'Группа', 'Куратор'], ).sum(numeric_only=False).reset_index()
    # TODO проверка количества строк. Должно быть не более 1, если больше то в файле присутствует еще одна группа,куратор,специальность

    # Копируем датафрейм, так как внутри функции он будет изменятся
    checked_df = temp_df.copy()
    # Получаем результат проверки
    verification_status = check_data(temp_df, df.shape[0])
    # Добавляем данные в итоговый датафрейм
    checked_df['Статус'] = verification_status

    base_df = base_df.append(checked_df, ignore_index=True)

# Сохраняем промежуточный  результат для того чтобы легче было отслеживать корректность данных
base_df.to_excel('Промежуточный результат.xlsx', index=False)

# Проводим итоговую группировку
# Нужно провести подсчет причин в графе Прочие.
itog_df = base_df.groupby(['Специальность']).sum(numeric_only=False).reset_index()

# Обрабатываем колонку Прочее
# Сначала приводим заменяем возможные 0.0 на строку Отсутствуют
itog_df['Прочее'] = itog_df['Прочее'].apply(lambda x: 'Отсутствуют' if not type(x) == str else x)
itog_df['Прочее'] = itog_df['Прочее'].apply(processing_column_other)
itog_df.to_excel('Итоговый результат.xlsx', index=False)
