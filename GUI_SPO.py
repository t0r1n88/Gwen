# -*- coding: utf-8 -*-

import pandas as pd
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog,QLabel
from PyQt5.QtGui import QIcon, QPixmap
import traceback
import sys


class Ui_MainWindow(object):
    # Создаем атрибут класса. Датафрейм с эталонными колонками
    template_df = pd.DataFrame(
        columns=['ФИО', 'Группа', 'Специальность', 'Куратор', 'Трудоустроен', 'ИП', 'Самозанятые', 'Призваны в ВС',
                 'Продолжают обучение',
                 'Находятся в отпуске по уходу за ребенком',
                 'Находящиеся под риском нетрудоустройства', 'Состоят на учете в центрах занятости', 'Прочее',
                 'Не определились',
                 'План.Трудоустроен'
            , 'План.ИП', 'План.Самозанятые', 'План.Призваны в ВС', 'План.Продолжают обучение',
                 'План.Находятся в отпуске по уходу за ребенком', 'План.Находящиеся под риском нетрудоустройства',
                 'План.Состоят на учете в центрах занятости', 'План.Прочее', 'План.Не определились', 'Всего',
                 'Лица с ограниченными возможностями здоровья', 'Инвалиды и дети-инвалиды',
                 'Инвалиды и дети-инвалиды (кроме учтенных в предыдущем столбце)',
                 'Имеют договор о целевом обучении',
                 'Лица с ограниченными возможностями здоровья (имеющие договор о целевом обучении)',
                 'Инвалиды и дети-инвалиды (имеющие договор о целевом обучении)',
                 'Инвалиды и дети-инвалиды (кроме учтенных в предыдущем столбце) (имеющие договор о целевом обучении)',
                 'Статус'])

    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(542, 537)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        # Метка
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setEnabled(True)
        self.label.setGeometry(QtCore.QRect(170, 10, 231, 61))
        font = QtGui.QFont()
        font.setPointSize(20)
        self.label.setFont(font)
        self.label.setObjectName("label")





        # Создаем базовый датафрейм для экземпляра класса
        self.base_df = Ui_MainWindow.template_df.copy()

        # Создаем кнопку выбора файлов
        self.getFileNamesButton = QtWidgets.QPushButton(self.centralwidget)
        self.getFileNamesButton.setGeometry(QtCore.QRect(200, 120, 161, 51))
        self.getFileNamesButton.setObjectName("pushButton")
        # Привязываем метод выбора файлов
        self.getFileNamesButton.clicked.connect(self.getFileNames)

        # Создаем кнопку выбора конечной папки
        self.getDirectoryButton = QtWidgets.QPushButton(self.centralwidget)
        self.getDirectoryButton.setGeometry(QtCore.QRect(200, 200, 161, 51))
        self.getDirectoryButton.setObjectName("pushButton_2")
        # Привязываем метод выбора папки
        self.getDirectoryButton.clicked.connect(self.getDirectory)

        # Создаем кнопку Выполнения
        self.makeButton = QtWidgets.QPushButton(self.centralwidget)
        self.makeButton.setGeometry(QtCore.QRect(200, 400, 161, 51))
        self.makeButton.setObjectName("pushButton_3")
        # Привязываем метод выполнения
        self.makeButton.clicked.connect(self.run_script)
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 542, 21))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Кассандра"))
        self.label.setText(_translate("MainWindow", "Обработка данных"))
        self.getFileNamesButton.setText(_translate("MainWindow", "Выберите файлы"))
        self.getDirectoryButton.setText(_translate("MainWindow", "Выберите конечную папку"))
        self.makeButton.setText(_translate("MainWindow", "Выполнить обработку"))

    def getDirectory(self):
        """
        Метод для получения пути к папке
        :return: Пути к папке
        """
        self.output_dir = QFileDialog.getExistingDirectory(MainWindow, "Выбрать Папку")

    # Для вывода ошибок pyqt в консоль
    def excepthook(exc_type, exc_value, exc_tb):
        tb = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
        print("Oбнаружена ошибка !:", tb)

    sys.excepthook = excepthook

    def getFileNames(self):
        """
        Метод для получения путей к файлам
        :return: Пути к файлам
        """
        # В filenames хранится список файлов
        """
        Можно выбирать допустимые типы файлов. Пример
             "Выбрать файл",
         ".",
         "Text Files(*.txt);;JPEG Files(*.jpeg);;\
         PNG Files(*.png);;GIF File(*.gif);;All Files(*)")

        """
        self.filenames, ok = QFileDialog.getOpenFileNames(MainWindow,
                                                          "Выберите несколько файлов",
                                                          ".",
                                                          "All Files(*.xlsx)")

    def check_columns_name(self, base_df, checked_df):
        """
    Метод для проверки корректности названий колонок
    :param base_df: эталонный датафрейм с колонками
    :param checked_df: проверяемый датафрейм
    :return: результат проверки
        """
        if (len(base_df.columns) == len(checked_df.columns)) and (all((base_df.columns == checked_df.columns))):
            return True
        else:
            return None

    def processing_type_df(self, df):
        """
    Метод для обработки датафрейма. Обработка nan, конвертирование типов колонок
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

    def check_correct_data(self, df, name_file):
        """
        Метод для проверки итогов по разделам:фактическое, планируемое, инвалиды
        :param df: датафрейм со списком студентов
        :param name_file: имя проверяемого файла
        :return: сохраняет результаты проверки в файлы Excel формата Проверка_Имя файла
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

        output_check_df.to_excel(f'{self.output_dir}/Итог проверки_{name_file}', index=False)

    def check_quantity_row(self, df):
        """
        Метод для проверки количества строк в сгруппированном датафрейме. На случай если в списке группы окажутся другие
        специальности или кураторы, ну или просто преподаватель опечатается
        :param df: сгруппированный датафрейм
        :return: статус проверки
        :return:
        """
        # В корректном файле должно быть 1 строка и 32 колонки
        if df.shape[0] != 1 or df.shape[1] != 32:
            return None
        else:
            return True

    def processing_column_other(self, text):
        """
        Метод для подсчета причин указанных в колонке Прочее
        :param text: Текст разделенный запятыми
        :return: Текст разделенный знаками переноса вида:
         семейные обстоятельства -3
         переезд -4
        :return:
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

    def run_script(self):
        """
        Метод запускающий обработку
        """
        # Перебираем файлы
        for file in self.filenames:
            # Получаем имя файла для того чтобы сохранять называть файлы проверки
            file_name = file.split('/')[-1]
            # Создаем датафрейм
            df = pd.read_excel(file,engine='openpyxl')
            # Проверяем корректность названий колонок
            if not self.check_columns_name(self.base_df, df):
                print(f'{file_name} Некорректное количество колонок или название колонок')
                continue
            # Обрабатываем колонки, приводим к типу инт, заполняем пропуски
            df = self.processing_type_df(df)
            # Проверяем на правильность заполнения, совпадают ли суммы
            checked_df = df.copy()
            self.check_correct_data(checked_df, file_name)

            # Удаляем столбец с ФИО, так как при группировке он будет мешать
            df = df.drop(['ФИО'], axis=1)
            # добавляем параметр numeric_only чтобы текстовые значения также суммировались
            temp_df = df.groupby(['Специальность', 'Группа', 'Куратор'], ).sum(numeric_only=False).reset_index()

            # Проверяем правильность группировки, есть ли в файле другие специальности,кураторы, группы
            if not self.check_quantity_row(temp_df):
                print(f'{file_name} В файле присутствуют другие специальности,группы,кураторы или есть ошибки в написании этих данных')
                continue
            self.base_df = self.base_df.append(temp_df, ignore_index=True)
        self.base_df = self.base_df.drop('ФИО', axis=1)
        self.base_df.to_excel(f'{self.output_dir}/Промежуточный результат.xlsx', index=False)

        # Проводим итоговую группировку
        # Нужно провести подсчет причин в графе Прочие.
        itog_df = self.base_df.groupby(['Специальность']).sum(numeric_only=False).reset_index()

        # Обрабатываем колонку Прочее

        itog_df['Прочее'] = itog_df['Прочее'].apply(self.processing_column_other)
        itog_df['План.Прочее'] = itog_df['План.Прочее'].apply(self.processing_column_other)

        # Сохраняем результат
        itog_df.to_excel(f'{self.output_dir}/Итоговый результат.xlsx', index=False)

    # для вывода ошибок в консоль
    sys.excepthook = excepthook


if __name__ == "__main__":
    # Создаем объект приложения(экземпляр класса)
    # Обязательная строка
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()

    sys.exit(app.exec_())
