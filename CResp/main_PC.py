""" Описание модуля
Стартовый модуль с прикручиванием UX к макету GUI, и запуском приложения

"""

""" Подключение модулей обработки """
from modules import CR_reader  as crr   # Считывание таблицы в базу разбора для конкретной группы
from modules import CR_parser  as crp   # Парсинг базы разбора в нормальную БД для каждой логической записи
from modules import CR_analyze as cra   # Анализатор БД, для подправки косяков и определения числа подгрупп
from modules import CR_writter as crw   # Форматная запись БД в таблицу EXCEL

""" Подключение элементов GUI """
from PyQt5.QtWidgets import QApplication, QFileDialog, QMainWindow, QWidget
from CR_GUI_PC       import Ui_MainWindow
from webbrowser      import open as open_link
from pyperclip       import copy as cp
from sys             import exit as close_app

""" Задание интерфейса """
class my_window(QMainWindow):
    """ Атрибуты для обработки расписания """
    path   = '' # Путь к обрабатываемой книге
    book   = '' # Открытая книга
    sheet  = '' # Выбор листа
    groups = '' # Строка с инфой о группах, периоде и т.п

    def __init__(self):
        super(my_window, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        # Реакция на выбор репозитория
        self.ui.action_1.triggered.connect(lambda: open_link('https://github.com/DK905/CoolResp'))

        # Реакция на вызов краткой справки
        self.ui.action_2.triggered.connect(lambda: self.ui.msgInf.exec())

        # Реакция на вызов технической инфы о проге
        self.ui.action_3.triggered.connect(lambda: self.ui.msgAbt.exec())

        # Реакция на вызов диалогового окна выбора загружаемого файла
        self.ui.pushButton_1.clicked.connect(self.OpenFile)

        # Реакция на подтверждение пути
        self.ui.pushButton_2.clicked.connect(self.ConfirmPath)

        # Реакция на смену листа
        self.ui.comboBox_1.currentIndexChanged.connect(self.SheetInfo)

        # Реакция на кнопку конвертации
        self.ui.pushButton_3.clicked.connect(self.Upgrade)

        # Реакция на кнопку сворачивания логов
        self.ui.toogleButton_1.clicked.connect(self.Turn)

        # Реакция на двойной клик по логу
        self.ui.listWidget.itemDoubleClicked.connect(self.CopyLog)

    # Обработка ошибки
    def ErDo(self, text):
        text = str(text)
        self.ui.listWidget.addItems([text, ''])
        self.Disintegration()
        self.ui.msgErr.setText(text)
        self.ui.msgErr.exec()

    # Диалоговое окно загрузки файла
    def OpenFile(self):
        path = QFileDialog.getOpenFileName(self,
                                                     'Выбрать файл', # Название диалогового окна
                                                     '.',            # Имя файла по умолчанию
                                                     'EXCEL таблицы(*.xls*);;Все файлы(*)') # Поддерживаемые типы файлов
        self.ui.textBox_1.setText(path[0])

    # Диалоговое окно сохранения
    def SaveFile(self, group, year):
        name =  f'Респа для {group} на {year} год' + '.xlsx'
        path = QFileDialog.getSaveFileName(self,
                                                     'Сохранить файл', # Название диалогового окна
                                                     name,             # Имя файла по умолчанию
                                                     'EXCEL таблицы(*.xlsx)') # Формат для сохранения

        if not path[0].endswith('.xlsx'): # Если нужно доставить формат
            return f'{path[0]}.xlsx'
        else: # Если всё ок
            return path[0]

    # Подгрузка основной инфы с листа
    def SheetInfo(self):
        ind = self.ui.comboBox_1.currentText()
        self.sheet = self.book[ind]
        self.groups = crr.choise_group(self.sheet)
        if type(self.groups).__name__ == 'BadDataError':
            self.ErDo(self.groups)
        else:
            self.ui.listWidget.addItem(f'Список групп успешно считан!')
            self.ui.comboBox_2.addItems(list(map(str, self.groups[2])))
            self.ui.comboBox_2.blockSignals(False)

    # Обнуление подгруженной информации
    def Disintegration(self):
        # Очистка свойств класса
        self.path = self.book = self.sheet = self.groups = ''
        # Очистка листов
        self.ui.comboBox_1.blockSignals(True)
        self.ui.comboBox_1.clear()
        self.ui.comboBox_1.clearEditText()
        # Очистка групп на листе
        self.ui.comboBox_2.blockSignals(True)
        self.ui.comboBox_2.clear()
        self.ui.comboBox_2.clearEditText()

    # Открытие файла (происходит при подтверждении пути к файлу)
    def ConfirmPath(self):
        # Обнуление атрибутов
        self.Disintegration()
        # Считывание пути с текстового поля пути
        self.path = self.ui.textBox_1.text()
        if self.ui.listWidget.currentRow() > 0:
            self.ui.listWidget.addItem(' ')
        self.ui.listWidget.addItems([f'Подключаемый путь:', f'«{self.path}»'])
        # Попытка считать книгу
        self.book = crr.read_book(self.path)
        # Если попытка провалилась, обнулить атрибуты и вывести ошибку
        if type(self.book).__name__ == 'BadDataError':
            self.ErDo(self.book)
        # Если всё норм, продолжить подключение файла
        else:
            self.ui.listWidget.addItem(f'Файл расписания успешно подключен!')
            # Попытка считать список листов в книге (бывают книги без листов)
            sheets = crr.choise_sheet(self.book)
            if type(sheets).__name__ == 'BadDataError':
                self.ErDo(self.sheets)
            else:
                self.ui.listWidget.addItem(f'Список листов успешно считан!')
                # Добавить считанные названия листов в меню выбора листа
                self.ui.comboBox_1.addItems(sheets)
                self.ui.comboBox_1.blockSignals(False)
                # Попытка считать с листа инфу о группах и т.п
                self.SheetInfo()
        
    # Конвертация расписания
    def Upgrade(self):
        # Если лист определён, то есть расписание, которое можно обработать
        if self.sheet:
            """ Временное отключение кнопок """
            self.ui.pushButton_1.setEnabled(False)
            self.ui.pushButton_2.setEnabled(False)
            self.ui.pushButton_3.setEnabled(False)

            """ Считывание показаний с переключателей """
            timey = self.groups[0]                        # Период расписания
            year  = self.groups[1]                        # Год расписания
            row_s = self.groups[3]                        # Начальная строка расписательной части
            group = self.ui.comboBox_2.currentText()      # Комбобокс выбора группы
            grp2  = self.ui.comboBox_3.currentIndex()     # Комбобокс выбора подгруппы "из 2-х"
            grp3  = self.ui.comboBox_4.currentIndex()     # Комбобокс выбора подгруппы "из 3-х"
            f1    = bool(self.ui.checkBox_1.checkState()) # Чекбокс сокращения препода
            f2    = bool(self.ui.checkBox_2.checkState()) # Чекбокс сокращения предмета
            f3    = bool(self.ui.checkBox_3.checkState()) # Чекбокс сокращения кабинета
            f4    = bool(self.ui.checkBox_4.checkState()) # Чекбокс сокращения подгруппы

            """ Обработка файла """
            # Ошибки в файле автоматически корректируются, но возможны исключения

            # Считывание расписания для выбранной группы
            book = crr.prepare(self.sheet, group, row_s, f3)
            # Если считывание провалилось, дропнуть ошибку
            if type(book).__name__ == 'BadDataError':
                self.ErDo(self.book)
            else:
                # Парсинг считанного расписания
                book = crp.parser(book, year, f1, f2, f4)
                # Если парсинг провалился, дропнуть ошибку
                if type(book).__name__ == 'BadDataError':
                    self.ErDo(self.book)
                else:
                    # Анализ запарсенного расписания 
                    a_bd = cra.analyze_bd(book)
                    # Если анализ провалился, дропнуть ошибку
                    if type(a_bd).__name__ == 'BadDataError':
                        self.ErDo(self.a_bd)
                    else:
                        # Форматирование расписания (с учётом анализа)
                        book = crw.create_resp(book, a_bd, grp2, grp3, timey, year)
                        # Если форматирование провалилось, дропнуть ошибку
                        if type(book).__name__ == 'BadDataError':
                            self.ErDo(self.book)
                        else:
                            # Выбор пути (путь к файлу, путь воды, путь огня, путь военкомата)
                            path = self.SaveFile(group, year)
                            if path == '.xlsx':
                                self.ErDo('Сохранение было отменено пользователем.')
                            else:
                                try:
                                    # Попытка сохранить расписание
                                    crw.save_resp(book, path)
                                    self.ui.listWidget.addItems([f'Расписание сохранено как:', f'«{path}»', ''])
                                except:
                                    self.ErDo('Сохранение не удалось!')

            """ Обратное включение кнопок """
            self.ui.pushButton_1.setEnabled(True)
            self.ui.pushButton_2.setEnabled(True)
            self.ui.pushButton_3.setEnabled(True)

        # Если лист не определён, то нет и расписания
        else:
            self.ErDo('Нечего конвертировать!')

    # Сворачивание логов
    def Turn(self):
        self.ui.toogleButton_1.setEnabled(False)
        if self.height() > 500:
            self.ui.toogleButton_1.setToolTip('Развернуть логи')
            self.ui.toogleButton_1.setText('▼')
            self.setFixedSize(581, 317)
        else:
            self.ui.toogleButton_1.setToolTip('Свернуть логи')
            self.ui.toogleButton_1.setText('▲')
            self.setFixedSize(581, 678)
        self.ui.toogleButton_1.setEnabled(True)

    # Копирование лога в буфер обмена
    def CopyLog(self, item):
        cp(item.text())
        self.ui.listWidget.addItems([f'Лог «{repr(item.text())}» скопирован!', ''])


""" Запуск интерфейса """
# Создание процесса приложения
app = QApplication([])

# Инициализация объекта окна приложения
application = my_window()

# Запуск окна приложения
application.show()

# Выход из приложения
close_app(app.exec())
