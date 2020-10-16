""" Описание модуля
Стартовый модуль с прикручиванием UX к макету GUI, и запуском приложения

"""

""" Информация о приложении """
ORGANIZATION_NAME = 'DK905'
ORGANIZATION_DOMAIN = 'vk.com/dk905'
APPLICATION_NAME = 'CoolResp'
SETTINGS_TRAY = 'settings/tray'

""" Подключение модулей обработки """
from modules.CR_dataset import er_list  # Импорт списка ошибок и рекомендаций по их фиксу
from modules import CR_reader  as crr   # Считывание таблицы в базу разбора для конкретной группы
from modules import CR_parser  as crp   # Парсинг базы разбора в нормальную БД для каждой логической записи
from modules import CR_analyze as cra   # Анализатор БД, для подправки косяков и определения числа подгрупп
from modules import CR_writter as crw   # Форматная запись БД в таблицу EXCEL

""" Подключение элементов GUI """
from PyQt5.QtWidgets import QApplication, QFileDialog, QMainWindow, QWidget
from GUI.PC.main_gui import Ui_MainWindow
from PyQt5.QtCore    import QCoreApplication, QSettings
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

        # Объект настроек приложения
        self.ui.settings = QSettings('CoolResp', 'DK905', self)
        self.LoadSet()

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

    # Закрытие приложения должно сохранять настройки
    def closeEvent(self, event):
        self.SaveSet()
        super().closeEvent(event)

    # Функция сохранения настроек приложения
    def SaveSet(self):
        try:
            self.ui.settings.beginGroup('All')
            self.ui.settings.setValue('PathFile',  self.ui.textBox_1.text())
            self.ui.settings.setValue('Switching', self.ui.toogleButton_1.text())
            self.ui.settings.endGroup()

            self.ui.settings.beginGroup('Briefs')
            self.ui.settings.setValue('Prepods',  self.ui.checkBox_1.checkState())
            self.ui.settings.setValue('Predmets', self.ui.checkBox_2.checkState())
            self.ui.settings.setValue('Cabinets', self.ui.checkBox_3.checkState())
            self.ui.settings.setValue('Groups',   self.ui.checkBox_4.checkState())
            self.ui.settings.endGroup()

            self.ui.settings.beginGroup('NGroups')
            self.ui.settings.setValue('Two',   self.ui.comboBox_3.currentIndex())
            self.ui.settings.setValue('Three', self.ui.comboBox_4.currentIndex())            
            self.ui.settings.endGroup()

        except:
            pass

    # Функция загрузки настроек приложения
    def LoadSet(self):
        try:
            # Группа общих настроек
            self.ui.settings.beginGroup('All')
            if self.ui.settings.value('PathFile'):
                self.ui.textBox_1.setText(self.ui.settings.value('PathFile'))
            if self.ui.settings.value('Switching') and self.ui.settings.value('Switching') != self.ui.toogleButton_1.text():
                self.Turn()
            self.ui.settings.endGroup()

            # Группа сокращалок
            self.ui.settings.beginGroup('Briefs')
            if self.ui.settings.value('Prepods'):
                self.ui.checkBox_1.setCheckState(self.ui.settings.value('Prepods'))
            if self.ui.settings.value('Predmets'):
                self.ui.checkBox_2.setCheckState(self.ui.settings.value('Predmets'))
            if self.ui.settings.value('Cabinets'):
                self.ui.checkBox_3.setCheckState(self.ui.settings.value('Cabinets'))
            if self.ui.settings.value('Groups'):
                self.ui.checkBox_4.setCheckState(self.ui.settings.value('Groups'))
            self.ui.settings.endGroup()

            # Группа выбора подгрупп
            self.ui.settings.beginGroup('NGroups')
            if self.ui.settings.value('Two'):
                self.ui.comboBox_3.setCurrentIndex(self.ui.settings.value('Two'))
            if self.ui.settings.value('Three'):
                self.ui.comboBox_4.setCurrentIndex(self.ui.settings.value('Three'))
            self.ui.settings.endGroup()

        except:
            pass

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

        try:
            stage = 'Поиск групп на листе расписания'
            self.groups = crr.choise_group(self.sheet)
            self.ui.listWidget.addItem(f'Список групп успешно считан!')
            self.ui.comboBox_2.addItems(list(map(str, self.groups[2])))
            self.ui.comboBox_2.blockSignals(False)

        except Exception as msg:
            self.ErDo(f'Ошибка на этапе «{stage}»\n{msg}')

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

        try:
            # Попытка считать книгу
            stage = 'Считывание книги'
            self.book = crr.read_book(self.path)
            self.ui.listWidget.addItem(f'Файл расписания успешно подключен!')

            # Попытка считать список листов в книге (бывают книги без листов)      
            stage = 'Получение списка листов в книге'
            sheets = crr.choise_sheet(self.book)
            self.ui.listWidget.addItem(f'Список листов успешно считан!')

            # Добавить считанные названия листов в меню выбора листа
            stage = 'Добавление листов в меню выбора'
            self.ui.comboBox_1.addItems(sheets)
            self.ui.comboBox_1.blockSignals(False)

            # Попытка считать с листа инфу о группах и т.п
            stage = 'Считывание инфы о группах с листа'
            self.SheetInfo()

        except Exception as msg:
            self.ErDo(f'Ошибка на этапе «{stage}»\n{msg}')
            
        
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
            f3    = bool(self.ui.checkBox_3.checkState()) # Чекбокс урезки корпуска кабинета
            f4    = bool(self.ui.checkBox_4.checkState()) # Чекбокс сокращения подгруппы

            """ Обработка файла """
            # Ошибки в файле автоматически корректируются, но исключения всегда найдут путь

            try:
                # Считывание расписания для выбранной группы
                stage = 'Считывание расписания из файла'
                book = crr.prepare(self.sheet, group, row_s, f3)

                # Парсинг считанного расписания
                stage = 'Парсинг расписания'
                book = crp.parser(book, timey, year, f1, f2, f4)

                # Анализ запарсенного расписания
                stage = 'Анализ расписания'
                a_bd = cra.analyze_bd(book)

                # Форматирование расписания (с учётом анализа)
                stage = 'Форматирование расписания'
                book = crw.create_resp(book, a_bd, grp2, grp3, timey, year)

                # Выбор пути (путь к файлу, путь воды, путь огня, путь военкомата)
                path = self.SaveFile(group, year)
                if path == '.xlsx':
                    self.ErDo('Сохранение было отменено пользователем.')
                else:
                    # Попытка сохранения расписания. Возможно не удастся из-за прав доступа к памяти или нехватки места
                    stage = 'Сохранение готового расписания'
                    crw.save_resp(book, path)
                    self.ui.listWidget.addItems([f'Расписание сохранено как:', f'«{path}»', ''])

            except Exception as msg:
                self.ErDo(f'Ошибка на этапе «{stage}»\n{msg}')

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
        if self.ui.toogleButton_1.text() == '▲':
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
if __name__ == '__main__':
    # Подключение свойств приложения
    QCoreApplication.setApplicationName(ORGANIZATION_NAME)
    QCoreApplication.setOrganizationDomain(ORGANIZATION_DOMAIN)
    QCoreApplication.setApplicationName(APPLICATION_NAME)

    # Создание процесса приложения
    app = QApplication([])

    # Инициализация объекта окна приложения
    application = my_window()

    # Запуск окна приложения
    application.show()

    # Выход из приложения
    close_app(app.exec())
