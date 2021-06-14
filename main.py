""" Описание модуля
Стартовый модуль с прикручиванием UX к макету GUI, и запуском приложения

Компиляция: "pyinstaller --onefile Compile.spec"
Генерация зависмостей: pip freeze > requirements.txt
Восстановление зависимостей: pip install -r requirements.txt
"""

""" Информация о приложении """
ORGANIZATION_NAME = 'DK905'
ORGANIZATION_DOMAIN = 'vk.com/dk905'
APPLICATION_NAME = 'CoolResp'
SETTINGS_TRAY = 'settings/tray'

""" Подключение модулей обработки """
# Импорт умолчаний API
from CoolRespProject.modules import CR_defaults as crd
# Считывание таблицы в базу разбора для конкретной группы
from CoolRespProject.modules import CR_reader as crr
# Парсинг базы разбора в нормальную БД для каждой логической записи
from CoolRespProject.modules import CR_parser as crp
# Швейцарский нож для дополнительной обработки
from CoolRespProject.modules import CR_swiss as crs
# Форматная запись БД в таблицу EXCEL
from CoolRespProject.modules import CR_writter as crw

""" Подключение элементов GUI """
from CoolRespProject.GUI.PC.main_gui import Ui_MainWindow
from PyQt5.QtWidgets import QApplication, QFileDialog, QMainWindow, QWidget
from PyQt5.QtCore import QCoreApplication, QSettings
from webbrowser import open as open_link
from pyperclip import copy as cp
from sys import exit as close_app


""" Задание интерфейса """


class CoolRespWindow(QMainWindow):
    """ Атрибуты для обработки расписания """
    path   = ''  # Путь к обрабатываемой книге
    book   = ''  # Открытая книга
    sheet  = ''  # Выбор листа
    groups = ''  # Строка с инфой о группах, периоде и т.п

    def __init__(self):
        super(CoolRespWindow, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        # Объект настроек приложения
        self.ui.settings = QSettings('CoolResp', 'DK905', self)
        self.load_settings()

        # Реакция на выбор репозитория
        self.ui.action_1.triggered.connect(lambda: open_link('https://github.com/DK905/CoolResp'))

        # Реакция на вызов краткой справки
        self.ui.action_2.triggered.connect(lambda: self.ui.msgInf.exec())

        # Реакция на вызов технической инфы о проге
        self.ui.action_3.triggered.connect(lambda: self.ui.msgAbt.exec())

        # Реакция на вызов диалогового окна выбора загружаемого файла
        self.ui.pushButton_1.clicked.connect(self.open_file)

        # Реакция на подтверждение пути
        self.ui.pushButton_2.clicked.connect(self.confirm_path)

        # Реакция на смену листа
        self.ui.comboBox_1.currentIndexChanged.connect(self.sheet_info)

        # Реакция на кнопку конвертации
        self.ui.pushButton_3.clicked.connect(self.upgrade)

        # Реакция на кнопку сворачивания логов
        self.ui.toogleButton_1.clicked.connect(self.turn_logs)

        # Реакция на двойной клик по логу
        self.ui.listWidget.itemDoubleClicked.connect(self.copy_logs)

    # Закрытие приложения должно сохранять настройки
    def closeEvent(self, event):
        self.save_settings()
        super().closeEvent(event)

    # Функция сохранения настроек приложения
    def save_settings(self):
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
    def load_settings(self):
        try:
            # Группа общих настроек
            self.ui.settings.beginGroup('All')
            if self.ui.settings.value('PathFile'):
                self.ui.textBox_1.setText(self.ui.settings.value('PathFile'))
            if self.ui.settings.value('Switching') and self.ui.settings.value('Switching') != self.ui.toogleButton_1.text():
                self.turn_logs()
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
    def error_action(self, text):
        text = str(text)
        self.ui.listWidget.addItems([text, ''])
        self.ui.msgErr.setText(text)
        self.ui.msgErr.exec()

    # Диалоговое окно загрузки файла
    def open_file(self):
        path = QFileDialog.getOpenFileName(self,
                                           'Выбрать файл',                         # Название диалогового окна
                                           '.',                                    # Имя файла по умолчанию
                                           'EXCEL таблицы(*.xls*);;Все файлы(*)')  # Поддерживаемые типы файлов
        self.ui.textBox_1.setText(path[0])

    # Диалоговое окно сохранения
    def save_file(self, book):
        name = f'{crs.create_name(book)}.xlsx'
        path = QFileDialog.getSaveFileName(self,
                                           'Сохранить файл',         # Название диалогового окна
                                           name,                     # Имя файла по умолчанию
                                           'EXCEL таблицы(*.xlsx)')  # Формат для сохранения

        if not path[0].endswith('.xlsx'):  # Если нужно дописать формат
            return f'{path[0]}.xlsx'
        else:  # Если всё ок
            return path[0]

    # Подгрузка основной инфы с листа
    def sheet_info(self):
        sheet_name = self.ui.comboBox_1.currentText()

        stage = 'Поиск групп на листе расписания'
        try:
            self.sheet = crr.take_sheet(self.book, sheet_name)
            self.groups = crr.choise_group(self.sheet)
            self.ui.listWidget.addItem(f'Список групп успешно считан!')
            # Обнуление списка групп
            self.ui.comboBox_2.blockSignals(True)
            self.ui.comboBox_2.clear()
            # Добавление новых групп
            self.ui.comboBox_2.addItems(list(map(str, self.groups[2])))
            self.ui.comboBox_2.blockSignals(False)

        except Exception as msg:
            self.error_action(f'Ошибка на этапе «{stage}»\n{msg}')

    # Обнуление подгруженной информации
    def disintegration(self):
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
    def confirm_path(self):
        # Обнуление атрибутов
        self.disintegration()

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
            sheets = crr.see_sheets(self.book)
            self.ui.listWidget.addItem(f'Список листов успешно считан!')

            # Добавить считанные названия листов в меню выбора листа
            stage = 'Добавление листов в меню выбора'
            self.ui.comboBox_1.addItems(sheets)
            self.ui.comboBox_1.blockSignals(False)

            # Попытка считать с листа инфу о группах и т.п
            stage = 'Считывание инфы о группах с листа'
            self.sheet_info()

        except Exception as msg:
            self.error_action(f'Ошибка на этапе «{stage}»\n{msg}')
            self.disintegration()

    # Конвертация расписания
    def upgrade(self):
        # Если лист определён, то есть расписание, которое можно обработать
        if self.sheet:
            """ Временное отключение кнопок """
            self.ui.pushButton_1.setEnabled(False)
            self.ui.pushButton_2.setEnabled(False)
            self.ui.pushButton_3.setEnabled(False)

            """ Обработка файла """
            # Ошибки в файле автоматически корректируются, но исключения всегда найдут путь

            try:
                # Считывание расписания для выбранной группы
                stage = 'Считывание расписания из файла'
                bd_process = crr.prepare(self.sheet,                        # Выбранный лист
                                         self.ui.comboBox_2.currentText(),  # Выбранная группа
                                         self.groups[3])                    # Диапазон расписания

                # Парсинг считанного расписания
                stage = 'Парсинг расписания'
                bd_parse = crp.parser(bd_process,      # Датасет расписания группы
                                      self.groups[0],  # Период расписания
                                      self.groups[1]   # Год расписания
                                      )

                # Форматирование расписания
                stage = 'Форматирование расписания'
                book = crw.create_resp(bd_parse,                                   # БД расписания
                                       str(self.ui.comboBox_3.currentIndex()),     # Подгруппа, где две подгруппы
                                       str(self.ui.comboBox_4.currentIndex()),     # Подгруппа, где три подгруппы
                                       bool(self.ui.checkBox_2.checkState()),      # Флаг сокращения предмета
                                       bool(self.ui.checkBox_1.checkState()),      # Флаг сокращения препода
                                       not bool(self.ui.checkBox_4.checkState()),  # Флаг сокращения подгруппы
                                       bool(self.ui.checkBox_3.checkState()))      # Флаг сокращения корпуса кабинета

                # Выбор пути к файлу
                stage = 'Сохранение расписания'
                path = self.save_file(bd_parse)
                if path == '.xlsx':
                    self.error_action('Сохранение было отменено пользователем.')
                else:
                    # Попытка сохранения расписания
                    stage = 'Сохранение готового расписания'
                    crw.save_resp(book, path)
                    self.ui.listWidget.addItems([f'Расписание сохранено как:', f'«{path}»', ''])

            except Exception as msg:
                self.error_action(f'Ошибка на этапе «{stage}»\n{msg}')

            """ Обратное включение кнопок """
            self.ui.pushButton_1.setEnabled(True)
            self.ui.pushButton_2.setEnabled(True)
            self.ui.pushButton_3.setEnabled(True)

        # Если лист не определён, то нет и расписания
        else:
            self.error_action('Нечего конвертировать!')

    # Сворачивание логов
    def turn_logs(self):
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
    def copy_logs(self, item):
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
    application = CoolRespWindow()

    # Запуск окна приложения
    application.show()

    # Выход из приложения
    close_app(app.exec())
