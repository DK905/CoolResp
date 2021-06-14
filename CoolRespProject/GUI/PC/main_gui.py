""" Модуль интерфейса конвертера расписания
В этом модуле реализован макет и UI часть GUI конвертера расписания.
Интерфейс построен для ПК, так как для телефона планируется специальная версия.

"""

# Функция возврата абсолютного пути (нужна для подгрузки ресурсов в компилируемую версию)
def resource_path(relative):
    import sys
    from os import path as os_path
    if hasattr(sys, '_MEIPASS'):
        return os_path.join(sys._MEIPASS, relative)
    else:
        return os_path.join(os_path.abspath('.'), relative)

""" Модули макета """
from PyQt5.QtWidgets import QAction, QCheckBox, QComboBox, QDialog, QFrame, QGroupBox, QHBoxLayout, QLabel
from PyQt5.QtWidgets import QLineEdit, QListWidget, QMenu, QMenuBar, QMessageBox, QPushButton, QWidget
from PyQt5.QtCore    import Qt, QCoreApplication, QMetaObject, QRect
from PyQt5.QtGui     import QIcon, QPixmap


class AboutBox(QDialog):
    """ Класс окна справки """
    def __init__(self):
        QDialog.__init__(self)
        self.setWindowIcon(QIcon(resource_path('CResp.ico')))
        self.setFixedSize(1000, 558)
        self.label_1 = QLabel(self)
        self.label_1.setGeometry(QRect(5, 3, 450, 550))

        self.title = QLabel(self)
        self.title.setGeometry(QRect(477, 10, 500, 50))
        self.title.setStyleSheet("background-color: 'white';"
                                 "font: 30pt \'Bookman Old Style\';"
                                 "text-align: 'center'")
        self.title.setFrameStyle(QFrame.Box)
        self.title.setAlignment(Qt.AlignCenter)
        self.title.setText('ОБОЗНАЧЕНИЯ')

        self.label_2 = QListWidget(self)
        self.label_2.setGeometry(QRect(477, 70, 500, 478))
        self.label_2.setStyleSheet("font: 10pt \'Bookman Antiqua\'")
        self.label_2.setTabKeyNavigation(True)
        self.label_2.addItems(['1)  Раздел технической инфы;',
                               '2)  Текстовое поле для ввода пути к файлу расписания;',
                               '3)  Кнопка выбора файла расписания через диалоговое окно;',
                               '4)  Кнопка подтверждения выбора файла;',
                               '5)  Список листов в выбранном файле;',
                               '6)  Список групп на выбранном листе;',
                               '7)  Выбор подгруппы для пар с делением на две подгруппы;',
                               '8)  Выбор подгруппы для пар с делением на три подгруппы;',
                               '9)  Чекбокс для сокращения должности преподавателей;',
                               '10) Чекбокс для сокращения предметов;',
                               '11) Чекбокс для сокращения учебных корпусов у кабинетов;',
                               '12) Чекбокс для сокращения подгрупп до номера;',
                               '13) Кнопка для конвертации расписания;',
                               '14) Виджет, куда выводятся логи;',
                               '15) Переключатель отображения логов\n(при нажатии, виджет логов сворачивается).'])


class Ui_MainWindow(object):
    """ Класс основного окна приложения """

    def setupUi(self, MainWindow):
        """ Инициализация макета и UI """
        
        """ Базовые параметры главного окна """
        MainWindow.setObjectName("MainWindow")
        MainWindow.setFixedSize(581, 678)
        self.cw_main = QWidget(MainWindow)
        self.cw_main.setObjectName("cw_main")
        MainWindow.setCentralWidget(self.cw_main)

        """ Уведомление об ошибке """
        self.msgErr = QMessageBox(self.cw_main)
        self.msgErr.setWindowModality(True)
        self.msgErr.setObjectName("msgErr")    

        """ Блок верхнего меню """
        # Меню
        self.menuBar = QMenuBar(MainWindow)
        self.menuBar.setGeometry(QRect(0, 0, 581, 26))
        self.menuBar.setObjectName("menuBar")
        self.menu = QMenu(self.menuBar)
        self.menu.setObjectName("menu")
        MainWindow.setMenuBar(self.menuBar)
        # Пункт №1
        self.action_1 = QAction(MainWindow)
        self.action_1.setObjectName("action_1")
        self.menu.addAction(self.action_1)
        # Пункт №2
        self.action_2 = QAction(MainWindow)
        self.action_2.setObjectName("action_2")
        self.menu.addAction(self.action_2)
        # Пункт №3
        self.action_3 = QAction(MainWindow)
        self.action_3.setObjectName("action_3")
        self.menu.addAction(self.action_3)
        
        self.menuBar.addAction(self.menu.menuAction())

        """ Краткая справка """
        self.msgInf = AboutBox()
        self.msgInf.setObjectName("msgInf")

        """ Техническая инфа """
        self.msgAbt = QMessageBox(self.cw_main)
        self.msgAbt.setWindowModality(True)
        self.msgAbt.setObjectName("msgAbt")

        """ Блок загрузки """
        # Блок
        self.groupBox_1 = QGroupBox(self.cw_main)
        self.groupBox_1.setGeometry(QRect(20, 10, 541, 61))
        self.groupBox_1.setObjectName("groupBox_1")
        # Поле ручного ввода пути
        self.textBox_1 = QLineEdit(self.groupBox_1)
        self.textBox_1.setGeometry(QRect(10, 20, 441, 31))
        self.textBox_1.setObjectName("textBox_1")
        # Кнопка выбора файла через диалог
        self.pushButton_1 = QPushButton(self.groupBox_1)
        self.pushButton_1.setGeometry(QRect(460, 20, 31, 31))
        self.pushButton_1.setObjectName("pushButton_1")
        # Кнопка подтверждения выбора
        self.pushButton_2 = QPushButton(self.groupBox_1)
        self.pushButton_2.setGeometry(QRect(500, 20, 31, 31))
        self.pushButton_2.setObjectName("pushButton_2")        

        """ Блок инфы, подгружаемой с листа """
        # Блок
        self.groupBox_2 = QGroupBox(self.cw_main)
        self.groupBox_2.setGeometry(QRect(20, 80, 301, 91))
        self.groupBox_2.setObjectName("groupBox_2")
        # Выбор листа в книге EXCEL
        self.groupBox_3 = QGroupBox(self.groupBox_2)
        self.groupBox_3.setGeometry(QRect(10, 20, 151, 61))
        self.groupBox_3.setObjectName("groupBox_3")
        self.comboBox_1 = QComboBox(self.groupBox_3)
        self.comboBox_1.setGeometry(QRect(10, 20, 131, 31))
        self.comboBox_1.setObjectName("comboBox_1")
        # Выбор группы на листе
        self.groupBox_4 = QGroupBox(self.groupBox_2)
        self.groupBox_4.setGeometry(QRect(180, 20, 111, 61))
        self.groupBox_4.setObjectName("groupBox_4")
        self.comboBox_2 = QComboBox(self.groupBox_4)
        self.comboBox_2.setGeometry(QRect(10, 20, 91, 31))
        self.comboBox_2.setObjectName("comboBox_2")

        """ Блок выбора подгрупп """
        # Блок
        self.groupBox_5 = QGroupBox(self.cw_main)
        self.groupBox_5.setGeometry(QRect(340, 80, 221, 91))
        self.groupBox_5.setObjectName("groupBox_5")
        # Деление на две подгруппы
        self.groupBox_6 = QGroupBox(self.groupBox_5)
        self.groupBox_6.setGeometry(QRect(10, 20, 91, 61))
        self.groupBox_6.setObjectName("groupBox_6")
        self.comboBox_3 = QComboBox(self.groupBox_6)
        self.comboBox_3.setGeometry(QRect(10, 20, 73, 31))
        self.comboBox_3.setObjectName("comboBox_3")
        self.comboBox_3.addItem("")
        self.comboBox_3.addItem("")
        self.comboBox_3.addItem("")
        # Деление на три подгруппы
        self.groupBox_7 = QGroupBox(self.groupBox_5)
        self.groupBox_7.setGeometry(QRect(120, 20, 91, 61))
        self.groupBox_7.setObjectName("groupBox_7")
        self.comboBox_4 = QComboBox(self.groupBox_7)
        self.comboBox_4.setGeometry(QRect(10, 20, 73, 31))
        self.comboBox_4.setObjectName("comboBox_4")
        self.comboBox_4.addItem("")
        self.comboBox_4.addItem("")
        self.comboBox_4.addItem("")
        self.comboBox_4.addItem("")

        """ Блок сокращений """
        # Блок
        self.groupBox_8 = QGroupBox(self.cw_main)
        self.groupBox_8.setGeometry(QRect(20, 180, 541, 60))
        self.groupBox_8.setObjectName("groupBox_8")
        self.horizontalLayout = QHBoxLayout(self.groupBox_8)
        self.horizontalLayout.setObjectName("horizontalLayout")
        # Переключатель "сокращать должности преподов?"
        self.checkBox_1 = QCheckBox(self.groupBox_8)
        self.checkBox_1.setObjectName("checkBox_1")
        self.horizontalLayout.addWidget(self.checkBox_1)
        # Переключатель "сокращать названия предметов?"
        self.checkBox_2 = QCheckBox(self.groupBox_8)
        self.checkBox_2.setObjectName("checkBox_2")
        self.horizontalLayout.addWidget(self.checkBox_2)
        # Переключатель "сокращать корпуса кабинетов?"
        self.checkBox_3 = QCheckBox(self.groupBox_8)
        self.checkBox_3.setObjectName("checkBox_3")
        self.horizontalLayout.addWidget(self.checkBox_3)
        # Переключатель "сокращать п/гр в подгруппах?"
        self.checkBox_4 = QCheckBox(self.groupBox_8)
        self.checkBox_4.setObjectName("checkBox_4")
        self.horizontalLayout.addWidget(self.checkBox_4)

        """ Блок конвертации """
        self.pushButton_3 = QPushButton(self.cw_main)
        self.pushButton_3.setGeometry(QRect(20, 250, 491, 31))
        self.pushButton_3.setObjectName("pushButton_3")

        """ Блок логов """
        # Кнопка сворачивания логов
        self.toogleButton_1 = QPushButton(self.cw_main)
        self.toogleButton_1.setGeometry(QRect(520, 250, 41, 31))
        self.toogleButton_1.setObjectName("toogleButton_1")
        # Виджет вывода логов
        self.listWidget = QListWidget(self.cw_main)
        self.listWidget.setGeometry(QRect(20, 295, 541, 341))
        self.listWidget.setStyleSheet("font: 9pt \'Lucida Console\'")
        self.listWidget.setTabKeyNavigation(True)
        self.listWidget.setObjectName("listWidget")

        """ Инициализация визуала и функционала """
        self.retranslateUi(MainWindow)
        QMetaObject.connectSlotsByName(MainWindow)


    def retranslateUi(self, MainWindow):
        """ Настройка графической составляющей интерфейса главного окна """

        _translate = QCoreApplication.translate
        """ Параметры основного окна """
        MainWindow.setWindowTitle(_translate("MainWindow", 'CoolResp'))
        MainWindow.setWindowIcon(QIcon(resource_path('CResp.ico')))

        """ Сообщение об ошибке """
        self.msgErr.setWindowTitle(_translate("MainWindow", 'Ошибка'))
        self.msgErr.setIcon(QMessageBox.Icon.Warning)

        """ Верхнее меню """
        self.menu.setTitle(_translate("MainWindow", 'Помощь'))
        self.menuBar.setToolTip('Краткая инфа')
        # Пункт №1
        self.action_1.setText(_translate("MainWindow", "GitHub"))
        self.action_1.setToolTip('Открыть репозиторий проекта')
        self.action_1.setIcon(QIcon(resource_path('Git.ico')))
        # Пункт №2
        self.action_2.setText(_translate("MainWindow", "Краткая справка"))
        self.action_2.setToolTip('Небольшая справка\nпо проге')
        self.action_2.setIcon(QIcon(resource_path('Help.ico')))
        self.action_2.setShortcut('F1')
        self.msgInf.setWindowTitle(_translate("MainWindow", 'Краткая справка'))
        self.msgInf.label_1.setPixmap(QPixmap(resource_path('Short_Help.png')))
        # Пункт №3
        self.action_3.setText(_translate("MainWindow", "О программе"))
        self.action_3.setToolTip('Инфа о разработке')
        self.action_3.setIcon(QIcon(resource_path('About.ico')))
        self.msgAbt.setWindowTitle(_translate("MainWindow", 'О программе'))
        self.msgAbt.setIcon(QMessageBox.Icon.Information)
        self.msgAbt.setText('''
CoolResp
Язык разработки: python 3.8.3
Автор: Мирославский И.С. (DK905)
universus114@mail.ru
                            ''')

        """ Блок загрузки """
        self.groupBox_1.setTitle(_translate("MainWindow", 'Загрузка'))
        # Поле ручного ввода пути
        self.textBox_1.setText('Путь к файлу расписания')
        self.textBox_1.setToolTip('Ручной ввод пути к файлу')
        # Кнопка выбора файла через диалог
        self.pushButton_1.setText(_translate("MainWindow", '...'))
        self.pushButton_1.setToolTip('Открыть диалог загрузки файла')
        self.pushButton_1.setShortcut('Ctrl+O')
        # Кнопка подтверждения выбора
        self.pushButton_2.setToolTip('Подтвердить выбор пути\n(Загружает файл)')
        self.pushButton_2.setIcon(QIcon(resource_path('CheckMark.ico')))

        """ Блок инфы, подгружаемой с листа """
        self.groupBox_2.setTitle(_translate("MainWindow", 'Базовая инфа'))
        # Выбор листа в книге EXCEL
        self.groupBox_3.setTitle(_translate("MainWindow", 'Лист'))
        self.comboBox_1.setToolTip('Выбрать лист из книги EXCEL')
        # Выбор группы на листе
        self.groupBox_4.setTitle(_translate("MainWindow", 'Группа'))
        self.comboBox_2.setToolTip('Выбрать группу с листа')

        """ Блок выбора подгрупп """
        self.groupBox_5.setTitle(_translate("MainWindow", 'Выбор своей подгруппы'))
        # Деление на две подгруппы
        self.groupBox_6.setTitle(_translate("MainWindow", 'Из 2-х'))
        self.groupBox_6.setToolTip('Выбрать подгруппу "где две подгруппы"')
        self.comboBox_3.setItemText(0, _translate("MainWindow", 'Пофиг'))
        self.comboBox_3.setItemText(1, _translate("MainWindow", '1'))
        self.comboBox_3.setItemText(2, _translate("MainWindow", '2'))
        # Деление на три подгруппы
        self.groupBox_7.setTitle(_translate("MainWindow", 'Из 3-х'))
        self.groupBox_7.setToolTip('Выбрать подгруппу "где три подгруппы"')
        self.comboBox_4.setItemText(0, _translate("MainWindow", 'Пофиг'))
        self.comboBox_4.setItemText(1, _translate("MainWindow", '1'))
        self.comboBox_4.setItemText(2, _translate("MainWindow", '2'))
        self.comboBox_4.setItemText(3, _translate("MainWindow", '3'))

        """ Блок сокращений """
        self.groupBox_8.setTitle(_translate("MainWindow", 'Сокращения'))
        # Переключатель "сокращать должности преподов?"
        self.checkBox_1.setText(_translate("MainWindow", 'Преподаватели'))
        self.checkBox_1.setToolTip('Сокращать должность преподов?')
        # Переключатель "сокращать названия предметов?"
        self.checkBox_2.setText(_translate("MainWindow", 'Предметы'))
        self.checkBox_2.setToolTip('Сокращать названия предметов?')
        # Переключатель "сокращать корпуса кабинетов?"
        self.checkBox_3.setText(_translate("MainWindow", 'Кабинеты'))
        self.checkBox_3.setToolTip('Сокращать корпуса кабинетов?')
        # Переключатель "сокращать п/гр в подгруппах?"
        self.checkBox_4.setText(_translate("MainWindow", 'Подгруппы'))
        self.checkBox_4.setToolTip('Скрывать "п/гр" у подгрупп?')

        """ Блок конвертации """
        self.pushButton_3.setText(_translate("MainWindow", 'Конвертировать'))
        self.pushButton_3.setToolTip('Нажать для конвертации расписания')
        self.pushButton_3.setShortcut('Ctrl+S')

        """ Блок логов """
        # Кнопка сворачивания логов
        self.toogleButton_1.setText(_translate("MainWindow", '▲'))
        self.toogleButton_1.setToolTip('Свернуть логи')
