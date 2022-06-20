r"""Кастомные исключения.

"""


class BasicException(Exception):
    """ Класс базового исключения """
    pass


""" Исключения модуля взаимодействия с Excel """


class FileNotExcel(BasicException):
    """ Открываемый файл не является Excel документом """

    def __init__(self):
        super().__init__(
            f'Не удалось открыть файл как Excel-книгу!'
        )


class CantFoundSheets(BasicException):
    """ Excel документ не содержит листов """

    def __init__(self):
        super().__init__(
            f'Не удалось обнаружить листы в Excel-книге!'
        )


class CantGetSheet(BasicException):
    """ Запрашиваемый лист невозможно получить """

    def __init__(self):
        super().__init__(
            f'Не удалось загрузить запрашиваемый лист Excel-книги!'
        )


class CantFoundPositionInfo(BasicException):
    """ Не обнаружены ключевые метки для позиционирования расписания """

    def __init__(self):
        super().__init__(
            f'Невозможно выявить содержательную часть расписания!\n'
            f'На листе не обнаружены позиционные метки. Образец:\n'
            f'https://github.com/DK905/CoolResp/blob/master/Design/Типовой%20шаблон.png'
        )


""" Исключения парсера """


class IncorrectDate(BasicException):
    """  """

    def __init__(self, cell, date):
        super().__init__(
            f'Набор дат некорректен!\n'
            f'«{cell}»\n'
            f'Даты: «{date}»'
        )


class IncorrectCell(BasicException):
    """  """

    def __init__(self, cell):
        super().__init__(
            f'Что-то не так в ячейке!\n'
            f'«{cell}»'
        )
