r"""Взаимодействие с исходным Excel-документом.
Данный модуль предназначен для:
    - Считывания изначального файла расписания из .xls или .xlsx;
    - Вычленения расписания конкретной выбранной группы из считанного файла.

"""

import CoolRespProject.modules_api.api_defaults as api_def
import CoolRespProject.modules_parser.cr_defaults as crd
import CoolRespProject.modules_parser.cr_swiss as crs
import pandas as pd
import numpy as np
import xlrd
import re


def read_book(file: 'Путь к загружаемому файлу',
              ext:  'Расширение документа' = ''
              ) ->  'Объект книги в формате типизированного словаря':

    """ Функция для считывания Excel-документа (xls, xlsx) в объект xlrd-книги """

    # Если расширение не указано, попытаться определить его
    if not ext:
        ext = crs.check_file_extension(file)
        if ext == api_def.EXTENSIONS['404']:
            return crd.READING_BOOK

    # Если старый формат EXCEL 97-2003
    if ext == 'xls':
        # В этом формате, нужно отдельно запрашивать информацию о ячейках
        reading_book = xlrd.open_workbook(file, formatting_info=True)
    # В новых форматах изменена структура хранения данных и информации о ячейках
    elif ext == 'xlsx':
        # То есть, флаг formatting_info уже не используется
        reading_book = xlrd.open_workbook(file)
    else:
        reading_book = crd.READING_BOOK

    return reading_book


def see_sheets(book: 'Загруженный объект книги'
               ) ->  'Список листов книги (для меню выбора листов)':

    """ Функция для отображения листов в книге """
    
    # Вернуть список листов
    return book.sheet_names()


def take_sheet(book: 'Загруженный объект книги',
               name: 'Имя нужного листа'
               ) ->  'Объект выбранного листа':

    """ Функция получения листа из объекта книги """

    return book.sheet_by_name(name)


def group_choice(sheet: 'Объект листа с загруженной книги'
                 ) ->   'Словарь {Список групп на листе, индексный диапазон расписания}':

    """ Функция для выделения описательной информации о группах и границах на листе """
    period = year = grp_list = ind_start = ind_end = None

    # Проход по значениям строк выбранного листа
    for ind_row in range(sheet.nrows):
        row = sheet.row_values(ind_row)
        # Если период расписания и границы его диапазона ещё не найдены, то проверить актуальную строку на их наличие
        if not period or not ind_end:
            # Период ищется среди значений строки значений (из-за плавающего положения)
            for col in row:
                # Период ищется только среди данных строчного типа
                if isinstance(col, str):
                    # Таблица данных расписания начинается сразу после строки "На период..."
                    if not period and re.search(r'[Нн]а\s*период', col):
                        # Выделение периода расписания. Если он указан неправильно, потом будет ложная коррекция парсера
                        period = re.findall(r'с\s*([\d.]+)[г.]*\s*по\s*([\d.]+)[г.]*',
                                                 col,
                                                 re.IGNORECASE)[0]
                        # Период приводится к самому частому формату даты по типу "04.03.21"
                        period = ['.'.join([part[-2:] for part in date.split('.')]) for date in period]
                        # Выделение года
                        year = 2000 + int(period[0][-2:])
                        # После нахождения строки с периодом и выделения её данных, поиск завершается
                        continue
                    # Расписание всегда идёт вплоть до "начальник УО"
                    elif ind_start and re.search(r'(?:(?:начальник)|(?:методист))\s*уо',
                                                 col,
                                                 re.IGNORECASE):
                        ind_end = ind_row
                        break

        # Найти список групп и начальную позицию (они могут быть не сразу после строки периода)
        if not grp_list and re.search(r'^[Дд]ни', row[0], re.MULTILINE):
            grp_list = [el for el in row if el]  # Отсеивание совсем пустых столбцов
            grp_list = [grp_list[i] for i in range(2, len(grp_list), 2)]  # Группы в строке идут с двойным шагом ячейки
            grp_list = [crs.string_float_to_string_int(group) for group in grp_list]
            ind_start = ind_row  # Начало диапазона расписания - строка групп
            continue  # Если нашли начало, то рано проверять конец

        if ind_end:  # Поиск останавливается при нахождении конца диапазона расписания
            break  # То есть, цикл полностью завершается

    # Если на листе обнаружены "список групп", "период расписания" и "диапазон строк расписания", скомпановать их
    if period and grp_list and ind_start and ind_end:
        group_information = [period, year, grp_list, (ind_start + 1, ind_end)]
    else:
        group_information = crd.GROUP_INFORMATION
    # Обнаруженная информация возвращается в качестве словаря
    return dict(zip(['period', 'year', 'groups_info', 'range'], group_information))


"""                 Начальная стадия разбора расписания

Данная стадия предполагает подготовку расписания выбранной группы к парсингу.
Сначала, среди всех листов и всех групп выбирается одна конкретная группа, для которой
вычленяется собственная таблица расписания.
По итогу стадии, будет создана первичная БД обработки

"""


def take_value(row: 'Координата строки',
               col: 'Координата столбца',
               mgl: 'Список границ объединённых ячеек',
               act: 'Нужна актуальная ячейка, или её правый сосед?'
               ) -> 'Координаты ячейки, из которой нужно взять значение':

    """ Функция получения координат истинного значения ячейки (в том числе, объединённой) """
    
    # Индикатор объединённости ячейки
    indicator = None
    
    # Проверка актуальной ячейки на объединённость
    for diapason in mgl:
        row_range = range(diapason[0], diapason[1])
        col_range = range(diapason[2], diapason[3])
        # Если ячейка объединённая
        if row in row_range and col in col_range:
            indicator = diapason[0], diapason[2]
            break

    # Если нужно найти правого соседа актуальной ячейки
    if not act:
        # Если актуальная ячейка - объединённая
        if indicator:
            # То это учитывается в столбце
            indicator = take_value(row, diapason[3], mgl, True)
        # Если актуальная ячейка - обычная
        else:
            # Получить значение из правого столбца
            indicator = take_value(row, col+1, mgl, True)

    # Если нужная ячейка - обычная, то вернуть её координаты
    return indicator if indicator else (row, col)
    

def what_col(title: 'Шапка подтаблицы расписания',
             group: 'Выбранная группа'
             ) ->   'Индекс столбца группы':

    """ Функция определения столбца группы """
    for ind, rec in enumerate(title):
        if crs.string_float_to_string_int(rec) == crs.string_float_to_string_int(group):
            return ind

    return None


def prepare(sheet: 'Выбранный лист',
            group: 'Выбранная группа',
            coord: 'Диапазон информации о расписании'
            ) ->   'Датафрейм предварительных данных':

    """ Функция для подготовки расписания выбранной группы к парсингу """

    """ Этап сбора набора данных """
    # Получить индекс главного столбца информации о парах группы
    # Строка списка групп расположена на строку выше чем начало диапазона, поэтому -1
    g = what_col(sheet.row_values(coord[0] - 1), group)
    if not g:
        return None

    # Выделить список границ объединённых ячеек
    merge_coords = sorted(sheet.merged_cells)

    # Задание датафрейма для базы предобработки
    df_prep = pd.DataFrame(columns=['day',    # День недели
                                    'num',    # Номер пары
                                    'rec',    # Информация о парах
                                    'cabs'])  # Кабинеты

    # Построчно обработать лист и выделить полноценное расписание группы
    for row in range(*coord):
        # День недели берётся из нулевой ячейки строки
        act_day = sheet.cell(*take_value(row, 0, merge_coords, True)).value

        # Номер пары берётся из первой ячейки строки
        act_num = sheet.cell(*take_value(row, 1, merge_coords, True)).value

        # Информация о парах берётся из главной ячейки группы
        act_rec = sheet.cell(*take_value(row, g, merge_coords, True)).value

        # Информация о кабинетах берётся следующей ячейкой, от правой ячейки относительно главной
        act_cab = sheet.cell(*take_value(row, g, merge_coords, False)).value

        new_row = {'day' : act_day,
                   'num' : act_num,
                   'rec' : act_rec,
                   'cabs': act_cab}

        df_prep.loc[df_prep.shape[0]] = new_row

    """ Этап предварительной обработки вида данных """

    """ Общая редакция (1) """
    # Заменить все пропуски, которые содержат лишь пробельные символы, на пустые значения
    df_prep.replace(to_replace=r'^\s*$', value=np.nan, regex=True, inplace=True)

    # Заменить все пропуски в номерах пары предыдущим значением
    # Иногда, номер пары забывают сделать объединённой ячейкой, но он должен быть
    df_prep[['day', 'num']] = df_prep[['day', 'num']].fillna(method='ffill')

    # Удалить все строки без данных о парах
    df_prep.dropna(subset=['rec'], inplace=True)

    """ Дни недели """
    # Сократить дни недели
    df_prep['day'].replace(crd.DAYS_REGEX, regex=True, inplace=True)

    # Привести столбец номеров пар к типу категориальных данных (оптимизация)
    df_prep['day'] = df_prep['day'].astype('category')

    """ Номера пар """
    # Номера пар считались как float из-за особенностей xlrd и пропусков
    # То есть, столбец номеров пар нужно привести к целочисленному типу (оптимизация)
    df_prep['num'] = df_prep['num'].astype('int8')

    """ Информация о парах """
    # Пробельные последовательности заменяются одиночным пробелом
    df_prep['rec'].replace(r'\s+', ' ', regex=True, inplace=True)

    # Заменить диф.зачёты и зачёты на кодовую последовательность (стабилизация парсинга)
    df_prep['rec'].replace(crd.EXAM_TYPES, regex=True, inplace=True)

    """ Кабинеты """
    # Заменить различные сокращения актового зала на обобщённое значение
    df_prep['cabs'] = df_prep['cabs'].map(lambda val: re.sub(r'[аА].*?[лЛ]', crd.DEF_EVENT_CAB, str(val)) if val else val)

    # Регулярное выражение для физкультурного зала
    pat_zal = re.compile(r'(?:с\s*?/\s*?з.*?т\s*?/\s*?з)|(?:1[\d]{2}\s*?[уУ][кК]\s*?№\s*?1[,;: ].*т\s*?/\s*?з)|(?:1[\d]{2}\s*?[уУ][кК]\s*?№\s*?1)')
    # Заменить различные сокращения физкультурного зала на обобщённое значение
    df_prep['cabs'] = df_prep['cabs'].map(lambda val: re.sub(pat_zal, crd.DEF_SPORT_CAB, str(val)) if val else val)

    # Убрать всю информацию о необходимости сменной обуви
    df_prep['cabs'] = df_prep['cabs'].map(lambda val: re.sub(r'с[м]?.+?об', '', str(val)) if val else val)

    # Заменить символы переноса строки на стандартный разделитель
    df_prep['cabs'] = df_prep['cabs'].map(lambda val: re.sub(r'\n+', ' ; ', str(val)) if val else val)

    # Заменить все последовательности пробельных символов одним пробелом
    df_prep['cabs'] = df_prep['cabs'].map(lambda val: re.sub(r'\s+', ' ', str(val)) if val else val)

    """ Общая редакция (2) """
    # Финальное удаление полностью пустых строк
    df_prep = df_prep.dropna(axis=0, how='all')

    """ Задание имени датафрейма """

    # Имя датафрейма - имя группы
    df_prep.name = group
    
    return df_prep

