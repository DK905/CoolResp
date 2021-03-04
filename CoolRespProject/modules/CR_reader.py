""" Информация о модуле
Данный модуль предназначен для считывания изначального файла расписания (в формате xls или xlsx),
и последующего вычленения из него расписания конкретной выбранной группы.

Первоначальное считывание идёт через pyexcel, так как данный модуль лучше всего справляется с данной задачей.
Можно было использовать openpyxl, но он поддерживает только xlsx.
Можно было использовать pandas,   но это слишком тяжёлая артиллерия для такого простого действия. 

"""

# Импорт функции для считывания EXCEL таблицы
from pyexcel import get_book_dict as pxl_book

# Импорт базовых команд поиска и замены из стандартного модуля регулярных выражений
from re import findall, fullmatch, match, search, split as rsplit, sub

# Импорт умолчаний
from CoolRespProject.modules.CR_dataset import days_names, defaults


def read_book(file: 'Путь к файлу таблицы расписания'
              ) ->  'Объект книги в формате типизированного словаря':

    """ функция для считывания книги xls/xlsx в специальный словарь из pyexcel """
    # Импорты для компиляции (у pyexcel есть проблемы с pyinstaller)
    import pyexcel_xls, pyexcel_xlsx

    # Если выбран файл с расширением .xls/.xlsx, то попробовать открыть его
    if search(r'(?m)\.xlsx?$', file):
        return pxl_book(file_name = file)


def choise_sheet(d_book: 'Типизированный словарь из read_book'
                 ) ->    'Список листов книги (для меню выбора листов)':

    """ Функция для отображения листов в книге """
    # Если в книге есть листы, вернуть их список
    if list(d_book.keys()):
        return(list(d_book.keys()))


def choise_group(sheet: 'Типизированный словарь из read_book'
                 ) ->   'Список [Период, год, список групп на листе, индекс начальной строки расписания]':

    """ Функция для выделения нужной информации из расписания """
    timey_wimey = year = grp_list = start = None
    for r_ind, row in enumerate(sheet):
        # Найти период расписания, если ещё не найден
        if not timey_wimey:
            for col in row:
                if search(r'[Нн]а\s*период', col):
                    # Выделение периода расписания. Если он указан неправильно, потом будут проблемы
                    # timey_wimey = findall(r'с\s*([\d.]{5,8})[г.]*\s*по\s*([\d.]{5,8})[г.]*', col)[0]
                    timey_wimey = findall(r'с\s*([\d.]+)[г.]*\s*по\s*([\d.]+)[г.]*', col)[0]
                    # Период приводится к самому частому формату по типу "04.03.2021"
                    timey_wimey = ['.'.join([part[-2:] for part in date.split('.')]) for date in timey_wimey]
                    # Выделение года
                    year = 2000 + int(timey_wimey[0][-2:])
                    break
        # Найти список групп и начальную позицию
        if not grp_list and search(r'(?m)^[Дд]ни', row[0]):
            # Отсеивание пустых столбцов
            grp_list = [el for el in row if el]
            grp_list = [grp_list[i] for i in range(2, len(grp_list), 2)]
            start = r_ind
            break
        # Если всё найдено, продолжать поиск не нужно
        if timey_wimey and grp_list:
            break
    # Если на листе есть список групп и период расписания, вернуть инфу
    if timey_wimey and grp_list:
        return [timey_wimey, year, grp_list, start]


"""                 Начальная стадия разбора расписания

Данная стадия предполагает подготовку расписания выбранной группы к препарированию.
Сначала, среди всех листов и всех групп выбирается одна конкретная группа, для которой
вычленяется собственная таблица расписания, которая в процессе частично форматируется.
По итогу стадии, будет создана первичная БД обработки

"""


def merged_cells(row: 'Текущая строка таблицы',
                 col: 'Текущий столбец таблицы'
                 ) -> 'Значение объединённой ячейки':

    """ Функция корректного считывания объединённых ячеек """
    act = col - 2
    # Если есть предмет и правый сосед не кабинет - ячейка общая
    if row[act] and not row[act+1]:
        return row[act]
    # Если ячейка 100% не общая или строка пройдена - вернуть пустую запись
    elif row[act+1] or act == 2 and not row[act]:
        return ''
    else:
        return merged_cells(row, act)


def exterminate(sheet: 'Лист, на котором удаляются пустые столбцы'
                ) ->   'Лист без пустых столбцов':

    """ Функция удаления пустых столбцов на листе """
    sheet = list(zip(*sheet))
    sheet = [rs for rs in sheet if any(rs)]
    sheet = [list(row) for row in zip(*sheet)]
    return sheet


def what_col(title: 'Шапка подтаблицы расписания',
             group: 'Выбранная группа'
             ) ->   'Индекс столбца группы':

    """ Функция определения столбца группы """
    for ind, rec in enumerate(title):
        if str(rec) == str(group):
            return ind


def repl_a(cab: 'Строка с кабинетами'
           ) -> 'Строка с единой записью кабинетов':

    """ Первичное форматирование списка кабинетов """
    # Единый формат для актового зала
    cab = sub(r'[аА].*?[лЛ]', defaults[3], cab)
    # Кабинет для ФЗК может разнообразно мимикрировать: "с/з, т/з", "1xx УК№1 с/з, т/з", "1xx УК№1"
    pat_zal = r'(?:с\s*?/\s*?з.*?т\s*?/\s*?з)|(?:1[\d]{2}\s*?[уУ][кК]\s*?№\s*?1[,;: ].*т\s*?/\s*?з)|(?:1[\d]{2}\s*?[уУ][кК]\s*?№\s*?1)'
    cab = sub(pat_zal, defaults[4], cab)
    cab = sub(r'с.+?об', '', cab)
    cab = sub(r'\n+', ' ; ', cab)
    cab = sub(r'\s+', ' ', cab)
    return cab


def repl_b(cab:  'Подстрока с одним кабинетом',
           sepr: 'Флаг сокращения кабинетов'
           ) ->  'Отформатированная подстрока кабинета':

    """ Форматирование записи корпусов кабинетов """
    cab = sub(r'[.;: ]', '', cab)
    # Если кабинеты записываются в полной форме
    if not sepr:
        cab = sub(r'[уУ][кК]?№', ' УК№', cab)
    else:
        cab = sub(r'\s*[уУ][кК]?№\d', '', cab)
    return cab


def swap_quiz(inf_cell: 'Ячейка с инфой о предметах'
              ) ->      'Ячейка с инфой о предметах (зачёты/дифы заменены на спец. коды)':

    """ Кодирование зачётов и дифов """
    pattern_dif = r'(?i)(?:зач[её]т[\s]*с[\s]*оценкой)|(?:диф[.\s]*зач[её]т)' # Отлов дифов
    pattern_zac = r'(?i)зач[её]т'  # Отлов зачётов
    inf_cell = sub(r'\s+', ' ', inf_cell)  # Пробельная чистка
    if search(pattern_dif, inf_cell):  # Замена дифов на 6D6D6D6
        inf_cell = sub(pattern_dif, '6D6D6D6', inf_cell)
    if search(pattern_zac, inf_cell):  # Замена зачётов на 7Z7Z7Z7
        inf_cell = sub(pattern_zac, '7Z7Z7Z7', inf_cell)
    return inf_cell


def prepare(sheet: 'Выбранный лист',
            group: 'Выбранная группа',
            start: 'Индекс начальной строки расписания',
            dv_yn: 'Флаг сокращения кабинетов'
            ) ->   'Урезка таблицы, подготовленная к парсингу':

    """ Функция для подготовки расписания выбранной подгруппы к препарированию """
    trash = []  # БД разбора

    end = 0
    # Расписание всегда идёт вплоть до "начальник УО"
    for i in range(start+1, len(sheet)+1):
        finita = [a for a in sheet[i] if search(r'(?i)начальник\s*уо', str(a))]
        if finita:
            end = i-1
            break

    sheet = sheet[start:end]  # Первичная БД

    if '' in sheet[0]:              # Если в шапке есть пустые значения, то в таблице могут быть пустые столбцы...
        sheet = exterminate(sheet)  # ...а пустые столбцы = штука, которая усложняет обработку
    g = what_col(sheet[0], group)   # Определение столбца с инфой для выбранной группы
    sheet = sheet[1:]               # Обрезка расписательной таблицы по шапке

    # Регулярный шаблон для выбора дня недели (на случай, если где-то его нет)
    days_pat = r'(?i)(.*л.*)|(.*в.*о.*к.*)|(.*с.*д.*)|(.*ч.*г.*)|(.*я.*)|(.*у.*б.*)'   

    for row in sheet:
        """ Выделение конкретных частей новой записи в БД """
        # День недели
        if not row[0]:  # День - общая ячейка, то есть значение есть только в левой верхней ячейке
            day = trash[-1][0]  # Следовательно, для общей ячейки, день берётся из предыдущей записи
        else:  # Но если день есть, то его нужно определить и привести к общему виду
            for ind_d, sovp in enumerate(findall(days_pat, row[0])[0]):
                if sovp:
                    day = days_names[ind_d]
                    break

        # Номер пары
        num = row[1]

        # Ячейка с инфой о паре
        if row[g] or row[g-1]:  # Если инфа лежит не где-то в левой части объединённой ячейки
            info = swap_quiz(row[g])
        elif g > 2:
            info = swap_quiz(merged_cells(row, g))
        else:
            info = ''

        # Кабинеты
        if row[g+1]:
            cabs = repl_a(row[g+1])
            cabs = [repl_b(cb, dv_yn) for cb in rsplit(r'[;,]', cabs)
                    if not fullmatch(r'[.,:; ]*', cb)]
            if not dv_yn:
                # Исправление ['210', '212 УК№1', '329', '331 УК№5', '410 УК№1'] в полных записях
                for ind_cb in range(len(cabs)-1, -1, -1):
                    # Если умолчание, то можно пропустить кабинет
                    if cabs[ind_cb] == defaults[3] or cabs[ind_cb] == defaults[4]:
                        continue
                    if match(r'\d+', cabs[ind_cb]) and not match(r'\d+ УК№\d', cabs[ind_cb]):
                        if search(r' УК№\d', cabs[ind_cb+1]):
                            cabs[ind_cb] += search(r' УК№\d', cabs[ind_cb+1])[0]
                        else:
                            cabs[ind_cb] += ' УК№?'
        else:
            cabs = []

        # Запись в БД
        trash.append([day, num, info, cabs])

    # Если в процессе считывания не было косяков, то вернуть базу разбора
    return trash
