"""
Основной модуль, из которого вызываются все остальные части
До окончания ревью, без интерфейса

"""
""" Подключение модулей """
import CR_reader  as crr # Считывание таблицы в базу разбора для конкретной группы
import CR_parser  as crp # Парсинг базы разбора в нормальную БД для каждой логической записи
import CR_analyze as cra # Анализатор БД, для подправки косяков и определения числа подгрупп
import CR_writter as crw # Форматная запись БД в таблицу EXCEL

""" Раздел констант """
# Сокращённая запись дней
days_names = {0: 'ПН', 1: 'ВТ', 2: 'СР', 3: 'ЧТ', 4: 'ПТ', 5: 'СБ'}

# Сокращённая запись типов пары
tip_list = {0: 'ДИФ', 1: 'зачёт', 2: 'лекция', 3: 'лекция',  4: 'ЛБ', 5: 'ПР'}

# В какую папку сохранять
path = 'F:/CoolResp/Резы/'

# Для каких подгрупп составляется расписание
pdgr = [0, 0]

# Обозначения для умолчаний
defaults = ['общ',      # 0 # Обозначение для общей пары
            'рейд',     # 1 # Обозначение для предмета без типа пары и подгруппы
            'ЛБ/ПР',    # 2 # Обозначение для предмета без типа пары, но с подгруппой
            '',         # 3 # Обозначение для даты по умолчанию. Задаётся периодом расписания
            '',         # 4 # Обозначение для года по умолчанию. Вычленяется из периода расписания
            'АКТ_зал',  # 5 # Обозначение для кабинета по умолчанию
            #           # 6 # Препод по умолчанию
            'Преображенский Ф.Ф.'
           ]


""" Раздел тестовых данных """
# Тесты с одной группой на листе
test1 = ['F:/CoolResp/Тесты/921_2_semestr_2019-2020.xls',       # Тест 0
         'F:/CoolResp/Тесты/MITE-91_2019-2020_II_semestr.xls',  # Тест 1
         'F:/CoolResp/Тесты/MIVT-91_2019-2020_II_semestr.xls',  # Тест 2
         'F:/CoolResp/Тесты/PE-51b_2018-2019_II_semestr.xlsx',  # Тест 3
         'F:/CoolResp/Тесты/PE-61b_2018-2019_II_semestr.xlsx',  # Тест 4
         'F:/CoolResp/Тесты/PE-71b_2019-2020_II_semestr.xlsx',  # Тест 5
         'F:/CoolResp/Тесты/PE-71_2_semestr_2018-2019.xls',     # Тест 6
         'F:/CoolResp/Тесты/PE-81_2_semestr_2019-2020.xls'      # Тест 7
         ]

# Тесты с одним листом, но несколькими группами
test2 = ['F:/CoolResp/Тесты/1_KURS_2018-2019_2semestr.xls',     # Тест 0 # 1 лист, 3 группы
         'F:/CoolResp/Тесты/PE-91_92b_2_semestr_2019-2020.xls'  # Тест 1 # 1 лист, 2 группы
        ]

# Тесты с несколькими листами и группами
test3 = ['F:/CoolResp/Тесты/2_kurs_2semestr_2018-2019.xls',     # Тест 0 # 2 листа, 2 группы
         'F:/CoolResp/Тесты/3_kurs_2018-2019_II_semestr.xls',   # Тест 1 # 2 листа, 2 группы
         'F:/CoolResp/Тесты/3_kurs_2019-2020_II_semestr.xls',   # Тест 2 # 2 листа, 2 группы
         'F:/CoolResp/Тесты/4_kurs_2018-2019_II_semestr.xls',   # Тест 3 # 2 листа, 2 группы
         ]

# Тесты с одной группой на листе
for i in range(8):
    print(test1[i])
    book   = crr.read_book(test1[i])
    sheets = crr.choise_sheet(book)
    sheet  = book[sheets[0]]

    sheet_info  = crr.choise_group(sheet)
    timey_wimey = sheet_info[0] # Период расписания
    year        = sheet_info[1] # Год
    defaults[3], defaults[4] = timey_wimey, year
    groups      = sheet_info[2] # Список групп на листе
    row_start   = sheet_info[3] # Индекс стартовой строки

    group = groups[0]
    name =  f'{path}/Респа для {str(groups[0])} на {year} год' + '.xlsx'
    bd_process = crr.prepare(sheet, group, row_start, days_names, True)
    bd_parse   = crp.parser(bd_process, defaults, tip_list, True, True)
    # crp.print_bd(bd_parse, group, timey_wimey, year)
    analysis   = cra.analyze_bd(bd_parse, defaults[0])
    crw.create_resp(bd_parse, analysis, pdgr, name, days_names, timey_wimey, year)

# Тесты с одним листом, но несколькими группами
for i in range(2):
    print(test2[i])
    book   = crr.read_book(test2[i])
    sheets = crr.choise_sheet(book)
    sheet  = book[sheets[0]]

    sheet_info  = crr.choise_group(sheet)
    timey_wimey = sheet_info[0] # Период расписания
    year        = sheet_info[1] # Год
    defaults[3], defaults[4] = timey_wimey, year
    groups      = sheet_info[2] # Список групп на листе
    row_start   = sheet_info[3] # Индекс стартовой строки

    group = groups[0]
    name =  f'{path}/Респа для {str(groups[0])} на {year} год' + '.xlsx'
    bd_process = crr.prepare(sheet, group, row_start, days_names, True)
    bd_parse   = crp.parser(bd_process, defaults, tip_list, True, True)
    # crp.print_bd(bd_parse, group, timey_wimey, year)
    analysis   = cra.analyze_bd(bd_parse, defaults[0])
    crw.create_resp(bd_parse, analysis, pdgr, name, days_names, timey_wimey, year)

    group = groups[1]
    name =  f'{path}/Респа для {str(groups[0])} на {year} год' + '.xlsx'
    bd_process = crr.prepare(sheet, group, row_start, days_names, True)
    bd_parse   = crp.parser(bd_process, defaults, tip_list, True, True)
    # crp.print_bd(bd_parse, group, timey_wimey, year)
    analysis   = cra.analyze_bd(bd_parse, defaults[0])
    crw.create_resp(bd_parse, analysis, pdgr, name, days_names, timey_wimey, year)

    if not i:
        group = groups[2]
        name =  f'{path}/Респа для {str(groups[0])} на {year} год' + '.xlsx'
        bd_process = crr.prepare(sheet, group, row_start, days_names, True)
        bd_parse   = crp.parser(bd_process, defaults, tip_list, True, True)
        # crp.print_bd(bd_parse, group, timey_wimey, year)
        analysis   = cra.analyze_bd(bd_parse, defaults[0])
        crw.create_resp(bd_parse, analysis, pdgr, name, days_names, timey_wimey, year)


for i in range(4):
    print(test3[i])
    book   = crr.read_book(test3[i])
    sheets = crr.choise_sheet(book)
    sheet  = book[sheets[0]]

    sheet_info  = crr.choise_group(sheet)
    timey_wimey = sheet_info[0] # Период расписания
    year        = sheet_info[1] # Год
    defaults[3], defaults[4] = timey_wimey, year
    groups      = sheet_info[2] # Список групп на листе
    row_start   = sheet_info[3] # Индекс стартовой строки

    group = groups[0]
    name =  f'{path}/Респа для {str(groups[0])} на {year} год' + '.xlsx'
    bd_process = crr.prepare(sheet, group, row_start, days_names, True)
    bd_parse   = crp.parser(bd_process, defaults, tip_list, True, True)
    # crp.print_bd(bd_parse, group, timey_wimey, year)
    analysis   = cra.analyze_bd(bd_parse, defaults[0])
    crw.create_resp(bd_parse, analysis, pdgr, name, days_names, timey_wimey, year)

    group = groups[1]
    name =  f'{path}/Респа для {str(groups[0])} на {year} год' + '.xlsx'
    bd_process = crr.prepare(sheet, group, row_start, days_names, True)
    bd_parse   = crp.parser(bd_process, defaults, tip_list, True, True)
    # crp.print_bd(bd_parse, group, timey_wimey, year)
    analysis   = cra.analyze_bd(bd_parse, defaults[0])
    crw.create_resp(bd_parse, analysis, pdgr, name, days_names, timey_wimey, year)


for i in range(4):
    print(test3[i])
    book   = crr.read_book(test3[i])
    sheets = crr.choise_sheet(book)
    sheet  = book[sheets[1]]

    sheet_info  = crr.choise_group(sheet)
    timey_wimey = sheet_info[0] # Период расписания
    year        = sheet_info[1] # Год
    defaults[3], defaults[4] = timey_wimey, year
    groups      = sheet_info[2] # Список групп на листе
    row_start   = sheet_info[3] # Индекс стартовой строки

    group = groups[0]
    name =  f'{path}/Респа для {str(groups[0])} на {year} год' + '.xlsx'
    bd_process = crr.prepare(sheet, group, row_start, days_names, True)
    bd_parse   = crp.parser(bd_process, defaults, tip_list, True, True)
    # crp.print_bd(bd_parse, group, timey_wimey, year)
    analysis   = cra.analyze_bd(bd_parse, defaults[0])
    crw.create_resp(bd_parse, analysis, pdgr, name, days_names, timey_wimey, year)

    group = groups[1]
    name =  f'{path}/Респа для {str(groups[0])} на {year} год' + '.xlsx'
    bd_process = crr.prepare(sheet, group, row_start, days_names, True)
    bd_parse   = crp.parser(bd_process, defaults, tip_list, True, True)
    # crp.print_bd(bd_parse, group, timey_wimey, year)
    analysis   = cra.analyze_bd(bd_parse, defaults[0])
    crw.create_resp(bd_parse, analysis, pdgr, name, days_names, timey_wimey, year)
