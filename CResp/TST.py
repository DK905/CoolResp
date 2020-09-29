""" Тестовый модуль для массовой конвертации

"""

""" Подключение модулей """
from modules import CR_reader  as crr   # Считывание таблицы в базу разбора для конкретной группы
from modules import CR_parser  as crp   # Парсинг базы разбора в нормальную БД для каждой логической записи
from modules import CR_analyze as cra   # Анализатор БД, для подправки косяков и определения числа подгрупп
from modules import CR_jsoner  as crj   # Сохранение БД парсинга и БД анализа в json (для оптимизации)
from modules import CR_writter as crw   # Форматная запись БД в таблицу EXCEL

# В какую папку сохранять
path_save = 'E:/CoolResp/Резы/'
path_load = 'E:/CoolResp/Тесты/'

# Для каких подгрупп составляется расписание
pdgr = [0, 0]


""" Раздел тестовых данных """
# Тесты с одной группой на листе
test1 = [f'{path_load}921_2_semestr_2019-2020.xls',       # Тест 0
         f'{path_load}MITE-91_2019-2020_II_semestr.xls',  # Тест 1
         f'{path_load}MIVT-91_2019-2020_II_semestr.xls',  # Тест 2
         f'{path_load}PE-51b_2018-2019_II_semestr.xlsx',  # Тест 3
         f'{path_load}PE-61b_2018-2019_II_semestr.xlsx',  # Тест 4
         f'{path_load}PE-71b_2019-2020_II_semestr.xlsx',  # Тест 5
         f'{path_load}PE-71_2_semestr_2018-2019.xls',     # Тест 6
         f'{path_load}PE-81_2_semestr_2019-2020.xls',     # Тест 7
         f'{path_load}PE-81b_2020-2021_I_semestr.xlsx'    # Тест 8
         ]

# Тесты с одним листом, но несколькими группами
test2 = [f'{path_load}1_KURS_2018-2019_2semestr.xls',     # Тест 0 # 1 лист, 3 группы
         f'{path_load}PE-91_92b_2_semestr_2019-2020.xls'  # Тест 1 # 1 лист, 2 группы
        ]

# Тесты с несколькими листами и группами
test3 = [f'{path_load}2_kurs_2semestr_2018-2019.xls',     # Тест 0 # 2 листа, 2 группы
         f'{path_load}3_kurs_2018-2019_II_semestr.xls',   # Тест 1 # 2 листа, 2 группы
         f'{path_load}3_kurs_2019-2020_II_semestr.xls',   # Тест 2 # 2 листа, 2 группы
         f'{path_load}4_kurs_2018-2019_II_semestr.xls',   # Тест 3 # 2 листа, 2 группы
         ]

from time import time
start_time = time()
print('start')

# Одиночный тест
book   = crr.read_book(test1[8])
sheets = crr.choise_sheet(book)
sheet  = book[sheets[0]]

sheet_info  = crr.choise_group(sheet)
timey_wimey = sheet_info[0] # Период расписания
year        = sheet_info[1] # Год
groups      = sheet_info[2] # Список групп на листе
row_start   = sheet_info[3] # Индекс стартовой строки

group = groups[0]
name =  f'{path_save}/Респа для {str(group)} на {year} год' + '.xlsx'
#print(name)
bd_process = crr.prepare(sheet, group, row_start, True)
bd_parse   = crp.parser(bd_process, year, True, False, True)
#crp.print_bd(bd_parse, group, timey_wimey, year)
analysis   = cra.analyze_bd(bd_parse)
crj.save_json(bd_parse, analysis, group)
#p_bd, a_bd = crj.load_json(f'Data/data_{group}.json')

#book       = crw.create_resp(bd_parse, analysis, pdgr[0], pdgr[1], timey_wimey, year)
#crw.save_resp(book, name)


"""
# Тесты с одной группой на листе
for i in range(8):
    book   = crr.read_book(test1[i])
    sheets = crr.choise_sheet(book)
    sheet  = book[sheets[0]]

    sheet_info  = crr.choise_group(sheet)
    timey_wimey = sheet_info[0] # Период расписания
    year        = sheet_info[1] # Год
    groups      = sheet_info[2] # Список групп на листе
    row_start   = sheet_info[3] # Индекс стартовой строки

    group = groups[0]
    name =  f'{path_save}/Респа для {str(group)} на {year} год' + '.xlsx'
    #print(name)
    bd_process = crr.prepare(sheet, group, row_start, True)
    bd_parse   = crp.parser(bd_process, year, True, False, True)
    #crp.print_bd(bd_parse, group, timey_wimey, year)
    analysis   = cra.analyze_bd(bd_parse)
    crj.save_json(bd_parse, analysis, group)
    p_bd, a_bd = crj.load_json(f'Data/data_{group}.json')

    book       = crw.create_resp(bd_parse, analysis, pdgr[0], pdgr[1], timey_wimey, year)
    crw.save_resp(book, name)

# Тесты с одним листом, но несколькими группами
for i in range(2):
    book   = crr.read_book(test2[i])
    sheets = crr.choise_sheet(book)
    sheet  = book[sheets[0]]

    sheet_info  = crr.choise_group(sheet)
    timey_wimey = sheet_info[0] # Период расписания
    year        = sheet_info[1] # Год
    groups      = sheet_info[2] # Список групп на листе
    row_start   = sheet_info[3] # Индекс стартовой строки

    group = groups[0]
    name =  f'{path_save}/Респа для {str(group)} на {year} год' + '.xlsx'
    #print(name)
    bd_process = crr.prepare(sheet, group, row_start, True)
    bd_parse   = crp.parser(bd_process, year, True, False, True)
    #crp.print_bd(bd_parse, group, timey_wimey, year)
    analysis   = cra.analyze_bd(bd_parse)
    book       = crw.create_resp(bd_parse, analysis, pdgr[0], pdgr[1], timey_wimey, year)
    crw.save_resp(book, name)

    group = groups[1]
    name =  f'{path_save}/Респа для {str(group)} на {year} год' + '.xlsx'
    #print(name)
    bd_process = crr.prepare(sheet, group, row_start, True)
    bd_parse   = crp.parser(bd_process, year, True, False, True)
    #crp.print_bd(bd_parse, group, timey_wimey, year)
    analysis   = cra.analyze_bd(bd_parse)
    book       = crw.create_resp(bd_parse, analysis, pdgr[0], pdgr[1], timey_wimey, year)
    crw.save_resp(book, name)

    if not i:
        group = groups[2]
        name =  f'{path_save}/Респа для {str(group)} на {year} год' + '.xlsx'
        #print(name)
        bd_process = crr.prepare(sheet, group, row_start, True)
        bd_parse   = crp.parser(bd_process, year, True, False, True)
        #crp.print_bd(bd_parse, group, timey_wimey, year)
        analysis   = cra.analyze_bd(bd_parse)
        book       = crw.create_resp(bd_parse, analysis, pdgr[0], pdgr[1], timey_wimey, year)
        crw.save_resp(book, name)


for i in range(4):
    book   = crr.read_book(test3[i])
    sheets = crr.choise_sheet(book)
    sheet  = book[sheets[0]]

    sheet_info  = crr.choise_group(sheet)
    timey_wimey = sheet_info[0] # Период расписания
    year        = sheet_info[1] # Год
    groups      = sheet_info[2] # Список групп на листе
    row_start   = sheet_info[3] # Индекс стартовой строки

    group = groups[0]
    name =  f'{path_save}/Респа для {str(group)} на {year} год' + '.xlsx'
    #print(name)
    bd_process = crr.prepare(sheet, group, row_start, True)
    bd_parse   = crp.parser(bd_process, year, True, False, True)
    #crp.print_bd(bd_parse, group, timey_wimey, year)
    analysis   = cra.analyze_bd(bd_parse)
    book       = crw.create_resp(bd_parse, analysis, pdgr[0], pdgr[1], timey_wimey, year)
    crw.save_resp(book, name)

    group = groups[1]
    name =  f'{path_save}/Респа для {str(group)} на {year} год' + '.xlsx'
    #print(name)
    bd_process = crr.prepare(sheet, group, row_start, True)
    bd_parse   = crp.parser(bd_process, year, True, False, True)
    #crp.print_bd(bd_parse, group, timey_wimey, year)
    analysis   = cra.analyze_bd(bd_parse)
    book       = crw.create_resp(bd_parse, analysis, pdgr[0], pdgr[1], timey_wimey, year)
    crw.save_resp(book, name)


for i in range(4):
    book   = crr.read_book(test3[i])
    sheets = crr.choise_sheet(book)
    sheet  = book[sheets[1]]

    sheet_info  = crr.choise_group(sheet)
    timey_wimey = sheet_info[0] # Период расписания
    year        = sheet_info[1] # Год
    groups      = sheet_info[2] # Список групп на листе
    row_start   = sheet_info[3] # Индекс стартовой строки

    group = groups[0]
    name =  f'{path_save}/Респа для {str(group)} на {year} год' + '.xlsx'
    #print(name)
    bd_process = crr.prepare(sheet, group, row_start, True)
    bd_parse   = crp.parser(bd_process, year, True, False, True)
    #crp.print_bd(bd_parse, group, timey_wimey, year)
    analysis   = cra.analyze_bd(bd_parse)
    book       = crw.create_resp(bd_parse, analysis, pdgr[0], pdgr[1], timey_wimey, year)
    crw.save_resp(book, name)

    group = groups[1]
    name =  f'{path_save}/Респа для {str(group)} на {year} год' + '.xlsx'
    #print(name)
    bd_process = crr.prepare(sheet, group, row_start, True)
    bd_parse   = crp.parser(bd_process, year, True, False, True)
    #crp.print_bd(bd_parse, group, timey_wimey, year)
    analysis   = cra.analyze_bd(bd_parse)
    book       = crw.create_resp(bd_parse, analysis, pdgr[0], pdgr[1], timey_wimey, year)
    crw.save_resp(book, name)
"""
print(f'Обработано за {time() - start_time} секунд')
