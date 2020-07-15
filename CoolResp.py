# Коробка с костылями, выкованная дворфами из грязнейшего говнокода и зачарованная эльфийскими регулярными выражениями

import pyexcel  # Модуль для обработки EXCEL таблиц
import re       # Модуль регулярных выражений
import datetime # Модуль обработки дат
import openpyxl # Модуль для сохранения EXCEL таблицы в .xlsx (с форматированием)

""" Считывание файла в книжный словарь """
# Тесты с несколькими группами и листами
#temp_table = pyexcel.get_book_dict(file_name='F:/CResp_Tests/Тесты/1_KURS_2018-2019_2semestr.xls')     # 1 Интересный тест
#temp_table = pyexcel.get_book_dict(file_name='F:/CResp_Tests/Тесты/2_kurs_2semestr_2018-2019.xls')     # 2  
#temp_table = pyexcel.get_book_dict(file_name='F:/CResp_Tests/Тесты/3_kurs_2018-2019_II_semestr.xls')   # 3  
#temp_table = pyexcel.get_book_dict(file_name='F:/CResp_Tests/Тесты/3_kurs_2019-2020_II_semestr.xls')   # 4  
#temp_table = pyexcel.get_book_dict(file_name='F:/CResp_Tests/Тесты/4_kurs_2018-2019_II_semestr.xls')   # 5  
#temp_table = pyexcel.get_book_dict(file_name='F:/CResp_Tests/Тесты/PE-91_92b_2_semestr_2019-2020.xls') # 6  

# Тесты где всего один лист и всего одна группа
#temp_table = pyexcel.get_book_dict(file_name='F:/CResp_Tests/Тесты/921_2_semestr_2019-2020.xls')       # 7  Интересный тест 
#temp_table = pyexcel.get_book_dict(file_name='F:/CResp_Tests/Тесты/MITE-91_2019-2020_II_semestr.xls')  # 8  
#temp_table = pyexcel.get_book_dict(file_name='F:/CResp_Tests/Тесты/MIVT-91_2019-2020_II_semestr.xls')  # 9  
#temp_table = pyexcel.get_book_dict(file_name='F:/CResp_Tests/Тесты/PE-51b_2018-2019_II_semestr.xlsx')  # 10 
#temp_table = pyexcel.get_book_dict(file_name='F:/CResp_Tests/Тесты/PE-61b_2018-2019_II_semestr.xlsx')  # 11 Интересный тест 
#temp_table = pyexcel.get_book_dict(file_name='F:/CResp_Tests/Тесты/PE-71b_2019-2020_II_semestr.xlsx')  # 12 
#temp_table = pyexcel.get_book_dict(file_name='F:/CResp_Tests/Тесты/PE-71_2_semestr_2018-2019.xls')     # 13 Белкина
#temp_table = pyexcel.get_book_dict(file_name='F:/CResp_Tests/Тесты/PE-81_2_semestr_2018-2019.xls')     # 14 
temp_table = pyexcel.get_book_dict(file_name='F:/CResp_Tests/Тесты/PE-81_2_semestr_2019-2020.xls')     # 15 Интересный тест

""" Выбор листа с книги """
if len(temp_table) > 1:
    # Если листов несколько, то нужно выбрать один из них
    print(*temp_table.keys(), sep=' | ')
    #sheet = input('Выберите лист: ')
    #sheet = 'МЕ,ОЕ 71б' # Для 2
    #sheet = 'ИТ,ВЕ-71б'
    #sheet = 'МЕ,ОЕ-61б' # Для 3
    #sheet = 'ИТ,ВЕ-61б'
    #sheet = 'МЕ,ОЕ-71б' # Для 4
    #sheet = 'ИТ,ВЕ-71б'
    #sheet = 'МЕ,ОЕ-51б' # Для 5
    #sheet = 'ИТ,ВЕ-51б'
else:
    # Если лист один, то автовыбор его
    for i in temp_table:
        sheet = i

""" Выделение позиции графика расписания на табличном листе
Принцип таков:
1) Корректное расписание всегда начинается через одну строку после указания его периода
2) Корректное расписание всегда заканчивается упоминанием должности "Начальник УО"
3) Т.е, расписание - всё что внутри 1-2, но начиная с 3-й строки
"""
for i, a in enumerate(temp_table[sheet]):
    if re.search(' *[Нн]а *период', a[0], re.A):
        start = i
    if re.search(' *[Нн]ачальник *УО', a[0], re.A):
        end = i

""" Загрузить график расписания групп на листе, и выгрузить книгу EXCEL из памяти """
temp_sheet = temp_table[sheet][start:end]
del(temp_table)

""" Удаление пустых столбцов (а они иногда встречаются) """
if '' in temp_sheet[0]:
    temp_sheet = list(zip(*temp_sheet))
    temp_sheet = [rs for rs in temp_sheet if any(rs)]
    temp_sheet = [list(row) for row in zip(*temp_sheet)]

""" Обособление периода расписания и групп """
# Период выделяется из "с <дата1> по <дата2>" в кортеж (дата1, дата2)
timey_wimey = re.findall(r'(с\s*[\d.]{5,8}[г.]*\s*по\s*[\d.]{5,8}[г.]*)', temp_sheet[0][0])[0]
period = re.findall(r'с\s*([\d.]{5,8})[г.]*\s*по\s*([\d.]{5,8})[г.]*', temp_sheet[0][0])[0]
year = str(2000+int(period[0][-2:]))

# Строка групп сохраняется как список групп на листе
#title = [temp_sheet[1][i] for i in range(2, len(temp_sheet[1]), 2)]
#print(f'Список групп на листе: {title}')
title = temp_sheet[1] # Строка групп

# Из расписания удаляются строки периода и списка групп
temp_sheet = temp_sheet[2:]


""" Процедура корректного считывания объединённых ячеек """
def merged_cells(row, col, cabs):
    act = col-2
    # Если есть предмет и правый сосед не кабинет - ячейка общая
    if row[act] and not row[act+1]:
        return [row[0], row[1], row[act], cabs]
    # Если ячейка 100% не общая или строка пройдена - вернуть пустую запись
    elif row[act+1] or act == 2 and not row[act]:
        return[row[0], '', '', '']
    else:
        return merged_cells(row, act, cabs)

""" Процедура отсеивания ошибок в ячейках (начальная стадия разбора) """
def errors_clear(sheet, i, row):
    if not row[0]: # Т.к день - общая ячейка
        row[0] = sheet[i-1][0]
    else: # Если день есть, привести к общему виду
        row[0] = re.sub('\s+', '', row[0])
        row[0] = row[0][0].upper() + row[0][1:]
    if not row[1]: # Один номер пары порой включает несколько строк
        row[1] = sheet[i-1][1]
    # Чистка кабинетов и удаление непонятной инфы
    pattern = r'(?:[ст]/з[.]{0,2})|(?:см.об[.]{0,2})'
    if row[-1]:
        cabs = row[-1]
        cabs = re.sub(pattern, '', re.sub(r'\s', ' ', cabs))
        cabs = re.split(r'[,;]', cabs)
        cabs = [re.sub(r'\d[ ]*У', lambda m: m[0][0]+' '+'У', re.sub(' ', '', j))
                   for j in cabs if not re.fullmatch(r'[ ,.;]*', j)]
        if cabs:
            row[-1] = cabs
        else: # Если была только непонятная инфа, то кабинет итак все знают
            row[-1] = ['Знамогде']
    return row

""" Процедура замены зачётов/дифов (мешают обработке) """
def swap_quiz(inf_data):
    pattern_dif = r'(?:зач[её]т[\s]*с[\s]*оценкой)|(?:диф[.\s]*зач[её]т)' # Отлов диф.зачётов
    pattern_zac = r'зач[её]т' # Отлов зачётов
    inf_data = re.sub(r'\s+', ' ', inf_data) # Пробельная чистка
    if re.search(pattern_dif, inf_data, flags = re.I): # Замена дифов на 6D6D6D6
        inf_data = re.sub(pattern_dif, '6D6D6D6', inf_data, flags = re.I)
    if re.search(pattern_zac, table[-1][2], flags = re.I): # Замена зачётов на 7Z7Z7Z7
        inf_data = re.sub(pattern_zac, '7Z7Z7Z7', inf_data, flags = re.I)
    return inf_data


""" Создание базы обработки (table) для одной из групп на листе
По сути, это контейнер с мусором, из которого хлам отправится на дальнейшую переработку
* Структура элемента формируемой базы обработки следующая:
    [0] - день
    [1] - номер пары
    [2] - ячейка с инфой по предметам, преподам, группам, и т.п
    [3] - список кабинетов
"""
# Для наглядности доп. условий, база обработки создаётся циклом а не через List Comprehension
table = []
# Таблицы сравнительно малы => можно хранить пустые строки (нужны для корректных дней, № пар)
if len(title) > 4: # Если группа одна, титул состоит из 4-х ячеек. Иначе - 4+
    # Выделение списка групп на листе
    g = {title[a]:a for a in range(2, len(title), 2)}
    #print(*g.keys(), sep=' | ')
    #v = g[input('Выберите группу: ')]
    v = g['МЕ-81б'] # Для 1
    #v = g['ИТ-81б']
    #v = g['ОЕ-81б']
    #v = g['МЕ-71б'] # Для 2
    #v = g['ОЕ-71б']
    #v = g['ИТ-71б']
    #v = g['ВЕ-71б']
    #v = g['МЕ-61б'] # Для 3
    #v = g['ОЕ-61б']
    #v = g['ИТ-61б']
    #v = g['ВЕ-61б']
    #v = g['МЕ-71б'] # Для 4
    #v = g['ОЕ-71б']
    #v = g['ИТ-71б']
    #v = g['ВЕ-71б']
    #v = g['МЕ-51б'] # Для 5
    #v = g['ОЕ-51б']
    #v = g['ИТ-51б']
    #v = g['ВЕ-51б']
    #v = g['ПЕ-91б'] # Для 6
    #v = g['ПЕ-92б']
    # Каждая строка расписания добавляется на разбор в формате [День, №, инф_ячейка, кабинеты]
    for ind, record in enumerate(temp_sheet):
        record = errors_clear(temp_sheet, ind, record)
        # Когда групп несколько, встречаются потоковые лекции (общие ячейки)
        if record[v] or record[v-1]:
            table.append([record[0], record[1], record[v], record[v+1]])
        else:
            table.append(merged_cells(record, v, record[v+1]))
        if table[-1][2]: # Стандарт пробельных символов в информации
            table[-1][2] = swap_quiz(table[-1][2])
else: # Если группа всего одна
    v = 2 # Столбец предмета в пайтоновской нумерации
    for ind, record in enumerate(temp_sheet):
        record = errors_clear(temp_sheet, ind, record)
        # В расписании одной группы нет общих ячеек с информацией о предмете
        table.append([record[0], record[1], record[v], record[v+1]])
        if table[-1][2]: # Стандарт пробельных символов в информации
            table[-1][2] = swap_quiz(table[-1][2])


""" Процедуры форматирования элементов записи """
# Форматирование преподов
def format_prep(prepod):
    repair_pat1 = r'^([а-я. \d]+)\s'
    repair_pat2 = r'([А-Я][а-я]+)\s*([А-Я.]{4})'
    if not re.search(r'\.$', prepod, flags = re.M):
        prepod += '.'
    prep = ['', re.search(repair_pat2, prepod)[1], re.search(repair_pat2, prepod)[2]]
    if re.search(repair_pat1, prepod, flags = re.M):
        prep[0] = re.search(repair_pat1, prepod, flags = re.M)[1]
        prep[0] = re.sub(r'\s+', ' ', prep[0])
        prep[0] = re.sub(r'\.\s+', '.', prep[0])
    else:
        del(prep[0])
    return ' '.join(prep)

# Форматирование типа пары
def format_tip(tip):
    tip_list = {0: 'Диф.Зачёт', 1: 'Зачёт', 2: 'Лекция', 3: 'Лекция',  4: 'Лаба', 5: 'Практика'}
    repair_pat = r'(.*?6.*D.*)|(.*?7.*Z.*)|(.*?т.*р.*я.*)|(.*?л.*к.*я.*)|(.*?л.*б.*р.*)|(.*?п.*а.*к.*)'
    for i, sovp in enumerate(re.findall(repair_pat, tip)[0]):
        if sovp:
            return tip_list[i]

# Форматирование типа подгруппы
def format_group(group):
    repair_pat = r'(\d)'
    return re.findall(repair_pat, group)[0] + 'п/гр'

# Абсолютное разбитие даты
def expand_dates(dates, year, day=7):
    dates = '; '.join(dates)
    any_dates = r'(?:с\s*[\d.]{5,8}[г.]*\s*по\s*[\d.]{5,8}[г.]*)|(?:[\d,]+[ .][\d]{2})'
    dates = re.findall(any_dates, dates)
    repair_pat1 = r'с\s*([\d.]{5,8})[г.]*\s*по\s*([\d.]{5,8})[г.]*' # Дата вида "с..по.."
    repair_pat2 = r'([\d,]+)[ .]([\d]{2})' # Дата вида "день, день,...,месяц
    all_dates = []
    for date in dates:
        if re.search(repair_pat1, date): # Пример входных данных: 'с 13.01 по 06.06'
            start_end = re.search(repair_pat1, date).groups()
            dt_start = list(map(int, start_end[0].split('.')))
            dt_final = list(map(int, start_end[1].split('.')))
            dt_start = datetime.date(int(year), dt_start[1], dt_start[0])
            if day < 7: # Если периодизация для дня недели, нужно найти корректный старт
                sdvig = 7 # Сдвиг - неделя при любом дне
                while dt_start.weekday() != day:
                    dt_start += datetime.timedelta(days=1)
            else: # Сдвиг - день при заполнении календарной шкалы
                sdvig = 1
            dt_final = datetime.date(int(year), dt_final[1], dt_final[0])
            while dt_start <= dt_final:
                all_dates.append(dt_start.strftime('%d.%m.%Y'))
                dt_start += datetime.timedelta(days=sdvig)
                # На случай, если заполняется календарная шкала
                if dt_start.weekday() == 6:
                    dt_start += datetime.timedelta(days=1)
        else: # Пример входных данных: '03,17,24,31.03'
            date = re.findall(repair_pat2, date)[0]
            days = date[0].split(',')
            month = date[1]
            all_dates.extend('.'.join([day, month, year]) for day in days)
    return all_dates

""" Регулярные шаблоны для переработки информации """
pat1 = r'([А-ЯЁA-Z][а-яaёa-zА-ЯЁA-Z, ()-]+)(?:[:;\n\d]|$)' # Отлов предмета (зачёты и дифы исключаются до отлова)
pat2 = r'(?:(?:с\s*(?:\d{2}[.,г;\s]*\s*)+по\s*(?:\d{2}(?:[.,г;\s]|$)\s*)+))|(?:(?:\d{2}(?:[.,г;\s]|$)\s*)+)' # Даты
pat4 = r'(?:[а-яёa-z]{2,}[а-яёa-z. \d]+\s{1,}[А-ЯЁA-Z][а-яёa-z ]+(?:[А-ЯЁA-Z][.]?){2})' # Отлов препода и должности
pat5 = r'[А-ЯЁA-Z][а-яёa-z ]+(?:[А-ЯЁA-Z][.]?){2}' # Отлов только ФИО препода (должность иногда опускают)
pat6 = r'(\d\s*[п]?\s*/\s*гр)' # Отлов подгрупп
pat7 = r'(?:[практика]{8})|(?:[лаб.раб ]{8})|(?:[лекция]{6})|(?:[теория]{6})|6D6D6D6|7Z7Z7Z7' # Отлов типа пары

# Разделение по предметам
pattern1 = r'(?:.*?(?:[А-ЯЁA-Z][а-яёa-zА-ЯЁA-Z, ()-]+(?:[:;\n\d]|$)).*?)(?=(?:[А-ЯЁA-Z][а-яёa-zА-ЯЁA-Z, ()-]+[:;\n\d])|(?:$))'
# Разделение по преподам
pattern2a = r'.*?(?:[а-яёa-z.]{2,}\s[А-ЯЁA-Z][а-яёa-z\s]+(?:[А-ЯЁA-Z][.]?){2})'
# Разделение по преподам: выделение подгрупп и преподов (если после инфы просто перечисление подгрупп-преподов)
pattern2b = r'(\d\s*[п]?\s*/\s*гр).*?((?:[а-яёa-z]{2,}[а-яёa-z. \d]+\s{1,})?[А-ЯЁA-Z][а-яёa-z ]+(?:[А-ЯЁA-Z][.]?){2})'
# Отлов даты в конце строки (для конкретных случаев, когда косяк в захвате даты)
pattern3 = r'(?:(?:(?:с\s*(?:\d{2}[.,г;\s]*\s*)+по\s*(?:\d{2}[.,г;\s]{0,2}\s*)+))|(?:(?:\d{2}[.,г;\s]{0,2}\s*)+))$'
# Отлов дат, типов пары и групп как [даты, типы пары, группы]
pattern4 = r'((?:(?:с\s*(?:\d{2}[.,г;\s]*\s*)+по\s*(?:\d{2}(?:[.,г;\s]|$)\s*)+))|(?:(?:\d{2}(?:[.,г;\s]|$)\s*)+))|((?:[практика]{8})|(?:[лаб.раб ]{8})|(?:[лекция]{6})|(?:[теория]{6})|6D6D6D6|7Z7Z7Z7)|(\d\s*[п]?\s*/\s*гр)'
# Для лингв. проверки случаев с несколькими преподами
#pattern5 = r'.*(?:[А-ЯЁA-Z][а-яaёa-zА-ЯЁA-Z, ()-]+)(?:[:;\n\d]|$).*?(?=\d\s*[п]?\s*/\s*гр)'
pattern5 = r'.*(?:[А-ЯЁA-Z][а-яaёa-zА-ЯЁA-Z, ()-]+)(?:[:;\n\d]|$).*?(?=(?:\d\s*[п]?\s*/\s*гр)|(?:[а-яёa-z]{2,}[а-яёa-z. \d]+\s{1,}[А-ЯЁA-Z][а-яёa-z ]+(?:[А-ЯЁA-Z][.]?){2}))'

""" Преобразование базы обработки в базу разбора (parse)
Здесь данные заботливо перерабатываются в примитивную БД (список списков)
"""
parse_title = [['Дни', '№', 'Препод', 'Предмет', 'Тип пары', 'п/гр', 'Даты']]
parse = []
default_para = 'общ'    # Обозначение для общей пары
default_type = 'Рейд'   # Обозначение для предмета без типа пары
default_date = timey_wimey # Дата по умолчанию = период
day_num = set()
default_prep = 'Тот-Чьё-Имя-Нельзя-Называть'  # Препод по умолчанию
#default_prep = 'Препод 404'  # Препод по умолчанию
for record in table:
    day_num.add(record[0]) # Множество пройденных дней (для определения текущего)
    if not record[2]: # Если инфы в ячейке предмета нет, то запись не обрабатывается
        continue

    # Список логических предметов в ячейке (элемент = всё что относится к предмету)
    divide = list(a for a in re.findall(pattern1, record[2], flags = re.M))
    # Цикл для обработки строки логического предмета:
    # "Методы оптимизации; 20,27.03; 03,10,17,24.04; 08.05.20г. - теория; доцент КондратьевВ.П.;"
    for ind in range(len(divide)-1, -1, -1):
        # Исправление захвата лишних дат (когда дата, относящаяся к предмету, стоит ПЕРЕД предметом и захватывается не туда)
        if ind > 0:
            dates = re.findall(pattern3, divide[ind-1], flags = re.M)
            if dates:
                divide[ind] = dates[0] + ' ' + divide[ind]
                divide[ind-1] = re.sub(pattern3, '', divide[ind-1], flags = re.M)

        razbor = [] # Список разбора. Один элемент при "Предмет(-ы)...препод", 2+ при "Предмет...преподы"
        # Если есть конструкция "предмет...препод1...преподN"
        if len(re.findall(pattern2a, divide[ind])) > 1:
            # Если даты и типы пар общие для подгрупп (т.е остаток = "гр1-препод1, гр2-препод2...")
            if not re.search(pat2, re.sub(pattern5, '', divide[ind])):
                part1 = re.findall(pattern5, divide[ind])[0] # Предмет; даты. - тип пары:
                part2 = re.sub(pattern5, '', divide[ind])    # 1п/гр.: препод1;...; 'Nп/гр.: преподN
                for pp in re.findall(pattern2b, part2):
                    razbor.append('; '.join([part1, pp[0], pp[1]]))
            else:
                # Шаг 1) Вырезать предмет (перед ним может быть дата, если первый тип - лекция/практика)
                predmet = re.search(pat1, divide[ind], flags = re.M)[1]
                for_sep = re.sub(predmet, '', divide[ind])
                # Шаг 2) Строка без предметов логически делится по преподам
                for razd in re.findall(pattern2a, for_sep):
                    razbor.append('; '.join([predmet, razd]))             
        else:
            razbor.append(divide[ind])
        # День недели - номер пары - предмет - препод - тип пары - подгруппа - даты
        for nabor in razbor:
            #print(nabor)
            # Вырезать предмет из набора в отдельную переменную (если предмет указан)
            if re.search(pat1, nabor, flags = re.M):
                predmet = re.search(pat1, nabor, flags = re.M)[1]
                nabor = re.sub(pat1, '', nabor)
            else: # Если предмета нет (а такое вообще возможно?)
                predmet = 'ОШИБКА, ПРЕДМЕТ НЕ НАЙДЕН'

            # Выделить наборы "даты - тип пары - группа"
            it_inf = iter([a for a in re.findall(pattern4, nabor, flags = re.M) if a]) # Типологический итератор
            dtg = [[[], [], []]] # Логический список дат/типов/групп конструкции
            pred = 9 # Индекс предыдущей группы совпадений, нужен для остановки после комбо + фулсета
            # Комбо, когда 
            for trash in it_inf:
                # Определить, что отловилось (даты - группа №0, типы - 1, группы - 2)
                for i, unint in enumerate(trash):
                    if unint: # [] = False, т.е если не False - элемент = улов
                        # Если не хватает подгруппы, а актуал не подгруппа, то это следующий тип
                        if i!=2 and dtg[-1][0] and dtg[-1][1] and not dtg[-1][2]:
                            dtg[-1][2].append(default_para)
                        if not dtg[-1].count([]) and i!= pred or dtg[-1][1] and i==1: # Если фулсет и не комбо, или следующий тип
                            dtg.append([[], [], []]) # То перейти к следующему элементу
                        dtg[-1][i].append(unint)
                        break # Нет смысла проверять остальное, если искомое нашлось
                pred = i
            # Последняя запись не проверяется на подгруппы (из-за особенностей итерирования)
            if not dtg[-1][2]: # Если нет подгруппы
                dtg[-1][2].append(default_para) # Значение по умолчанию

            # Исправление возможных ошибок в изначальном расписании
            i = 0 # Индексация выносится за цикл, чтобы не париться с косяками при удалении
            while True:
                # Разделение записи по подгруппам, если оно возможно
                if len(dtg[i][2]) > 1:
                    j = len(dtg[i][2])
                    for gr in dtg[i][2]:
                        dtg = dtg[:i+1] + [[dtg[i][0], dtg[i][1], [gr]]] + dtg[i+1:]
                    del(dtg[i])
                    i += j-1
                # Исправление разделения записей вида "ФЗК;  практика: лаб.раб.: 15.04.20г. - 1п/гр"
                if not dtg[i][0]:
                    if i > 0: # Если косяк в "не первой" записи, то склеить её с предыдущей
                        dtg[i-1][1][0] += ', '+dtg[i][1][0] # Склейка типов пары
                        del(dtg[i])
                        i -= 1 # На случай, если дальше что-то будет
                    elif i+1 != len(dtg): # Если косяк в первой записи, то склеить её со следующей
                        dtg[i+1][1][0] += ', '+dtg[i][1][0] # Склейка типов пары
                        del(dtg[i])
                        i -= 1 # На случай, если дальше что-то будет
                # Если запись не содержит тип пары, то он либо был упомянут ранее, либо таков замысел
                if i>0 and not dtg[i][1]:
                    dtg[i][1] = dtg[i-1][1] # Есть предыдущая запись? Взять тип пары из неё
                elif not dtg[i][1]:
                    dtg[i][1].append(default_type) # Запись первая? Тип по умолчанию
                # Если все записи были обработаны, исправление ошибок завершается. Иначе, следующая запись
                if i+2 > len(dtg):
                    break
                else:
                    i += 1

            # Форматирование дат, типа пары и подгруппы
            for f in range(len(dtg)):
                if not dtg[f][0]:
                    dtg[f][0].append(default_date) # Дата по умолчанию
                dtg[f][0] = expand_dates(dtg[f][0], year, len(day_num)-1)
                if dtg[f][1][0] != default_type:
                    dtg[f][1] = [format_tip(dtg[f][1][0])]
                if dtg[f][2][0] != default_para:
                    dtg[f][2] = [format_group(dtg[f][2][0])]
            
            # Выделить препода (выделяется после выделения типов пар)
            # Если выделить до, то при ошибке в синтаксисе изначального расписания можно поймать:
            # Препод = "теория ст.преподаватель БелкинаА.В."
            if re.findall(pat4, nabor): # Если есть препод и должность
                prepod = [format_prep(re.findall(pat4, nabor)[0])]
            elif re.findall(pat5, nabor): # Если есть только препод
                prepod = [format_prep(re.findall(pat5, nabor)[0])]
            else: # Если препод не указан
                if re.findall(pat7, razbor[i], flags = re.I): # Если у текущей пары есть тип, то она не "особая"
                    prepod = [format_prep(parse[-1][3])]
                else: # Пара особая = мероприятие, час куратора и т.п = препод не указывается
                    prepod = [default_prep]

            # Для каждого набора "тип пары - подгруппа - даты" создать запись в БД
            for info in dtg:
                parse.append([record[0],  # День недели
                              record[1],  # Номер пары
                              predmet,    # Предмет
                              prepod[0],  # Препод
                              info[1][0], # Тип пары
                              info[2][0], # Подгруппа
                              info[0]     # Даты
                              #', '.join(info[0])     # Даты (для записи в ячейку экселевской таблицы)
                              ])

# Тестовый вывод для сверки. Дат нет т.к много места занимают. Кабинетов нет т.к пока не добавлены
parse_title = 'Расписание на '+year+'-й год. Учебная часть семестра идёт с '+period[0]+' по '+period[1]
print(f"\n{parse_title : ^188}\n")
print('='*187)
print(f"| {'День' : ^15} | {'№' : ^3} | {'Предмет' : ^65} | {'Препод' : ^55} | {'Тип' : ^15} | {'Для кого' : ^15} |")
print('-'*187)
for i, record in enumerate(parse):
    if i>0 and record[0]!=parse[i-1][0]:
        print('-'*187)
    print(f"| {record[0] : ^15} | {record[1] : ^3} | {record[2] : ^65} | {record[3] : ^55} | {record[4] : ^15} | {record[5] : ^15} |")
print('='*187)

# Сохранение и форматирование нужно будет делать через openpyxl: pyexcel не поддерживает работу со стилями и объединением ячеек :с
