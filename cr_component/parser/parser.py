r"""Парсинг датафрейма с ячейками расписания конкретной группы.

"""

import cr_component.parser.additional as cr_add
import cr_component.parser.defaults as cr_def
import cr_component.parser.exceptions as cr_err
from datetime import date as dt_date, timedelta
import pandas as pd
import numpy as np
import re


def format_prep(lecturer: 'Форматируемая подстрока преподавателя'
                ) -> 'Отформатированная подстрока преподавателя':

    """ Функция форматирования записи преподавателя """
    prep_pat1 = r'(?m)^([а-я. \d]+)\s'          # Должность препода
    prep_pat2 = r'([А-Я][а-я]+)\s*([А-Я.]{4})'  # Препод без должности
    # Если у препода в инициале не хватает второй точки, то доставить её
    if not re.search(r'(?m)\.$', lecturer):
        lecturer += '.'

    # Поставить пробел между фамилией и инициалами, доставить должность если нужно
    prep = ['', re.search(prep_pat2, lecturer)[1], re.search(prep_pat2, lecturer)[2]]
    # Если должности нет, удалить её поле
    if not re.search(prep_pat1, lecturer):
        del(prep[0])
    else:
        prep[0] = re.search(prep_pat1, lecturer)[1]
        prep[0] = re.sub(r'\s+', ' ', prep[0])
        prep[0] = re.sub(r'\.\s+', '.', prep[0])

    return ' '.join(prep)


def format_tip(tip: 'Подстрока форматируемого типа пары'
               ) -> 'Отформатированный тип пары':

    """ Функция форматирования типа пары """
    f_pat = r'(.*?6.*D.*)|(.*?7.*Z.*)|((?:.*?т.*р.*я.*)|(?:.*?л.*к.*я.*))|(.*?а.*б.*)|(.*?п.*а.*к.*)'
    for i, eq in enumerate(re.findall(f_pat, tip)[0]):
        if eq:
            return cr_def.TIP_LIST[i]


def format_group(group: 'Подстрока форматируемой подгруппы'
                 ) -> 'Отформатированная подгруппа':

    """ Функция форматирования подгруппы """
    repair_pat = r'(\d)'
    return re.findall(repair_pat, group)[0]


def expand_dates(dates: 'Список подстрок с датами предмета',
                 year:  'Год (берётся из периода расписания)',
                 day:   'День недели, соответствующий списку дат',
                 cell:   'Обрабатываемая запись (для исключений)'
                 ) -> 'Список объектов дат':

    """ Функция абсолютного разбития дат """
    # Шаблон разделения дат по месяцам
    dates = '; '.join(dates).split(';')
    # Шаблон для дат вида "с...по..."
    dat_pat = r'(?i)с\s*([\d\.]{4,13})[г\.]*\s*по\s*([\d\.]{4,13})[г\.]*'
    # Список для дат записи
    all_dates = []
    for date_string in dates:
        if len(date_string) < 5:
            continue
        # Если дата вида "с 13.01 по 06.06"
        if re.search(dat_pat, date_string):
            start_end = re.search(dat_pat, date_string).groups()
            dt_start = list(map(int, list(filter(None, start_end[0].split('.')))))
            dt_final = list(map(int, list(filter(None, start_end[1].split('.')))))
            dt_start = dt_date(year, dt_start[1], dt_start[0])
            # Пока день недели не совпадёт с тем что "сейчас" в базе (на случай косяков в периоде)
            if dt_start.weekday() != day:
                if dt_start.weekday() < day:
                    dt_start += timedelta(days=day-dt_start.weekday())
                else:
                    dt_start += timedelta(days=day-dt_start.weekday())
            dt_final = dt_date(year, dt_final[1], dt_final[0])
            # Пока не достигнут конец периода
            while dt_start <= dt_final:
                all_dates.append(dt_start)
                dt_start += timedelta(days=7)
        # Если дата вида "03,17,24,31,03" или "22,02,01.03.19г"
        elif date_string:
            date = list(map(int, re.findall(r'(\d+)', date_string)))
            # Костыль. Если последнее число - год, то выкинуть его
            if date and year in [date[-1], date[-1]+2000]:
                del(date[-1])
            # Прогон в обратную сторону, чтобы отловить косяки типа "22,02,19,01.03.19г"
            if len(date) < 1:
                raise cr_err.IncorrectDate(cell, date_string)
            month, date = date[-1], date[:-1]
            for d in range(len(date)-1, -1, -1):
                # Установка текущей даты
                try:
                    dt_start = dt_date(year, month, date[d])
                except:
                    month = date[d]
                    if month not in list(range(1, 12+1)):
                        raise cr_err.IncorrectDate(cell, date_string)
                    dt_start = dt_date(year, month, date[d])

                # Проверка даты на правильность дня
                if dt_start.weekday() != day:
                    dt_start += timedelta(days=day-dt_start.weekday())
                else:
                    dt_start += timedelta(days=day-dt_start.weekday())
                all_dates.append(dt_start)

                # Проверка на обновление месяца (если число "слева" - предыдущий месяц, но не день)
                if d > 1 and date[d-1] == dt_start.month-1 and date[d-1] != dt_start.day-7:
                    month = date[d]

    # После абсолютного разбития, возвращается отсортированный список дат записи без повторов
    return sorted(list(set(all_dates)))


def repl_b(cab: 'Подстрока с одним кабинетом'
           ) -> 'Отформатированная подстрока кабинета':

    """ Форматирование записи корпусов кабинетов """
    cab = re.sub(r'[.;: ]', '', cab)          # Удаление лишних разделителей
    cab = re.sub(r'[уУ][кК]?№', ' УК№', cab)  # Форматирование полной записи кабинета

    return cab


def parser(stuff: 'База обработки',
           timey: 'Период расписания',
           year:  'Год расписания'
           ) -> ('База парсинга', 'Список ошибок'):

    """ Парсер предварительного набора данных расписания группы """

    """ Регулярные шаблоны для переработки информации """
    # Краткие шаблоны для обычных случаев
    # Отлов предмета (зачёты и дифы исключаются до отлова)
    pat1 = r'(?m)([А-ЯЁA-Z][А-ЯЁA-Zа-яёa-z, -]{2,}?[А-ЯЁA-Zа-яёa-z( )]+?)(?=(?: с \d)|[:;\n\d]|$)'
    # Отлов препода и должности (если есть)
    pat2 = r'(?:[а-яёa-z]{2,}[а-яёa-z. \d]+\s{1,})?[А-Я][а-я]+\s*[А-Я.]{3,4}'
    # Даты
    pat3 = r'(?mi)(?:(?:с\s*(?:\d{1,2}[.,г;\s]*\s*)+по\s*(?:\d{1,2}(?:[.,г;\s]|$)\s*)+))|(?:(?:\d{1,2}(?:[.,г;\s]|$)\s*)+)'
    # Отлов подгрупп
    pat4 = r'\d\s*[п]?\s*/\s*гр'
    # Отлов типа пары
    pat5 = r'(?i)(?:[практика]{8})|(?:[лаб.раб ]{8})|(?:[лекция]{6})|(?:[теория]{6})|6D6D6D6|7Z7Z7Z7'

    # Длинные шаблоны для особых случаев
    # Разделение по преподам
    pattern1a = r'(?m)(?:.*?[А-ЯЁA-Z][А-ЯЁA-Zа-яёa-z, -]{2,}?[А-ЯЁA-Zа-яёa-z( )]+?(?=(?: с \d)|[:;\n\d]|$)(?:(?:.(?![А-ЯЁA-Z][А-ЯЁA-Zа-яёa-z, -]{2,}?[А-ЯЁA-Zа-яёa-z( )]+?(?=(?: с \d)|[:;\n\d]|$)))*[А-Я][а-я]+\s*[А-Я.]{3,4}))|(?:.*?[А-ЯЁA-Z][А-ЯЁA-Zа-яёa-z, -]{2,}?[А-ЯЁA-Zа-яёa-z( )]+?(?=(?: с \d)|[:;\n\d]|$))'
    # Разделение по предметам
    pattern1b = r'(?m)(?:.*?(?:[А-ЯЁA-Z][А-ЯЁA-Zа-яёa-z, -]{2,}?[А-ЯЁA-Zа-яёa-z( )]+?(?=(?: с \d)|[:;\n\d]|$)).*?)(?=(?:[А-ЯЁA-Z][а-яёa-zА-ЯЁA-Z, ()-]{3,}[:;\n\d])|(?:$))'
    # Разделение по преподам
    pattern2a = r'.*?(?:[а-яёa-z.]{2,}\s[А-ЯЁA-Z][а-яёa-z\s]+(?:[А-ЯЁA-Z][.]?){2})'
    # Разделение по преподам: выделение подгрупп и преподов (если после инфы просто перечисление подгрупп-преподов)
    pattern2b = r'(\d\s*[п]?\s*/\s*гр).*?((?:[а-яёa-z]{2,}[а-яёa-z. \d]+\s{1,})?[А-ЯЁA-Z][а-яёa-z ]{2,}(?:[А-ЯЁA-Z][.]?){2})'
    # Отлов даты в конце строки (для конкретных случаев, когда косяк в захвате даты)
    pattern3 = r'(?m)(?:(?:(?:[сС]\s*(?:\d{2}[.,г;\s]*\s*)+[пП][оО]\s*(?:\d{2}[.,г;\s]{0,2}\s*)+))|(?:(?:\d{2}[.,г;\s]{0,2}\s*)+))$'
    # Отлов дат, типов пары и групп как [даты, типы пары, группы]
    pattern4 = r'(?m)((?:(?:[сС]\s*(?:\d{2,4}[.,г;\s]*\s*)+[пП][оО]\s*(?:\d{2,4}(?:[.,г;\s]|$)\s*)+))|(?:(?:\d{2}(?:[.,г;\s]|$)\s*)+))|((?:[практика]{8})|(?:[лаб.раб ]{8})|(?:[лекция]{6})|(?:[теория]{6})|6D6D6D6|7Z7Z7Z7)|(\d\s*[п]?\s*/\s*гр)'
    # Для лингв. проверки случаев с несколькими преподами
    pattern5 = r'(?m).*(?:[А-ЯЁA-Z][А-ЯЁA-Zа-яёa-z, -]{2,}?[А-ЯЁA-Zа-яёa-z( )]+?)(?=(?: с \d)|[:;\n\d]|$).*?(?=(?:\d\s*[п]?\s*/\s*гр)|(?:[а-яёa-z]{2,}[а-яёa-z. \d]+\s{1,}[А-ЯЁA-Z][а-яёa-z ]+(?:[А-ЯЁA-Z][.]?){2}))'

    # Стандартизация форматного умолчания препода
    cr_def.DEF_TEACHER = format_prep(cr_def.DEF_TEACHER)

    """ Обработка начальной БД """
    # Датафрейм базы парсинга
    parse = pd.DataFrame(columns=cr_def.DEF_COLUMNS)

    day_num = set()  # Множество для определения пройденных дней

    """ Проход по всем строкам расписания группы """
    for index, record in stuff.iterrows():
        start = parse.shape[0]      # Для сохранения первой строки новой ячейки
        day_num.add(record['day'])  # Множество пройденных дней (для определения текущего)

        """ Первичное обрубание: по преподам (перед предметом может идти доп. инфа вроде дат, но не препод) """
        # ...но после последнего препода в логическом наборе предмета, начинается инфа о след. предмете
        divide = list(a for a in re.findall(pattern1a, record['rec']))

        # Из-за особенностей работы с индексами, обработка оптимальна через while (т.к длина divide может меняться)
        # Минусы: обработка может замедляться
        # Плюсы: для полной обработки не нужно повторно проходить по разбору (уменьшение временной сложности)

        ind = 0  # Абсолютный индекс актуального элемента в divide
        while True:
            rec = divide[ind]

            # Можно обойтись и без условия, но оно помогает избежать лишних проверок (нужно для оптимизации)
            # Цикл для разделения хлама вида "ВМ...МО...Кондратьев В.П." или "Ивент...нормальный предмет и его препод"
            if len(re.findall(pat1, rec)) > 1:
                # Для комфортного добавления препода в строку
                lecturer = re.search(pat2, rec)[0]
                # Относительный индекс для отслеживания новых записей
                i = 0

                # Обособление предметных записей
                for predm_rec in re.findall(pattern1b, rec):
                    # Выделить предмет в отдельную запись
                    divide = divide[:ind+i+1] + [predm_rec] + divide[ind+i+1:]
                    # Увеличить относительный индекс
                    i += 1
                    # Ивентовый предмет без препода - нечто вида "09.01.20; Час куратора" (тип пары никогда не указан)
                    # Если предмет не ивентовый (т.е есть тип пары), а препода нет - "и дайте этому предмету препода"
                    if re.search(pat5, predm_rec) and not re.search(pat2, predm_rec):
                        divide[ind+i] += '; ' + lecturer

                # Удалить изначальную разделяемую запись
                del(divide[ind])

            """ Этап формирования разборного списка записей "предмет и инфа о нём" """
            # В разборном списке 1 элемент по умолчанию, 2+ если запись имеет вид "Предмет...преподы"
            parse_list = []

            """ Проверка на конструкцию с несколькими преподами "предмет...препод1...преподN" """
            if len(re.findall(pattern2a, divide[ind])) > 1:
                # Если даты и типы пар общие для подгрупп (т.е остаток = "гр1-препод1, гр2-препод2...")
                if not re.search(pat3, re.sub(pattern5, '', divide[ind])):
                    part1 = re.findall(pattern5, divide[ind])[0]  # Предмет; даты. - тип пары:
                    part2 = re.sub(pattern5, '', divide[ind])     # 1п/гр.: препод1;...; 'Nп/гр.: преподN

                    for pp in re.findall(pattern2b, part2):
                        # Так как инфа в записи общая для ячейки, и меняются только препод с подгруппой...
                        # То для каждого типа пары нужно добавить подгруппу...
                        # Подгруппа указывается один раз, из-за чего могут быть баги...
                        # При этом, оригинальное указание подгруппы затирается
                        parse_list.append('; '.join([re.sub(pat5,
                                                            lambda m: m[0] + ': ' + pp[0],
                                                            re.sub(pat4, '', part1)),
                                                     pp[1]]))
                else:
                    # Шаг 1) Вырезать предмет (перед ним может быть дата, если первый тип - лекция/практика)
                    edu_pair = re.search(pat1, divide[ind])[1]
                    for_sep = re.sub(edu_pair, '', divide[ind])
                    # Шаг 2) Строка без предметов логически делится по преподам
                    for razd in re.findall(pattern2a, for_sep):
                        parse_list.append('; '.join([edu_pair, razd]))
            else:
                parse_list.append(divide[ind])

            """ Прогон по каждому логическому набору в разборном списке """
            for logic_set in parse_list:
                # Вырезать предмет из набора в отдельную переменную (если предмет указан)
                if re.search(pat1, logic_set):
                    edu_pair = re.search(pat1, logic_set)[1]
                    logic_set = re.sub(pat1, '', logic_set)
                # Если предмета нет (такого быть не должно, но мало ли)
                else:
                    edu_pair = 'TBD'

                """ Выделение наборов "даты - тип пары - группа" """
                it_inf = iter(a for a in re.findall(pattern4, logic_set) if a)  # Типологический итератор
                dtg = [[[], [], []]]  # Логический список дат/типов/групп конструкции
                pred = 9  # Индекс предыдущей группы совпадений, нужен для остановки после комбо + фулсета
                # Комбо, когда очередь из одинаковых типов пар
                for trash in it_inf:
                    # Определить, что отловилось (даты - группа №0, типы - 1, группы - 2)
                    for i, unint in enumerate(trash):
                        if unint:
                            # Если не хватает подгруппы, а актуал не подгруппа, то это уже следующий набор
                            if i != 2 and dtg[-1][0] and dtg[-1][1] and not dtg[-1][2]:
                                dtg[-1][2].append(cr_def.DEF_GROUPS)
                            # Если фулсет и не комбо, или следующий набор
                            if not dtg[-1].count([]) and i != pred or dtg[-1][1] and i == 1:
                                # То перейти к следующему элементу
                                dtg.append([[], [], []])
                            dtg[-1][i].append(unint)
                            # Нет смысла проверять остальное, если искомое нашлось
                            break
                    pred = i

                """ Проверка последней записи """
                if not dtg[-1][2]:                  # Если нет подгруппы
                    dtg[-1][2].append(cr_def.DEF_GROUPS)  # Значение по умолчанию

                """ Исправление возможных ошибок в изначальном расписании """
                # Длина списка может меняться, из-за чего приходится использовать while
                i = 0
                while True:
                    """ Разделение записи по подгруппам, если оно возможно """
                    if len(dtg[i][2]) > 1:
                        j = len(dtg[i][2])
                        for gr in dtg[i][2]:
                            dtg = dtg[:i+1] + [[dtg[i][0], dtg[i][1], [gr]]] + dtg[i+1:]
                        del(dtg[i])
                        i += j-1

                    """ Исправление разделения записей вида "ФЗК;  практика: лаб.раб.: 15.04.20г. - 1п/гр" """
                    # Если в записи косяк
                    if not dtg[i][0]:
                        # Если косяк в "не первой" записи, то склеить её с предыдущей
                        if i > 0:
                            # Проверка наличия записи (во избежание ошибок разбития даты)
                            if not dtg[i-1][1] or not dtg[i][1]:
                                raise cr_err.IncorrectCell(rec)
                            # Склейка типов пары
                            dtg[i-1][1][0] += ', '+dtg[i][1][0]
                            del(dtg[i])
                            # Переиндексация на случай, если дальше что-то будет
                            i -= 1
                        # Если косяк в первой записи, то склеить её со следующей
                        elif i+1 != len(dtg):
                            # Склейка типов пары
                            dtg[i+1][1][0] += ', '+dtg[i][1][0]
                            del(dtg[i])
                            # Переиндексация на случай, если дальше что-то будет
                            i -= 1

                    """ Если запись не содержит тип пары, то он либо был упомянут ранее, либо таков замысел """
                    # Если есть предыдущая запись
                    if i > 0 and not dtg[i][1]:
                        # Взять тип пары из неё
                        dtg[i][1] = dtg[i-1][1]
                    # Если типа пары нет
                    else:
                        # Взять значение по умолчанию
                        dtg[i][1].append(cr_def.DEF_TYPE_PAIR)

                    """ Проверка на завершение цикла коррекции """
                    # Если все записи были обработаны, исправление ошибок завершается
                    if i+2 > len(dtg):
                        break
                    # Если ещё остались какие-то записи, продолжить коррекцию
                    else:
                        i += 1

                """ Форматирование дат, типа пары и подгруппы """
                for f in range(len(dtg)):
                    # Если даты нет
                    if not dtg[f][0]:
                        # Поставить датой период расписания
                        dtg[f][0] = [f'с {timey[0]} по {timey[1]}']
                    # Раскрыть даты в список дат
                    dtg[f][0] = expand_dates(dtg[f][0], year, len(day_num)-1, rec)

                    # Если тип пары не был определён
                    if not dtg[f][1]:
                        # Использовать тип пары по умолчанию
                        dtg[f][1].append(cr_def.DEF_TYPE_PAIR)
                    # Если тип пары есть и это не значение по умолчанию
                    elif dtg[f][1][0] != cr_def.DEF_TYPE_PAIR:
                        # Отформатировать его
                        dtg[f][1] = [format_tip(dtg[f][1][0])]

                    # Если подгруппа не значение по умолчанию
                    if dtg[f][2][0] != cr_def.DEF_GROUPS:
                        # Отформатировать её
                        dtg[f][2] = [format_group(dtg[f][2][0])]

                """ Выделить препода (выделяется после обособления типов пар) """
                # Если выделить до, то при ошибке в синтаксисе изначального расписания можно поймать:
                # Препод = "теория ст.преподаватель БелкинаА.В."

                # Если есть препод, отловить его с должностью (если указана)
                if re.findall(pat2, logic_set):
                    lecturer = [format_prep(re.findall(pat2, logic_set)[0])]
                # Если препод не указан
                else:
                    # Если у текущей пары есть тип, то она не "особая"
                    if re.findall(pat5, parse_list[i]):
                        lecturer = [format_prep(parse.iloc[-1]['teacher'])]
                    # Пара особая = мероприятие, час куратора и т.п = стандартный препод
                    else:
                        lecturer = [cr_def.DEF_TEACHER]

                """ Для каждого набора "тип пары - подгруппа - даты" создать запись в БД """
                for info in dtg:
                    # Финальная корректировка типа
                    if len(parse) and info[1][0] == cr_def.DEF_TYPE_PAIR:
                        # Если предмет как предыдущий, а тип базовый, то ошибка в типе
                        if parse.iloc[-1]['item_name'] == edu_pair:
                            info[1] = [parse.iloc[-1]['type']]
                        # Если тип базовый, но есть подгруппы - сменить на базовый №2
                        elif info[2][0] != cr_def.DEF_GROUPS:
                            info[1][0] = cr_def.DEF_TYPE_GRPS

                    """ Занесение в базу """
                    # Кабинеты добавляются отдельно, так как с ними много проблем
                    new_row = {'day'      : record['day'],  # День недели  # 0
                               'num'      : record['num'],  # Номер пары   # 1
                               'item_name': edu_pair,       # Предмет      # 2
                               'teacher'  : lecturer[0],    # Препод       # 3
                               'type'     : info[1][0],     # Тип пары     # 4
                               'pdgr'     : info[2][0],     # Подгруппа    # 5
                               'date_pair': info[0],        # Даты         # 6
                               'cab'      : np.nan          # Кабинет      # 7
                               }

                    parse.loc[parse.shape[0]] = new_row

            if ind+1 == len(divide):
                break
            else:
                ind += 1

        """ Логика распределения кабинетов """
        # Обработка ячейки кабинетов (если есть что обрабатывать)
        if record['cabs']:
            # Отдельная обработка каждого кабинета в ячейке
            cabs = [repl_b(cb) for cb in re.split(r'[;,]', record['cabs'])
                    if not re.fullmatch(r'[.,:; ]*', cb)]  # Но некорректные записи отсеиваются

            # Дополнительная коррекция данных
            # Исправление ['210', '212 УК№1', '329', '331 УК№5', '410 УК№1'] в полных записях
            for ind_cb in range(len(cabs)-1, -1, -1):
                # Если кабинет уже заменён базовым значением из констант, то его форматировать нельзя
                if cabs[ind_cb] == cr_def.DEF_CABS or cabs[ind_cb] == cr_def.DEF_SPORT_CAB:
                    continue
                # Если для кабинета не указан корпус
                if (ind_cb + 1 < len(cabs) and
                        re.match(r'\d+', cabs[ind_cb]) and
                        not re.match(r'\d+ УК№\d', cabs[ind_cb])):
                    if re.search(r' УК№\d', cabs[ind_cb+1]):                     # Если он указан дальше
                        cabs[ind_cb] += re.search(r' УК№\d', cabs[ind_cb+1])[0]  # Взять корпус оттуда
                    else:                                                        # Неоткуда брать корпус?
                        cabs[ind_cb] += ' УК№?'                                  # Обозначить этот факт

        # Если кабинетов нет
        else:
            cabs = [cr_def.DEF_CABS]
        
        it_cb = iter(cabs)    # Итератор кабинетов

        # Проход по всем записям в строке
        for ind in range(start, parse.shape[0]):
            """ Подсчитать количество уникальных записей и преподов """
            n_rec = n_prp = 0
            for i in range(start, parse.shape[0]):
                # Если нет совпадения с предыдущей записью (без учёта подгруппы, дат, кабинета)
                if i == 0 or not parse.loc[i-1][:5].eq(parse.loc[i][:5]).min():
                    # То запись - уникальная
                    n_rec += 1
                    # А если нет совпадения, то больше шансов что препод отличается
                    i_end = i-1 if (i-1) >= 0 else parse.shape[0]-1
                    if parse.loc[i_end]['teacher'] != parse.loc[i]['teacher']:
                        n_prp += 1

            """ Расстановка кабинетов """
            # Если ячейке сопряжён только один кабинет (или в ней лишь одна запись)
            if len(cabs) == 1 or start+1 == parse.shape[0]:
                # Если препод по умолчанию, а записей несколько
                if parse.loc[ind]['teacher'] == cr_def.DEF_TEACHER and n_rec > 1:
                    # Поставить кабинет по умолчанию
                    parse.loc[ind, 'cab'] = cr_def.DEF_CABS
                # Если запись одна
                else:                        
                    # Поставить первый кабинет из списка кабинетов ячейки
                    parse.loc[ind, 'cab'] = cabs[0]
            else:
                # Если ФЗК для нескольких групп, то могут быть косяки
                if (ind and
                        parse.loc[ind-1][:3].eq(parse.loc[ind][:3]).min() and
                        parse.loc[ind-1]['cab'] == cr_def.DEF_SPORT_CAB):
                    parse.loc[ind, 'cab'] = cr_def.DEF_SPORT_CAB
                    continue
                # Если число записей совпадает с числом кабинетов
                if n_rec == len(cabs):
                    # Итерироваться при смене УНИКАЛЬНОЙ записи (без учёта подгрупп)                    
                    if ind == start or not parse.loc[ind-1][:5].eq(parse.loc[ind][:5]).min():
                        """ Итерация """
                        try:
                            parse.loc[ind, 'cab'] = next(it_cb)
                        except:
                            parse.loc[ind, 'cab'] = parse.loc[ind-1]['cab']
                    else:
                        parse.loc[ind, 'cab'] = parse.loc[ind-1]['cab']
                else:
                    # Если преподов столько же сколько и кабинетов
                    if n_prp == len(cabs):
                        # Итерироваться при смене препода
                        if ind == start or parse.loc[ind-1]['teacher'] != parse.loc[ind]['teacher']:
                            """ Итерация """
                            try:
                                parse.loc[ind, 'cab'] = next(it_cb)
                            except:
                                parse.loc[ind, 'cab'] = parse.loc[ind-1]['cab']
                        else:
                            parse.loc[ind, 'cab'] = parse.loc[ind-1]['cab']
                    else:
                        # Итерироваться при одном из условий:
                        if (ind == start or
                            # 1) Сменился препод и не переход с лекцию на лекцию PAIR_TYPES.values
                            (parse.loc[ind-1]['teacher'] != parse.loc[ind]['teacher'] and
                             not (parse.loc[ind-1]['type'] == parse.loc[ind]['type'] == cr_def.TIP_LIST[2])) or
                            # 2) Препод не сменился и переход с/на лекцию
                            (parse.loc[ind-1]['teacher'] == parse.loc[ind]['teacher'] and
                             (parse.loc[ind-1]['type'] == cr_def.TIP_LIST[2] or
                              parse.loc[ind]['type'] == cr_def.TIP_LIST[2]))):
                            """ Итерация """
                            try:
                                parse.loc[ind, 'cab'] = next(it_cb)
                            except:
                                parse.loc[ind, 'cab'] = parse.loc[ind-1]['cab']
                        else:
                            parse.loc[ind, 'cab'] = parse.loc[ind-1]['cab']

    # Применение технической стандартизации к датафрейму
    parse = cr_add.use_standard(parse)

    # Имя берётся из предыдущей стадии эволюции БД
    try:
        parse.name = stuff.name
    except:
        parse.name = cr_def.DEF_NAME

    # Если с парсингом всё было норм, вернуть базу запарсенного расписания
    return parse
