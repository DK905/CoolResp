""" Информация о модуле
Данный модуль предназначен для парсинга разобранной базы расписания группы.

Известные проблемы:
    1) Кабинеты (в 1% случаев) могут быть ошибочными
    UPD1: Процентовка снижена с 5% до 1% путём переработки логики кабинетов
    UPD2: Для ошибок я сам не понимаю, по какой безумной логике расставлялись кабинеты

"""

# Импорт базовых команд поиска и замены из стандартного модуля регулярных выражений
from re import findall, fullmatch, match, search, split as rsplit, sub

# Импорт команд для обработки дат из модуля datetime
from datetime import date as dt_date, timedelta

# Импорт умолчаний и сокращений для типов пар
from modules.CR_dataset import BadDataError, defaults, tip_list


def format_prep(prepod : 'Подстрока форматируемого препода',
                dv_yn  : 'Флаг сокращения должности препода'
                ) ->     'Отформатированный препод':

    """ Функция форматирования препода """
    prep_pat1 = r'(?m)^([а-я. \d]+)\s'         # Должность
    prep_pat2 = r'([А-Я][а-я]+)\s*([А-Я.]{4})' # Препод без должности
    # Если у препода в инициале не хватает второй точки, то доставить её
    if not search(r'(?m)\.$', prepod):
        prepod += '.'

    # Поставить пробел между фамилией и инициалами, доставить должность если нужно
    prep = ['', search(prep_pat2, prepod)[1], search(prep_pat2, prepod)[2]]
    if dv_yn or not search(prep_pat1, prepod):
        del(prep[0])
    else:
        prep[0] = search(prep_pat1, prepod)[1]
        prep[0] = sub(r'\s+', ' ', prep[0])
        prep[0] = sub(r'\.\s+', '.', prep[0])

    return ' '.join(prep)


def format_tip(tip : 'Подстрока форматируемого типа пары'
               ) ->  'Отформатированный тип пары':

    """ Функция форматирования типа пары """
    f_pat = r'(.*?6.*D.*)|(.*?7.*Z.*)|(.*?т.*р.*я.*)|(.*?л.*к.*я.*)|(.*?а.*б.*)|(.*?п.*а.*к.*)'
    for i, sovp in enumerate(findall(f_pat, tip)[0]):
        if sovp:
            return tip_list[i]


def format_group(group : 'Подстрока форматируемой подгруппы',
                 dv_yn : 'Флаг сокращения подгруппы'
                 ) ->    'Отформатированная подгруппа':

    """ Функция форматирования подгруппы """
    repair_pat = r'(\d)'
    if dv_yn:
        return findall(repair_pat, group)[0]
    else:
        return f'{findall(repair_pat, group)[0]}п/гр'


def expand_dates(dates : 'Список подстрок с датами предмета',
                 year  : 'Год (берётся из периода расписания)',
                 day   : 'День недели, соответствующий списку дат'
                 ) ->    'Список объектов дат':

    """ Функция абсолютного разбития дат """
    # Шаблон разделения дат по месяцам. Идеален для 99% случаев, но в 1% встречает "22,02,01.03.19г"
    dates = rsplit(r';', '; '.join(dates))
    # Шаблон для дат вида "с...по..."
    dat_pat = r'с\s*([\d.]{5,8})[г.]*\s*по\s*([\d.]{5,8})[г.]*'
    # Список для дат записи
    all_dates = []
    for date in dates:
        # Если дата вида "с 13.01 по 06.06"
        if search(dat_pat, date):
            start_end = search(dat_pat, date).groups()
            dt_start = list(map(int, start_end[0].split('.')))
            dt_final = list(map(int, start_end[1].split('.')))
            dt_start = dt_date(year, dt_start[1], dt_start[0])
            # Пока день недели не совпадёт с тем что "сейчас" в базе (на случай косяков в периоде)
            while dt_start.weekday() != day:
                dt_start += timedelta(days = 1)
            dt_final = dt_date(year, dt_final[1], dt_final[0])
            # Пока не достигнут конец периода
            while dt_start <= dt_final:
                all_dates.append(dt_start)
                dt_start += timedelta(days = 7)
        # Если дата вида "03,17,24,31,03" или "22,02,01.03.19г"
        # Примечание: такую дату можно было бы парсить гораздо проще, но в 700+ тестовых записей было целых два косяка
        elif date: 
            date = list(map(int, findall(r'(\d+)', date)))
            # Костыль. Если последнее число - год, то выкинуть его
            if date and date[-1]+2000 == year:
                del(date[-1])
            # Прогон в обратную сторону, чтобы отловить косяки типа "22,02,19,01.03.19г"
            month, date = date[-1], date[:-1]
            for d in range(len(date)-1, -1, -1):
                if not d or d and date[d]%7 == date[d-1]%7: # Если нормальная череда дней
                    all_dates.append(dt_date(year, month, date[d]))
                else: # Если какой-то косяк
                    if date[d] == year: # Если указан год (как в примере с 22,02,19)
                        continue
                    else:         # Если указан месяц, обновить
                        month = date[d]

    # После абсолютного разбития, возвращается отсортированный список дат записи без повторов
    return sorted(list(set(all_dates)))


def parser(trash : 'База обработки',
           year  : 'Год расписания',
           g_yn  : 'Флаг сокращения подгрупп',
           p_yn  : 'Флаг сокращения преподов'
           ) ->    'База парсинга':

    """ Парсер базы обработки """

    """ Регулярные шаблоны для переработки информации """
    # Краткие шаблоны для обычных случаев
    pat1 = r'(?m)([А-ЯЁA-Z][А-ЯЁA-Zа-яaёa-z, -]{2,}?[А-ЯЁA-Zа-яaёa-z( )]+?)(?=(?: с \d)|[:;\n\d]|$)' # Отлов предмета (зачёты и дифы исключаются до отлова)
    pat2 = r'(?:[а-яёa-z]{2,}[а-яёa-z. \d]+\s{1,})?[А-Я][а-я]+\s*[А-Я.]{3,4}' # Отлов препода и должности (если есть)
    pat3 = r'(?m)(?:(?:с\s*(?:\d{2}[.,г;\s]*\s*)+по\s*(?:\d{2}(?:[.,г;\s]|$)\s*)+))|(?:(?:\d{2}(?:[.,г;\s]|$)\s*)+)' # Даты
    pat4 = r'\d\s*[п]?\s*/\s*гр' # Отлов подгрупп
    pat5 = r'(?i)(?:[практика]{8})|(?:[лаб.раб ]{8})|(?:[лекция]{6})|(?:[теория]{6})|6D6D6D6|7Z7Z7Z7' # Отлов типа пары

    # Длинные шаблоны для особых случаев
    # Разделение по преподам
    pattern1a = r'(?m)(?:.*?[А-ЯЁA-Z][А-ЯЁA-Zа-яaёa-z, -]{2,}?[А-ЯЁA-Zа-яaёa-z( )]+?(?=(?: с \d)|[:;\n\d]|$)(?:(?:.(?![А-ЯЁA-Z][А-ЯЁA-Zа-яaёa-z, -]{2,}?[А-ЯЁA-Zа-яaёa-z( )]+?(?=(?: с \d)|[:;\n\d]|$)))*[А-Я][а-я]+\s*[А-Я.]{3,4}))|(?:.*?[А-ЯЁA-Z][А-ЯЁA-Zа-яaёa-z, -]{2,}?[А-ЯЁA-Zа-яaёa-z( )]+?(?=(?: с \d)|[:;\n\d]|$))'
    # Разделение по предметам
    pattern1b = r'(?m)(?:.*?(?:[А-ЯЁA-Z][А-ЯЁA-Zа-яaёa-z, -]{2,}?[А-ЯЁA-Zа-яaёa-z( )]+?(?=(?: с \d)|[:;\n\d]|$)).*?)(?=(?:[А-ЯЁA-Z][а-яёa-zА-ЯЁA-Z, ()-]{3,}[:;\n\d])|(?:$))'
    # Разделение по преподам
    pattern2a = r'.*?(?:[а-яёa-z.]{2,}\s[А-ЯЁA-Z][а-яёa-z\s]+(?:[А-ЯЁA-Z][.]?){2})'
    # Разделение по преподам: выделение подгрупп и преподов (если после инфы просто перечисление подгрупп-преподов)
    pattern2b = r'(\d\s*[п]?\s*/\s*гр).*?((?:[а-яёa-z]{2,}[а-яёa-z. \d]+\s{1,})?[А-ЯЁA-Z][а-яёa-z ]{2,}(?:[А-ЯЁA-Z][.]?){2})'
    # Отлов даты в конце строки (для конкретных случаев, когда косяк в захвате даты)
    pattern3 = r'(?m)(?:(?:(?:с\s*(?:\d{2}[.,г;\s]*\s*)+по\s*(?:\d{2}[.,г;\s]{0,2}\s*)+))|(?:(?:\d{2}[.,г;\s]{0,2}\s*)+))$'
    # Отлов дат, типов пары и групп как [даты, типы пары, группы]
    pattern4 = r'(?m)((?:(?:с\s*(?:\d{2}[.,г;\s]*\s*)+по\s*(?:\d{2}(?:[.,г;\s]|$)\s*)+))|(?:(?:\d{2}(?:[.,г;\s]|$)\s*)+))|((?:[практика]{8})|(?:[лаб.раб ]{8})|(?:[лекция]{6})|(?:[теория]{6})|6D6D6D6|7Z7Z7Z7)|(\d\s*[п]?\s*/\s*гр)'
    # Для лингв. проверки случаев с несколькими преподами
    pattern5 = r'(?m).*(?:[А-ЯЁA-Z][А-ЯЁA-Zа-яaёa-z, -]{2,}?[А-ЯЁA-Zа-яaёa-z( )]+?)(?=(?: с \d)|[:;\n\d]|$).*?(?=(?:\d\s*[п]?\s*/\s*гр)|(?:[а-яёa-z]{2,}[а-яёa-z. \d]+\s{1,}[А-ЯЁA-Z][а-яёa-z ]+(?:[А-ЯЁA-Z][.]?){2}))'

    """ Приведение препода по умолчанию к общей форме """
    try:
        defaults[5] = format_prep(defaults[5], p_yn)
    except:
        pass

    """ Обработка начальной БД """
    try:
        parse = []      # База разбора
        day_num = set() # Множество для определения пройденных дней

        """ Проход по всем строкам расписания группы """
        for record in trash:
            start = len(parse)      # Для сохранения первой строки новой ячейки
            day_num.add(record[0])  # Множество пройденных дней (для определения текущего)
            if not record[2]: # Если инфы в ячейке предмета нет, то запись не обрабатывается
                continue

            """ Первичное обрубание: по преподам (перед предметом может идти доп. инфа вроде дат, но не препод) """
            # ...но после последнего препода в логическом наборе предмета, начинается инфа о след. предмете
            divide = list(a for a in findall(pattern1a, record[2]))

            # Из-за особенностей работы с индексами, обработка оптимальна через while (т.к длина divide может меняться)
            # Минусы: нужно внимательно отлавливать бесконечный цикл
            # Плюсы: для полной обработки не нужно повторно проходить по разбору (уменьшение временной сложности)
            ind = 0 # Абсолютный индекс актуального элемента в divide
            while True:
                rec = divide[ind]

                # Можно обойтись и без условия, но оно помогает избежать лишних проверок (нужно для оптимизации)
                # Цикл для разделения хлама вида "ВМ...МО...Кондратьев В.П." или "Ивент...нормальный предмет и его препод"
                if len(findall(pat1, rec)) > 1:
                    prepod = search(pat2, rec)[0] # Для комфортного добавления препода в строку
                    i = 0 # Относительный индекс для отслеживания новых записей
                    for predm_rec in findall(pattern1b, rec):
                        # Выделить предмет в отдельную запись
                        divide = divide[:ind+i+1] + [predm_rec] + divide[ind+i+1:]
                        i += 1 # Увеличить относительный индекс
                        # Ивентовый предмет без препода - нечто вида "09.01.20; Час куратора" (типа пары никогда нет)
                        # Если предмет не ивентовый (т.е есть тип пары), а препода нет - "и дайте этому предмету препода"
                        if search(pat5, predm_rec) and not search(pat2, predm_rec):
                            divide[ind+i] += '; '+prepod
                    # Удалить изначальную разделяемую запись
                    del(divide[ind])
                    rec = divide[ind]

                """ Этап формирования разборного списка записей "предмет и инфа о нём" """
                # В разборном списке 1 элемент по умолчанию, 2+ если запись имеет вид "Предмет...преподы"
                razbor = []

                """ Проверка на конструкцию с несколькими преподами "предмет...препод1...преподN" """
                if len(findall(pattern2a, divide[ind])) > 1:
                    # Если даты и типы пар общие для подгрупп (т.е остаток = "гр1-препод1, гр2-препод2...")
                    if not search(pat3, sub(pattern5, '', divide[ind])):
                        part1 = findall(pattern5, divide[ind])[0] # Предмет; даты. - тип пары:
                        part2 = sub(pattern5, '', divide[ind])    # 1п/гр.: препод1;...; 'Nп/гр.: преподN
                        for pp in findall(pattern2b, part2):
                            # Так как инфа в записи общая для ячейки, и меняются только препод с подгруппой...
                            # То для каждого типа пары нужно добавить подгруппу (она указывается один раз, из-за чего могут быть баги)
                            # При этом, оригинальное указание подгруппы затирается
                            razbor.append('; '.join([sub(pat5, lambda m: m[0]+': '+pp[0], sub(pat4, '', part1)), pp[1]]))
                    else:
                        # Шаг 1) Вырезать предмет (перед ним может быть дата, если первый тип - лекция/практика)
                        predmet = search(pat1, divide[ind])[1]
                        for_sep = sub(predmet, '', divide[ind])
                        # Шаг 2) Строка без предметов логически делится по преподам
                        for razd in findall(pattern2a, for_sep):
                            razbor.append('; '.join([predmet, razd]))
                else:
                    razbor.append(divide[ind])

                """ Прогон по каждому логическому набору в разборном списке """
                for nabor in razbor:
                    # Вырезать предмет из набора в отдельную переменную (если предмет указан)
                    if search(pat1, nabor):
                        predmet = search(pat1, nabor)[1]
                        nabor = sub(pat1, '', nabor)
                    else: # Если предмета нет (такого быть не должно, но мало ли)
                        predmet = 'Not found'

                    """ Выделение наборов "даты - тип пары - группа" """
                    it_inf = iter([a for a in findall(pattern4, nabor) if a]) # Типологический итератор
                    dtg = [[[], [], []]] # Логический список дат/типов/групп конструкции
                    pred = 9 # Индекс предыдущей группы совпадений, нужен для остановки после комбо + фулсета
                    # Комбо, когда очередь из одинаковых типов пар
                    for trash in it_inf:
                        # Определить, что отловилось (даты - группа №0, типы - 1, группы - 2)
                        for i, unint in enumerate(trash):
                            if unint:
                                # Если не хватает подгруппы, а актуал не подгруппа, то это уже следующий набор
                                if i!=2 and dtg[-1][0] and dtg[-1][1] and not dtg[-1][2]:
                                    dtg[-1][2].append(defaults[0])
                                if not dtg[-1].count([]) and i!=pred or dtg[-1][1] and i==1: # Если фулсет и не комбо, или следующий набор
                                    dtg.append([[], [], []]) # То перейти к следующему элементу
                                dtg[-1][i].append(unint)
                                break # Нет смысла проверять остальное, если искомое нашлось
                        pred = i

                    """ Проверка последней записи """
                    if not dtg[-1][2]: # Если нет подгруппы
                        dtg[-1][2].append(defaults[0]) # Значение по умолчанию

                    """ Исправление возможных ошибок в изначальном расписании """
                    i = 0 # Опять же, длина списка может меняться, из-за чего приходится использовать while
                    while True:
                        """ Разделение записи по подгруппам, если оно возможно """
                        if len(dtg[i][2]) > 1:
                            j = len(dtg[i][2])
                            for gr in dtg[i][2]:
                                dtg = dtg[:i+1] + [[dtg[i][0], dtg[i][1], [gr]]] + dtg[i+1:]
                            del(dtg[i])
                            i += j-1

                        """ Исправление разделения записей вида "ФЗК;  практика: лаб.раб.: 15.04.20г. - 1п/гр" """
                        if not dtg[i][0]:
                            if i > 0: # Если косяк в "не первой" записи, то склеить её с предыдущей
                                dtg[i-1][1][0] += ', '+dtg[i][1][0] # Склейка типов пары
                                del(dtg[i])
                                i -= 1 # На случай, если дальше что-то будет
                            elif i+1 != len(dtg): # Если косяк в первой записи, то склеить её со следующей
                                dtg[i+1][1][0] += ', '+dtg[i][1][0] # Склейка типов пары
                                del(dtg[i])
                                i -= 1 # На случай, если дальше что-то будет

                        """ Если запись не содержит тип пары, то он либо был упомянут ранее, либо таков замысел """
                        if i>0 and not dtg[i][1]:
                            dtg[i][1] = dtg[i-1][1]   # Есть предыдущая запись? Взять тип пары из неё
                        else:
                            dtg[i][1].append(defaults[1]) # Типа нет? По умолчанию

                        """ Проверка на завершение цикла коррекции """
                        if i+2 > len(dtg): # Если все записи были обработаны, исправление ошибок завершается
                            break
                        else:              # Если ещё остались какие-то записи, продолжить коррекцию
                            i += 1

                    """ Форматирование дат, типа пары и подгруппы """
                    for f in range(len(dtg)):
                        dtg[f][0] = expand_dates(dtg[f][0], year, len(day_num)-1)
                        if not dtg[f][1]: # Хз как, но порой тип пары всё равно отсутствует
                            dtg[f][1].append(defaults[1])
                        elif dtg[f][1][0] != defaults[1]:
                            dtg[f][1] = [format_tip(dtg[f][1][0])]
                        if dtg[f][2][0] != defaults[0]:
                            dtg[f][2] = [format_group(dtg[f][2][0], g_yn)]

                    """ Выделить препода (выделяется после обособления типов пар) """
                    # Если выделить до, то при ошибке в синтаксисе изначального расписания можно поймать:
                    # Препод = "теория ст.преподаватель БелкинаА.В."
                    if findall(pat2, nabor): # Если есть препод, отловить его с должностью (если указана)
                        prepod = [format_prep(findall(pat2, nabor)[0], p_yn)]
                    else: # Если препод не указан
                        if findall(pat5, razbor[i]): # Если у текущей пары есть тип, то она не "особая"
                            prepod = [format_prep(parse[-1][3], p_yn)]
                        else: # Пара особая = мероприятие, час куратора и т.п = стандартный препод
                            prepod = [defaults[5]]

                    """ Для каждого набора "тип пары - подгруппа - даты" создать запись в БД """
                    for info in dtg:
                        # Финальная корректировка типа
                        if len(parse) and info[1][0]==defaults[1]:
                            if parse[-1][2] == predmet: # Если предмет как в предыдущем, а тип базовый, то ошибка в типе
                                info[1] = [parse[-1][4]]
                            elif info[2][0] != defaults[0]: # Если тип базовый, но есть подгруппы - сменить на базовый №2
                                info[1][0] = defaults[2]                            

                        # Есть случаи, когда номер пары - объединённая ячейка
                        if record[1]:    # Если не объединённая, и номер есть, то всё ок
                            num = record[1]
                        elif len(parse): # Если объединённая, взять предыдущий номер
                            num = parse[-1][1]
                        else:            # Если же номера нет и взять не откуда, поставить первый
                            num = 1

                        """ Занесение в базу """
                        # Кабинеты добавляются отдельно, так как с ними много проблем
                        parse.append([record[0],   # День недели  # 0
                                      num,         # Номер пары   # 1
                                      predmet,     # Предмет      # 2
                                      prepod[0],   # Препод       # 3
                                      info[1][0],  # Тип пары     # 4
                                      info[2][0],  # Подгруппа    # 5
                                      info[0],     # Даты         # 6
                                      ''           # Кабинет      # 7
                                      ])

                        """ Проверка на разброс одной логической записи по двум записям (когда криво записано в изначальном файле) """
                        # "Информатика; 01,22,29.03; 05,12.04.19г. - 2п/гр.; лаб.раб: 19,26.04; 10,17,24.05.19г. -2п/гр.: доцент ОбвинцевО.А.;"
                        if len(parse)>1 and parse[-2][:-2]==parse[-1][:-2]:
                            parse[-2][-2].extend(parse[-2][-2])
                            del(parse[-1])
                            # Удаление повторов дат
                            parse[-1][6] = sorted(list(set(parse[-1][6])))

                if ind+1 == len(divide):
                    break
                else:
                    ind += 1

            """ Логика распределения кабинетов """
            # Если кабинеты есть
            if record[-1]:
                cabs = record[-1] # Список кабинетов для строки
            # Если кабинетов нет
            else:
                cabs = [defaults[3]]
            it_cb = iter(cabs)    # Итератор кабинетов

            # Проход по всем записям в строке
            for ind in range(start, len(parse)):

                """ Подсчитать количество уникальных записей и преподов """
                n_rec = n_prp = 0
                for i in range(start, len(parse)):
                    # Если нет совпадения с предыдущей записью (без учёта подгруппы, дат, кабинета)
                    if i == 0 or parse[i-1][:5] != parse[i][:5]:
                        # То запись - уникальная
                        n_rec += 1
                        # А если нет совпадения, то больше шансов что препод отличается
                        if parse[i-1][3] != parse[i][3]:
                            n_prp += 1

                """ Расстановка кабинетов """
                # Если ячейке сопряжён только один кабинет (или в ней лишь одна запись)
                if len(cabs) == 1 or start+1 == len(parse):
                    # Если препод по умолчанию, а записей несколько
                    if parse[ind][3] == defaults[5] and n_rec > 1:
                        # Поставить кабинет по умолчанию
                        parse[ind][7] = defaults[3]
                    # Если запись одна
                    else:                        
                        # Поставить первый кабинет из списка кабинетов ячейки
                        parse[ind][7] = cabs[0]
                else:
                    # Если ФЗК для нескольких групп, то могут быть косяки
                    if ind and parse[ind-1][:3] == parse[ind][:3] and parse[ind-1][7] == defaults[4]:
                        parse[ind][7] = defaults[4]
                        continue
                    # Если число записей совпадает с числом кабинетов
                    if n_rec == len(cabs):
                        # Итерироваться при смене УНИКАЛЬНОЙ записи (без учёта подгрупп)
                        if ind == start or parse[ind-1][:5] != parse[ind][:5]:
                            """ Итерация """
                            try:
                                parse[ind][7] = next(it_cb)
                            except:
                                parse[ind][7] = parse[ind-1][7]
                        else:
                            parse[ind][7] = parse[ind-1][7]
                    else:
                        # Если преподов столько же сколько и кабинетов
                        if n_prp == len(cabs):
                            # Итерироваться при смене препода
                            if ind == start or parse[ind-1][3] != parse[ind][3]:
                                """ Итерация """
                                try:
                                    parse[ind][7] = next(it_cb)
                                except:
                                    parse[ind][7] = parse[ind-1][7]
                            else:
                                parse[ind][7] = parse[ind-1][7]
                        else:
                            # Итерироваться при одном из условий:
                            if (ind == start or
                                # 1) Сменился препод и не переход с лекцию на лекцию
                                parse[ind-1][3] != parse[ind][3] and not (parse[ind-1][4] == parse[ind][4] == tip_list[2]) or
                                # 2) Препод не сменился и переход с/на лекцию
                                parse[ind-1][3] == parse[ind][3] and (parse[ind-1][4] == tip_list[2] or parse[ind][4] == tip_list[2])):
                                """ Итерация """
                                try:
                                    parse[ind][7] = next(it_cb)
                                except:
                                    parse[ind][7] = parse[ind-1][7]
                            else:
                                parse[ind][7] = parse[ind-1][7]
                                

        # Если с парсингом всё было норм, вернуть базу запарсенного расписания
        return parse

    except:
        # Если попалась ошибка, которой не было в тестах, то в расписании был какой-то мощный косяк
        return BadDataError('Парсинг расписания не удался, где-то было встречено ядрёное исключение!')


def print_bd(bd     : 'База парсинга',
             group  : 'Группа',
             period : 'Период расписания',
             year   : 'Год'
             ) ->      None:

    """ Функция для консольного вывода базы парсинга (только для тестов) """
    parse_title = 'Расписание '+str(group)+' на '+str(year)+'-й год. Учебная часть семестра идёт с '+period[0]+' по '+period[1]
    print()
    print(f"\n{parse_title : ^188}\n")
    print('='*187)
    print(f"| {'День' : ^4} | {'№' : ^1} | {'Предмет' : ^80} | {'Препод' : ^43} | {'Тип' : ^15} | {'Для кого' : ^11} | {'Каб' : ^11} |")
    print('-'*187)
    for i, record in enumerate(bd):
        if i>0 and record[0]!=bd[i-1][0]:
            print('-'*187)
        print(f"| {record[0] : ^4} | {record[1] : ^1} | {record[2] : ^80} | {record[3] : ^43} | {record[4] : ^15} | {record[5] : ^11} | {record[7] : ^11} |")
    print('='*187)
    print()