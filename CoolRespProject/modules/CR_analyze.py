""" Информация о модуле
Данный модуль предназначен для анализа запарсенной базы расписания.

Под анализом понимается простой процесс определения количества подгрупп каждого типа пары для каждого предмета.
Можно было бы обойтись и без анализа, но он позволяет:
    1) Определять ошибки при составлении расписания (когда у 1-й подгруппы есть только практики, а у 2-й - только лабы)
    2) Составлять форматированное расписание только для своих подгрупп (что может быть круче персонального расписания?)

"""

# Импорт умолчаний
from CoolRespProject.modules.CR_dataset import defaults


def analyze_bd(parse: 'БД с запарсенной инфой'
               ) ->   'Словарь: ключ - предмет, значение - [кол-во подгрупп типа пары, замена типа пары при ошибках]':

    """ Функция анализа базы парсинга """

    # Составление основы начальной базы анализа подгрупп
    # Здесь ключ - предмет, а значение - список наборов "тип пары - подгруппы"
    a_bd = {}
    for record in parse:
        if record[2] not in a_bd:
            a_bd[record[2]] = [[record[4], record[5]]]
        else:
            if [record[4], record[5]] not in a_bd[record[2]]:
                a_bd[record[2]].append([record[4], record[5]])

    # Переработка базы анализа подгрупп
    for predmet in sorted(a_bd.keys()):
        # Словари хороших, плохих и временных данных
        good, bad, temp = {}, {}, {}
        for info in sorted(a_bd[predmet]):
            if info[0] not in temp:
                temp[info[0]] = [info[1]]
            else:
                temp[info[0]].append(info[1])

        # Прогон каждого набора 'Тип пары': [подгруппы]
        for info in temp.keys():
            # Если подгруппа - общее умолчание, то всё хорошо
            if temp[info] == [defaults[0]]:
                good[info] = temp[info]

            # Если подгруппы есть, то среди них могут быть пропуски
            # Пропуски возникают, когда в расписании лабу случайно обозвали практикой (и т.п)
            elif len(temp[info]) > 1:
                # Наличие пропусков, максимальный индекс подгруппы (для пропусков)
                all_ok, n_group = True, 0

                # Прогон по всем подгруппам
                for g_ind, grp in enumerate(temp[info], start = 1):
                    # Если где-то номер подгруппы не совпадает с тем что должен быть, всё плохо
                    if grp != defaults[0] and g_ind != int(grp[0]):
                        if info not in bad:
                            bad[info] = []
                        bad[info].append(grp)
                        all_ok = False

                if all_ok:
                    good[info] = temp[info]

            # Если есть подгруппы, но в списке всего одна - что-то пошло не так
            else:
                bad[info] = [temp[info][0]]

        # Если есть как хорошие, так и плохие типы пар
        if bad and good:
            # Прогон каждого плохого типа пары
            for ind_bad in bad:
                # Типы пар, которые могут юзаться как заменитель
                maybe = {k: v for k, v in good.items() if v != [defaults[0]]}

                # Если есть ровно один заменитель, добавить к нему возможные подгруппы
                # Занести заменитель как значение плохиша
                if len(maybe) == 1:
                    for val in bad[ind_bad]:
                        if val not in good[list(maybe.keys())[0]]:
                            good[list(maybe.keys())[0]].append(val)

                    bad[ind_bad] = list(maybe.keys())[0]

                # Если заменителя в хорошем нет, то он может появиться в плохом
                elif not len(maybe):
                    good[ind_bad] = bad[ind_bad]

                # Если потенциальных заменителей несколько, выбрать тот что ближе соответствует подгруппам
                else:
                    # Выделить максимальную подгруппу в плохише
                    n_group = max(list(map(int, (a[0] for a in bad[ind_bad]))))
                    all_ok, zam = False, False
                    # Точный подбор заменителя
                    for mb in maybe:
                        # Проверка всех подгрупп
                        for g in maybe[mb]:
                            if int(g[0]) == n_group:
                                zam, all_ok = mb, True
                                break
                    if not all_ok:
                        # Примерный подбор заменителя (если точный не помог)
                        for mb in maybe:
                            if n_group in range(int(g[0]) - 1, int(g[0]) + 2):
                                zam, all_ok = mb, True
                                break
                    good[zam].extend(ind_bad.items()[0][1])
                    for val in bad[ind_bad]:
                        if val not in good[zam]:
                            good[zam].append(val)
                    bad[ind_bad] = zam

        # Если хорошего списка нет, то попалось кривое исключение
        elif bad:
            good = bad
            bad = {}

        # После обработки ошибок, нужно переработать temp как 'Ключ': [Кол-во подгрупп, заменитель]
        for t in temp:
            # Фикс лишнего добавления. Делается здесь, чтобы избежать ошибок итерирования по словарю bad
            if t in bad and t in good:
                bad.pop(t)
            # Если ключа нет в плохих, или он есть и там и там, то всё ок
            if t in [list(bad.keys()), list(good.keys())]:
                temp[t] = [len(good[t]), t]
            # Если ключа нет в хороших, то взять кол-во подгрупп из плохих, заменитель из него же
            elif t in list(bad.keys()):
                temp[t] = [len(temp[t]), bad[t]]
            # Если ключа нет в плохих, то заменитель не нужен => вместо него ставится ключ
            else:
                temp[t] = [len(good[t]), t]
        a_bd[predmet] = temp

    # Если словарь анализа был составлен, вернуть его
    return a_bd
