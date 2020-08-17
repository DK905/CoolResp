""" Информация о модуле
Данный модуль предназначен для анализа запарсенной базы расписания.

Под анализом понимается простой процесс определения количества подгрупп каждого типа пары для каждого предмета.
Можно было бы обойтись и без анализа, но он позволяет:
    1) Определять ошибки при составлении расписания (когда у 1-й подгруппы есть только практики, а у 2-й - только лабы)
    2) Составлять форматированное расписание только для своих подгрупп (что может быть круче персонального расписания?)

"""

# Импорт умолчаний
from modules.CR_dataset import BadDataError, defaults


def analyze_bd(parse : 'БД с запарсенной инфой'
               ) ->    'Словарь: ключ - предмет, значение - [кол-во подгрупп типа пары, замена типа пары при ошибках]':

    """ Функция анализа базы парсинга """
    try:
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
                    zbs, mx = True, 0 # Наличие пропусков, максимальная подгруппа (для пропусков)

                    # Прогон по всем подгруппам
                    for g_ind, grp in enumerate(temp[info], start = 1):
                        # Если где-то номер подгруппы не совпадает с тем что должен быть, всё плохо
                        if grp != defaults[0] and g_ind != int(grp[0]):
                            if info not in bad:
                                bad[info] = []
                            bad[info].append(grp)
                            zbs = False
                    if zbs == True:
                        good[info] = temp[info]

                # Если есть подгруппы, но в списке всего одна - что-то пошло не так
                else:
                    bad[info] = [temp[info][0]]

            # Если есть как хорошие, так и плохие типы пар
            if bad and good:
                # Прогон каждого плохого типа пары
                for b in bad:
                    maybe = {k: v for k, v in good.items() if v != [defaults[0]]} # Типы пар, которые могут юзаться как заменитель

                    # Если есть ровно один заменитель, добавить к нему возможные подгруппы. Занести заменитель как значение плохиша
                    if len(maybe) == 1:
                        for val in bad[b]:
                            if val not in good[list(maybe.keys())[0]]:
                                good[list(maybe.keys())[0]].append(val)
                        bad[b] = list(maybe.keys())[0]

                    # Если заменителя в хорошем нет, то мб он появится в плохом
                    elif not len(maybe):
                        good[b] = bad[b]

                    # Если потенциальных заменителей несколько, выбрать тот что ближе соответствует подгруппам
                    else:
                        # Выделить максимальную подгруппу в плохише
                        mx = max(list(map(int, (a[0] for a in bad[b]))))
                        zbs, zam = False, False
                        for mb in maybe: # Точный подбор заменителя
                            for g in maybe[mb]: # Проверка всех подгрупп
                                if int(g[0]) == mx:
                                    zam, zbs = mb, True
                                    break
                        if not zbs:
                            for mb in maybe: # Примерный подбор заменителя (если точный не помог)
                                if mx in range(int(g[0])-1, int(g[0])+2):
                                    zam, zbs = mb, True
                                    break
                        good[zam].extend(b.items()[0][1])
                        for val in bad[b]:
                            if val not in good[zam]:
                                good[zam].append(val)
                        bad[b] = zam

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

    except:
        # Проблем быть не должно, но если и были - вернуть ошибку
        return BadDataError('Анализ расписания не удался, что-то пошло не так!')