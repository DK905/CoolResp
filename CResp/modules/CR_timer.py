""" Информация о модуле
Данный модуль предназначен для выборки расписания на определённый день недели

"""
# Импорт умолчаний и сокращений для типов пар
try:
    from modules.CR_dataset import BadDataError, days_names, time_budni, time_vihod
except:
    from CR_dataset import BadDataError, days_names, time_budni, time_vihod

# Импорт стандартной даты из модуля темпоральной обработки
from datetime import date as dt_date, timedelta

def what_resp(p_bd : 'БД с запарсенной инфой о респе',
              a_bd : 'БД с инфой о распределении подгрупп по предметам',
              c_dt : 'Выбранная дата, для которой выводится респа',
              grp2 : 'Выбранная подгруппа (где две подгруппы)',
              grp3 : 'Выбранная подгруппа (где три подгруппы)',
              ) ->   'Мини-БД расписания на выбранную дату':

    """ Функция вывода расписания для указанного дня """
    try:
        # Предварительное задание мини-БД на выбранную дату
        info = [[] for para in range(7)]
        # Включая краткую текстовую запись нужного дня
        info[0] = days_names[c_dt.isoweekday()]

        # Поиск по каждой записи в запарсенном...
        for rec in p_bd:
            # ...а также, каждого дня в записи
            for day in rec[6]:
                # И если день совпадает с выбранным...
                if day == c_dt:
                    # ...проверить соответствие подгрупп
                    n_act = a_bd[rec[2]][rec[4]][0]
                    if n_act==1 or n_act==2 and rec[5]==grp2 or n_act==3 and rec[5]==grp3:
                        info[rec[1]] = [rec[2], rec[3], rec[4], rec[7]]
                        print(info[rec[1]])

        if info[1:] == [[] for para in range(6)]:
            return 'Пар нет'
        else:
            return info

    except:
        # Если какой-то косяк в обработке
        return BadDataError('Упс, что-то не так с вашими данными...')

if __name__ == '__main__':
    import CR_jsoner as crj

    p_bd, a_bd = crj.load_json(f'E:/CoolResp/CResp/Data/data_ПЕ-81б.json')

    what_resp(p_bd, a_bd, dt_date.today() + timedelta(days = 0), 0, 0)
