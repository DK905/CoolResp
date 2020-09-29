""" Информация о модуле
Данный модуль предназначен для работы с БД, сохранёнными в json
То есть, для сохранения/открытия запарсенной БД и БД анализа в/из json

"""

# Импорт умолчаний
try:
    from modules.CR_dataset import BadDataError
except:
    from CR_dataset import BadDataError

# Задание директории для сохранения данных программы
data_dyr = 'Data'

# Импорт команд обработки json файлов
from json import dump as jdump, load as jload

# Импорт команд взаимодействия с файловой системой
from os import mkdir, path

# Импорт стандартной даты из модуля темпоральной обработки
from datetime import date as dt_date

def save_json(p_bd  : 'БД с запарсенной инфой',
              a_bd  : 'БД с инфой об анализе БД парсинга',
              group : 'Название группы (для удобства)'
              ) ->     None:

    """ Функция сохранения БД парсинга и БД анализа в json формате """
    try:
        # Замена дат на строковое представление
        p_bd = list(zip(*p_bd))
        p_bd[6] = [[f'{dt:%d.%m.%y}' for dt in rec] for rec in p_bd[6]]
        p_bd = [list(row) for row in zip(*p_bd)]

        # Проверка существования директории настроек
        if not path.isdir(data_dyr):
            mkdir(data_dyr)

        # Запись БД парсинга и БД анализа как единой структуры из двух элементов
        with open(f'{data_dyr}/data_{group}.json', 'w', encoding='utf-8') as write_bd:
            jdump([p_bd, a_bd], write_bd, indent=4, ensure_ascii=False)
    except:
        # Если проблемы при сохранении, дропнуть ошибку
        return BadDataError('Сохранение БД не удалось!')

def load_json(file : 'Путь к json файлу с БД парсинга + анализа'
              ) ->   'БД парсинга, БД анализа':

    """ Функция загрузки БД парсинга и БД анализа из json формата """
    try:
        # Тестовое чтение БД анализа
        with open(file, 'r', encoding='utf-8') as read_bd:
            loaded = jload(read_bd)
            # В процессе, восстанавливаются даты как объекты
            p_bd, a_bd = list(zip(*loaded[0])), loaded[1]
            p_bd[6] = [[dt_date(2000+int(dt[6:8]), int(dt[3:5]), int(dt[0:2])) for dt in rec] for rec in p_bd[6]]
            p_bd = [list(row) for row in zip(*p_bd)]
            return p_bd, a_bd
    except:
        # Если проблемы при сохранении, дропнуть ошибку
        return BadDataError('Загрузка БД парсинга не удалась!'), BadDataError('Загрузка БД анализа не удалась!')
