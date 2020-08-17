""" Информация о модуле
Данный модуль предназначен для работы с БД, сохранёнными в json
То есть, для сохранения/открытия запарсенной БД и БД анализа в/из json

"""
# Что нужно сделать:
# Добавить запись/считывание базы парсинга, продумать совмещение с базой анализа

# Импорт умолчаний
from modules.CR_dataset import BadDataError
from json import dump as jdump, load as jload

def save_json(p_bd  : 'БД с запарсенной инфой',
              a_bd  : 'БД с инфой об анализе БД парсинга',
              group : 'Название группы (для удобства)'
              ) ->     None:
    try:
        # Тестовая запись БД анализа
        with open(f'data_{group}.json', 'w', encoding='utf-8') as write_bd:
            jdump(a_bd, write_bd, indent=4, ensure_ascii=False)
    except:
        # Если проблемы при сохранении, дропнуть ошибку
        return BadDataError('Сохранение БД не удалось!')

def load_json(file : 'Путь к json файлу с БД парсинга + анализа'
              ) ->   'БД парсинга, БД анализа':
    try:
        # Тестовое чтение БД анализа
        with open(file, 'r', encoding='utf-8') as read_bd:
            return jload(read_bd)
    except:
        # Если проблемы при сохранении, дропнуть ошибку
        return BadDataError('Загрузка БД не удалась!')
