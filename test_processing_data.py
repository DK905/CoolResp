"""
Тестовый модуль чисто для отлова крашей при обработке
То есть, он может отловить лишь какие-нибудь ошибки самого процесса обработки, но не правильность результатов
P.S Возможно, он даже не отлавливает, а просто формирует файлы для ручной проверки
P.P.S Кажется, я переборщил с обработкой исключений

Для всех файлов в path_load происходит обработка каждой группы на каждом листе
Результат парсинга сохраняется в path_json
Результат форматирования сохраняется в path_save
"""

""" Подключение модулей """
from CoolRespProject.modules import CR_reader  as crr   # Считывание таблицы в базу разбора для конкретной группы
from CoolRespProject.modules import CR_parser  as crp   # Парсинг базы разбора в нормальную БД для каждой логической записи
from CoolRespProject.modules import CR_analyze as cra   # Анализатор БД, для подправки косяков и определения числа подгрупп
from CoolRespProject.modules import CR_jsoner  as crj   # Сохранение БД парсинга и БД анализа в json (для оптимизации)
from CoolRespProject.modules import CR_writter as crw   # Форматная запись БД в таблицу EXCEL
from os import walk                                     # Для прохода по файлам в папке

path_load = 'Tests/Datasets/'  # Путь загрузки обрабатываемых файлов
path_save = 'Tests/Results/'   # Путь сохранения итогов обработки
path_json = 'Tests/Jsons/'     # Путь сохранения json файлов
pdgr = (0, 0)                  # Номера выбранных подгрупп


def test_processing(file_name: 'Название тестируемого файла'
                    ) ->        None:

    # Полный путь к проверяемому файлу
    file_path = f'{path_load}{file_name}'

    # Считывание EXCEL таблицы по пути name в переменную book
    book = crr.read_book(file_path)

    # Получить список названий листов в книге
    sheets = crr.choise_sheet(book)

    # Проход по всем названиям листов в книге
    for sheet_n in range(len(sheets)):
        # Выбрать лист №sheet_n
        sheet = book[sheets[sheet_n]]

        # Получить информацию о выбранном листе
        sheet_info  = crr.choise_group(sheet)
        timey_wimey = sheet_info[0]  # Период расписания
        year        = sheet_info[1]  # Год
        groups      = sheet_info[2]  # Список групп на листе
        row_start   = sheet_info[3]  # Индекс стартовой строки

        # Проход по всем группам на листе
        for group in groups:
            name = f"Респа для {group} на [{' - '.join(timey_wimey)}]"
            print(name)

            # Создание базы с данными для предварительной обработки
            bd_process = crr.prepare(sheet, group, row_start, True)
            # for row in bd_process:
            #     print(row)

            # Создание базы с запарсенными данными
            bd_parse = crp.parser(bd_process, timey_wimey, year, True, False, True)
            # Спринтить максимальный номер пары в парсинге
            print(f'Максимальный номер пары для группы был равен {max(row[1] for row in bd_parse)}-й паре')

            # Обычный вывод строк из базы парсинга
            # for row in bd_parse:
            #     print(row)

            # Форматный вывод базы парсинга
            # crp.print_bd(bd_parse, group, timey_wimey, year)

            # Выделение базы анализа данных из парсинга
            analysis = cra.analyze_bd(bd_parse)

            # Сохранение данных в json
            crj.save_json(bd_parse, analysis, path_json, name)

            # Загрузка данных из json
            # bd_parse, analysis = crj.load_json(f'{path_json}{name}.json')

            # Получение объекта форматированного расписания
            f_book = crw.create_resp(bd_parse, analysis, pdgr[0], pdgr[1], timey_wimey, year)

            # Сохранение форматированного расписания
            crw.save_resp(f_book, f'{path_save}{name}.xlsx')
    print()


for root, d, files in walk(path_load):
    for file in files:
        print(f'Открытие файла {file}')
        test_processing(file)
