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
from CoolRespProject.modules import CR_defaults as crd  # Умолчания
from CoolRespProject.modules import CR_parser   as crp  # Парсинг базы разбора в БД респы для каждой логической записи
from CoolRespProject.modules import CR_reader   as crr  # Считывание таблицы в базу разбора для конкретной группы
from CoolRespProject.modules import CR_swiss    as crs  # Швейцарский нож
from CoolRespProject.modules import CR_writter  as crw  # Модуль-форматтер
import os
import re
import xlrd
import pandas as pd
import numpy as np
pd.set_option('display.max_colwidth', None)
# pd.set_option('display.max_rows', 905)


path_load = 'Tests/Datasets/'  # Путь загрузки обрабатываемых файлов
path_save = 'Tests/Results/'   # Путь сохранения итогов обработки
path_json = 'Tests/Jsons/'     # Путь сохранения json файлов
path_best = 'Tests/JSBest/'    # Где хранятся идеальные json файлы


# Конвертация базы парсинга в датафрейм
def parse_convert_to_dataframe(bd_parse):
    df = pd.DataFrame(data=bd_parse,
                      columns=['day',        # День недели
                               'num',        # Номер пары
                               'item_name',  # Название предмета
                               'teacher',    # Препод
                               'type',       # Тип пары
                               'pdgr',       # Подгруппа
                               'date_pair',  # Дата
                               'cab'         # Кабинет
                              ]).explode('date_pair').reset_index(drop=True).drop_duplicates()

    # Выделение реального периода расписания
    date_min = df['date_pair'].dropna().min().strftime('%d.%m.%Y')
    date_max = df['date_pair'].dropna().max().strftime('%d.%m.%Y')

    # Имя датафрейма
    df.name = f'Респа для {group} на [{date_min} - {date_max}]'
    
    return df


rc = rc_act = rc_best = difs_cab_count = difs_all_count = 0
g2, g3 = '0', '0'

for root, d, files in os.walk(path_load):
    for file in files:
        print(f'Открытие файла {file}')
        
        # Полный путь к проверяемому файлу
        file_path = f'{path_load}{file}'

        # Считывание EXCEL таблицы по пути name в переменную book
        book = crr.read_book(file_path)

        # Получить список названий листов в книге
        sheets = crr.see_sheets(book)

        # Проход по всем названиям листов в книге
        for sheet_name in sheets:
            # Выбрать лист №sheet_n
            sheet = crr.take_sheet(book, sheet_name)

            # Получить информацию о выбранном листе
            sheet_info  = crr.choise_group(sheet)
            timey_wimey = sheet_info[0]  # Период расписания
            year        = sheet_info[1]  # Год
            groups      = sheet_info[2]  # Список групп на листе
            start_end   = sheet_info[3]  # Диапазон расписания

            # Проход по всем группам на листе
            for group in groups:
                # Создание базы с данными для предварительной обработки
                bd_process = crr.prepare(sheet,     # Лист расписания
                                         group,     # Выбранная группа
                                         start_end  # Диапазон расписания
                                        )

                # Создание базы с запарсенными данными
                bd_parse = crp.parser(bd_process,   # БД предобработки
                                      timey_wimey,  # Формальный период расписания
                                      year          # Год расписания
                                     )
                
                # Формирование полноценной базы данных расписания
                df = bd_parse
                
#                 # Вывод строк с пропусками 
#                 # df[df.isna().any(axis=1)]
                
                # Сравнение
                df.name = crs.create_name(df)
                df.name
                
                # Сохранение в JSON
                crs.database_to_json(df, path_json)
            
                # Сохранение в .XSLX
                f_book = crw.create_resp(df, g2, g3)
                crw.save_resp(f_book, f'Tests/Results/{df.name}.xlsx')
                
                path_to_act = f'{path_json}{df.name}.json'
                path_to_best_if = f'{path_best}{df.name}.json'
                if os.path.exists(path_to_best_if):
                    # Количество записей в датафрейме
                    rc += df.shape[0]
                    
                    # Текущий результат
                    # print(path_to_act)
                    df1 = pd.read_json(path_to_act)
                    df1['date_pair'] = pd.to_datetime(df1['date_pair']).dt.strftime('%Y-%m-%d')
                    rc_act += df1.shape[0]
                    
                    # "Идеальный" результат с кабинетами
                    # print(path_to_best_if)
                    df2 = pd.read_json(path_to_best_if).explode('date_pair').reset_index(drop=True)
                    df2['date_pair'] = pd.to_datetime(df2['date_pair']).dt.strftime('%Y-%m-%d')
                    rc_best += df2.shape[0]
                    
                    # Проверка результата
                    difs = crs.df_differences(df1, df2)
                    if not difs.empty:
                        difs_all_count += difs[difs['_merge'] == 'right_only'].shape[0]
#                         save_bd = bd_process
#                         save_difs = difs
#                         save_difs.merge(save_bd, how='left')
                    
                    
                    # Текущий результат
                    # print(path_to_act)
                    df1 = pd.read_json(path_to_act).drop(['cab'], axis=1)
                    df1['date_pair'] = pd.to_datetime(df1['date_pair']).dt.strftime('%Y-%m-%d')
                    
                    # "Идеальный" результат
                    # print(path_to_best_if)
                    df2 = pd.read_json(path_to_best_if).drop(['cab'], axis=1).explode('date_pair').reset_index(drop=True)
                    df2['date_pair'] = pd.to_datetime(df2['date_pair']).dt.strftime('%Y-%m-%d')
                    
                    # Проверка результата
                    difs = crs.df_differences(df1, df2)
                    if not difs.empty:
                        difs_cab_count += difs[difs['_merge'] == 'right_only'].shape[0]
#                         save_bd = bd_process
#                         save_difs = difs
#                         save_difs.merge(save_bd, how='left')
                else:                    
                    # файл не существует
                    print(f'Лучшего результата нет ({df.name})')
                
                
print(f'\nВсего выявлено: {rc} записей  ')
print(f'Всего сохранено: {rc_act} записей  ')
print(f'Всего должно быть: {rc_best} записей  ')

print(f'\nВсего отличий: {difs_all_count - difs_cab_count} записей отличаются от правильных кабинетами  ')
print(f'Всего отличий: {difs_cab_count} записей отличаются от правильных чем-то ещё  ')

diffs_full = difs_all_count / rc_best
diffs_cabs = (difs_all_count  / rc_best) - (difs_cab_count / rc_best)
print(f'\nПроцент коррапта: {diffs_cabs:.2%}  из-за кабинетов  ')
print(f'Процент коррапта: {diffs_full - diffs_cabs:.2%}  из-за прочего  ')

print(f'\nТочность: {1 - diffs_cabs:.2%}  с учётом кабинетов  ')
print(f'Точность: {1 - (diffs_full - diffs_cabs):.2%}  без учёта кабинетов  ')
