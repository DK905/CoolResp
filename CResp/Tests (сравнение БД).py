# Сравнение двух версий json баз парсинга
from json import load as jload
from os import listdir

path1 = 'E:/CoolResp/CResp/Data1'
path2 = 'E:/CoolResp/CResp/Data2'

f1 = listdir(path1)
f2 = listdir(path2)
files = [f for f in f1 + f2 if f in f1 and f in f2]

for file in files:
    f1, f2 = f'{path1}/{file}', f'{path2}/{file}'
    with open(f1, 'r', encoding='utf-8') as r1, open(f2, 'r', encoding='utf-8') as r2:
        f1, f2 = jload(r1)[0], jload(r2)[0]

    t = [r for r in f1 + f2 if r not in f1 or r not in f2]
    print(file)
    if t:
        for rec in t:
            print(rec)
    else:
        print('Изменений нет')
    print()
