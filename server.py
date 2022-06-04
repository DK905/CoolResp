r"""Серверная часть API для парсинга расписания УрТИСИ СибГУТИ.

Генерация зависмостей: pip freeze > requirements.txt
Восстановление зависимостей: pip install -r requirements.txt
"""

# # # Импорт общих модулей
import os
import pandas as pd
from hashlib import sha256
import datetime

# # # Импорт модулей для работы сервера
from flask import Flask, request, abort, jsonify, Response, send_file, make_response

import cr_component.api.defaults as api_def
import cr_component.api.database as api_db
# Программный планировщик заданий (если нужны действия "по таймеру")
# from apscheduler.triggers.interval import IntervalTrigger
# Функция для безопасной работы с именами файлов (если имена файлов будут использоваться в коде)
from werkzeug.utils import secure_filename

# # # Импорт модулей для работы парсера
import cr_component.parser.additional as cr_add
import cr_component.parser.defaults as cr_def
import cr_component.parser.reader as cr_read
import cr_component.parser.parser as cr_parse
import cr_component.parser.writer as cr_write


# Создание Flask-приложения
app = Flask(__name__)
# JSON-ответы не должны кодироваться в ASCII
app.config['JSON_AS_ASCII'] = False

# Базовая директория приложения
base_dir = os.path.abspath(os.path.dirname(__file__))

# Директории сохранения файлов
app.config['UPLOAD_FOLDER'] = os.path.join(base_dir, 'Temporary', 'Uploads')
app.config['JSON_FOLDER'] = os.path.join(base_dir, 'Temporary', 'JSON')
app.config['OTHER_FOLDER'] = os.path.join(base_dir, 'Temporary', 'Other')
# Создание директорий (если их нет)
for directory in [app.config['UPLOAD_FOLDER'],
                  app.config['JSON_FOLDER'],
                  app.config['OTHER_FOLDER']]:
    if not os.path.exists(directory):
        os.makedirs(directory)

# Максимально допустимый размер загружаемого файла
app.config['MAX_CONTENT_LENGTH'] = api_def.MAX_SIZE

# Путь к файлу базы данных
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///' + os.path.join(base_dir, 'Temporary', 'database.db')
# Индикатор отслеживания модификаций объектов базы данных
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False


@app.route('/read-book/', methods=['POST'])
def read_book():
    """ Функция загрузки файла на сервер """

    if 'file' not in request.files:
        abort(Response('Запрос должен содержать файл', 404))

    # Сохранение данных из запроса в переменную
    file_request = request.files['file']
    file_name = secure_filename(file_request.filename)
    file = file_request.stream.read()

    # Проверка хеша на наличие в базе данных
    file_hash = sha256()
    file_hash.update(file)
    file_hash = file_hash.hexdigest()
    file_info = api_db.get_info_from_files(file_hash)

    # Если хеш не присутствует в базе данных, то файл новый и его нужно проверить
    if not file_info:
        # Определение расширения файла по первым байтам, как XLS/XLSX/UNKNOWN
        file_ext = cr_add.check_object_extension(file)
        if file_ext == api_def.EXTENSIONS['404']:
            abort(Response('Файл должен иметь расширение «.xls» или «.xlsx»', 415))

        # Сохранение файла в постоянную память сервера
        path_save = os.path.join(app.config['UPLOAD_FOLDER'], file_hash)  # f'{file_hash}.{file_ext}'
        with open(path_save, 'wb') as saved_file:
            saved_file.write(file)

        # Занесение данных о файле в базу данных
        api_db.add_info_to_files(file_hash, file_name, file_ext)
    else:
        path_save = os.path.join(app.config['UPLOAD_FOLDER'], file_info['hash'])
        file_ext = file_info['extension']

    # Обновление времени последнего взаимодействия с файлом
    api_db.update_time_in_files(file_hash)

    # Подгрузка скачанной книги
    book = cr_read.read_book(path_save, file_ext)
    # Получение словаря листов и групп на них
    sheets = {sheet: cr_read.group_choice(cr_read.take_sheet(book, sheet))['groups_info']
              for sheet in cr_read.see_sheets(book)}

    # Возврат массива данных: хеша файла и словаря "лист: группы"
    return jsonify({
                    'file_hash': file_hash,
                    'sheets': sheets,
                    })


@app.route('/parse-book/<file_hash>/<sheet_name>/<group_name>/<result_type>', methods=['POST'])
def parse_book(file_hash, sheet_name, group_name, result_type):
    """ Функция парсинга расписания """

    # Запрос из БД информации о файле
    file_info = api_db.get_info_from_files(file_hash)

    # Если файла нет на сервере, вернуть ошибку 404
    if not file_info:
        abort(Response('Неправильный хеш!', 404))

    # Если файл на сервере есть, то обновить время последнего взаимодействия
    api_db.update_time_in_files(file_hash)

    # Если для файла нет запарсенных результатов, сгенерировать их
    json_info = api_db.get_info_from_json_on_file(file_hash, sheet_name, group_name)
    if not json_info:
        # Подгрузка файла книги
        try:
            file_ext = file_info['extension']
            book = cr_read.read_book(os.path.join(app.config['UPLOAD_FOLDER'], file_hash),
                                     file_ext)

            # Получение данных о группах на выбранном листе
            sheet = cr_read.take_sheet(book, sheet_name)
            period, year, groups, start_end = cr_read.group_choice(sheet).values()

            # Парсинг
            bd_process = cr_read.prepare(sheet, group_name, start_end)
            df = cr_parse.parser(bd_process, period, year)
            df_json = cr_add.dataframe_to_json(df)

            # Вычислить хеш виртуального JSON-файла
            json_hash = sha256()
            json_hash.update(bytes(df_json, 'utf-8'))
            json_hash = json_hash.hexdigest()

            # Сохранить виртуальный файл
            cr_add.dataframe_to_json(df, app.config['JSON_FOLDER'], json_hash)
            api_db.add_info_to_json(json_hash, file_hash, sheet_name, group_name)
        except Exception as err:
            print(err)
            abort(Response(str(err), 512))
    # Если результаты были сгенерированы ранее, то использовать их
    else:
        json_hash = json_info['hash']
        df = cr_add.json_to_dataframe(os.path.join(app.config['JSON_FOLDER'], json_hash), group_name)
        df_json = cr_add.dataframe_to_json(df)

    # Тип результата может быть json или excel
    if result_type == 'json':
        return df_json
    elif result_type == 'excel':
        # Выделение флагов для форматирования
        g2 = request.args.get('g2', default='0', type=bool)
        g3 = request.args.get('g3', default='0', type=bool)
        f1 = request.args.get('f1', default=False, type=bool)  # Сокращение названий предметов
        f2 = request.args.get('f2', default=True, type=bool)   # Сокращение должностей преподавателей
        f3 = request.args.get('f3', default=True, type=bool)   # Сокращение записи подгрупп
        f4 = request.args.get('f4', default=False, type=bool)  # Сокращение учебных корпусов кабинетов
        result = cr_write.create_resp(df, g2, g3, f1, f2, f3, f4)

        # Преобразование названия в транслитерацию, из-за используемой FLASK кодировки latin-51
        symbols = (u"абвгдеёжзийклмнопрстуфхцчшщъыьэюяАБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯ",
                   u"abvgdeejzijklmnoprstufhzcss_y_euaABVGDEEJZIJKLMNOPRSTUFHZCSS_Y_EUA")
        tr = {ord(a): ord(b) for a, b in zip(*symbols)}
        file_name = f'{cr_add.create_name(df)}.xlsx'.translate(tr)

        # Сохранение excel-книги и возврат её
        cr_write.save_resp(result, os.path.join(app.config['OTHER_FOLDER'], file_name))
        return make_response(send_file(os.path.join(app.config['OTHER_FOLDER'], file_name)))
    else:
        return Response('None', 404)


@app.route('/merge-books', methods=['POST'])
def merge_books():
    """ Функция сведения разных файлов расписания в один датафрейм """

    # Базовые параметры сведения
    merged_df = pd.DataFrame(columns=['group', *cr_def.DEF_COLUMNS])
    # year = request.args.get('year', default=datetime.date.today().year, type=int)
    year = request.args.get('year', default=0, type=int)

    # Список сводимых файлов. Если не были переданы конкретные хеши, обработать всю имеющуюся базу
    files = request.get_json(silent=True)
    if not files:
        files = os.listdir(app.config['UPLOAD_FOLDER'])

    # Получение хешей всех результатов парсинга сводимых файлов
    for file_hash in files:
        json_hashes = api_db.get_info_from_json_on_file(file_hash=file_hash)
        for json_hash in json_hashes:
            file, group = json_hash['hash'], json_hash['group']
            df = cr_add.json_to_dataframe(os.path.join(app.config['JSON_FOLDER'], file))
            df_year = int(df['date_pair'].dt.year.mean())
            if 2000 < year != df_year:
                continue

            df.insert(0, 'group', group)
            merged_df = merged_df.append(df, ignore_index=True)

    return cr_add.dataframe_to_json(merged_df)


@app.route('/get-books', methods=['GET', 'POST'])
def get_books():
    """ Функция возврата данных о всех имеющихся файлах """

    pass


def delete_file(filename):
    """ Функция удаления файла с сервера """
    os.remove(os.path.join(app.config['UPLOAD_FOLDER'], filename))

    # Удаление результатов парсинга, связанных с исходным файлом
    jsons = api_db.delete_info_from_json(filename)
    for json_hash in jsons:
        os.remove(os.path.join(app.config['JSON_FOLDER'], json_hash))

    # Удаление исходного файла и его записи, а также связанных хешей результатов
    api_db.delete_info_from_files(filename)


def check_file_lifetime(filename):
    """ Функция проверки последней активности файла на сервере """

    # Запрос из БД информации о файле
    file = api_db.get_info_from_files(filename)
    # Если в БД значится файл, вернуть разницу с последним измерением
    if file:
        time_now = api_db.now_in_prefer_timezone().replace(tzinfo=None)
        time_mod = file['modified_at']
        delta = (time_now - time_mod).total_seconds()
        return delta
    # Если файла в БД нет, его срок жизни истёк
    else:
        return api_def.FILE_LIFETIME+1


def check_directory(path):
    """ Функция удаления неактивных файлов, и файлов не отмеченных в базе данных """

    # Получение всех директорий и файлов по заданному пути
    for root, d, files in os.walk(path):
        # Проход по всем файлам в заданной директории
        for file in files:
            # Получение срока жизни текущего файла
            check = check_file_lifetime(file)
            # print(f'{check} VS {api_def.FILE_LIFETIME}')
            # Если срок жизни превышен, удалить файл
            if check > api_def.FILE_LIFETIME:
                delete_file(file)


def check_database():
    """ Функция удаления неактивных записей в БД, и записей, чьи файлы не обнаружены """

    # Получение всех записей из базы данных
    for record in api_db.FileInfo.query.all():
        file_hash = record.hash
        # Отсеивание записей, файлы к которым не обнаружены
        if not os.path.exists(os.path.join(app.config['UPLOAD_FOLDER'], file_hash)):
            api_db.delete_info_from_files(file_hash)
        # Отсеивание записей, чей срок жизни истёк
        check = check_file_lifetime(file_hash)
        if check > api_def.FILE_LIFETIME:
            delete_file(file_hash)


if __name__ == '__main__':
    # Импорт и подключение базы данных
    api_db.db.app = app
    api_db.db.init_app(app)
    api_db.db.create_all()

    # # Очистка неактивных/необозначенных файлов
    # check_directory(app.config['UPLOAD_FOLDER'])
    # # Очистка неактивных/необозначенных записей в БД
    # check_database()

    # Запустить приложение в режиме дебага, с доступом через порт 5000
    app.run(debug=True, port=5000)

    # Запустить приложение в обычном режиме, с доступом через порт заданный хостингом
    # app.run(int(os.getenv('PORT')))
