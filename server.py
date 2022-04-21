r"""Серверная часть API для парсинга расписания УрТИСИ СибГУТИ.

Генерация зависмостей: pip freeze > requirements.txt
Восстановление зависимостей: pip install -r requirements.txt
"""

# # # Импорт общих модулей
import os
from hashlib import sha256

# # # Импорт модулей для работы сервера
from flask import Flask, request, abort, jsonify, Response
from CoolRespProject.modules_api import api_defaults as api_def
from CoolRespProject.modules_api import api_database as api_db
# Программный планировщик заданий (если нужны действия "по таймеру")
# from apscheduler.triggers.interval import IntervalTrigger
# Функция для безопасной работы с именами файлов (если имена файлов будут использоваться в коде)
# from werkzeug.utils import secure_filename

# # # Импорт модулей для работы парсера
from CoolRespProject.modules_parser import cr_reader as crr
from CoolRespProject.modules_parser import cr_parser as crp
import CoolRespProject.modules_parser.cr_swiss as crs


# Создание Flask-приложения
app = Flask(__name__)
# JSON-ответы не должны кодироваться в ASCII
app.config['JSON_AS_ASCII'] = False

# Базовая директория приложения
base_dir = os.path.abspath(os.path.dirname(__file__))

# Путь к директории загрузки файлов расписания
app.config['UPLOAD_FOLDER'] = os.path.join(base_dir, 'Temporary', 'Uploads')
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
    # filename = secure_filename(file_request.filename)
    file = file_request.stream.read()

    # Проверка хеша на наличие в базе данных
    file_hash = sha256()
    file_hash.update(file)
    file_hash = file_hash.hexdigest()
    # api_db.delete_info(file_hash)
    file_info = api_db.get_info(file_hash)

    # Если хеш не присутствует в базе данных, то файл новый и его нужно проверить
    if not file_info:
        # Определение расширения файла по первым байтам, как XLS/XLSX/UNKNOWN
        file_ext = crs.check_object_extension(file)
        # Если расширение файла не поддерживается, вернуть код 415 - неподдерживаемый сервером медиа-формат
        if file_ext == api_def.EXTENSIONS['404']:
            abort(Response('Файл должен иметь расширение .xls или .xlsx', 415))

        # Сохранение файла в постоянную память сервера
        path_save = os.path.join(app.config['UPLOAD_FOLDER'], file_hash)  # f'{file_hash}.{file_ext}'
        with open(path_save, 'wb') as saved_file:
            saved_file.write(file)

        # Занесение данных о файле в базу данных
        api_db.add_info(file_hash, file_ext)
    else:
        path_save = os.path.join(app.config['UPLOAD_FOLDER'], file_info['hash'])
        file_ext = file_info['extension']

    # Подгрузка скачанной книги
    book = crr.read_book(path_save, file_ext)
    # Получение словаря листов и групп на них
    # sheets = {sheet: crr.group_choice(crr.take_sheet(book, sheet))['groups_info'] for sheet in crr.see_sheets(book)}
    sheets = {}
    for sheet in crr.see_sheets(book):
        sheets[sheet] = crr.group_choice(crr.take_sheet(book, sheet))['groups_info']

    # Обновление времени последнего взаимодействия с файлом
    api_db.update_time(file_hash)

    # Возврат массива данных: хеша файла и словаря "лист: группы"
    return jsonify({
                    'file_hash': file_hash,
                    'sheets': sheets,
                    })


@app.route('/parse-book/<file_hash>/<sheet_name>/<group_name>', methods=['POST'])
def parse_book(file_hash, sheet_name, group_name):
    """ Функция парсинга расписания """

    # Запрос из БД информации о файле
    file_info = api_db.get_info(file_hash)

    # Если файла нет на сервере, вернуть ошибку 404
    if not file_info:
        abort(Response('Неправильный хеш!', 404))

    # Если файл на сервере есть, то обновить время последнего взаимодействия
    api_db.update_time(file_hash)
    # Выделение расширения файла из запроса к БД
    file_ext = file_info['extension']
    # # Выделение флагов для форматирования
    # f1 = request.args.get('f1', default=False, type=bool)
    # f2 = request.args.get('f2', default=False, type=bool)
    # f3 = request.args.get('f3', default=False, type=bool)
    # f4 = request.args.get('f4', default=False, type=bool)
    # print(f1, f2, f3, f4)

    # Подгрузка книги по хешу
    book = crr.read_book(os.path.join(app.config['UPLOAD_FOLDER'], file_hash),
                         file_ext)
    # Получение данных о нужной группе на нужном листе
    sheet = crr.take_sheet(book, sheet_name)
    period, year, groups, start_end = crr.group_choice(sheet).values()

    # Выделение
    bd_process = crr.prepare(sheet, group_name, start_end)

    # Парсинг
    df = crp.parser(bd_process, period, year)

    # f_book = crw.create_resp(df, g2, g3)
    # crw.save_resp(f_book, f'Tests/Results/{crs.create_name(df)}.xlsx')

    return df.to_json(orient='records',    # Стиль JSON файла
                      indent=4,            # Уровень отступов внутри файла
                      force_ascii=False,   # Запись в ASCII?
                      date_format='iso')   # Формат записи даты


def delete_file(filename):
    """ Функция удаления файла с сервера """

    # Удаление файла по заданному пути
    os.remove(os.path.join(app.config['UPLOAD_FOLDER'], filename))
    # Удаление из БД записи о файле
    api_db.delete_info(filename)


def check_file_lifetime(filename):
    """ Функция проверки последней активности файла на сервере """

    # Запрос из БД информации о файле
    file = api_db.get_info(filename)
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
            api_db.delete_info(file_hash)
        # Отсеивание записей, чей срок жизни истёк
        check = check_file_lifetime(file_hash)
        if check > api_def.FILE_LIFETIME:
            delete_file(file_hash)


if __name__ == '__main__':
    # Импорт и подключение базы данных
    api_db.db.app = app
    api_db.db.init_app(app)
    api_db.db.create_all()
    # print(api_db.FileInfo.query.all())

    # Очистка неактивных/необозначенных файлов
    check_directory(app.config['UPLOAD_FOLDER'])
    # Очистка неактивных/необозначенных записей в БД
    check_database()

    # Запустить приложение в режиме дебага, с доступом через порт 5000
    app.run(debug=True, port=5000)

    # Запустить приложение в обычном режиме, с доступом через порт 5000
    # app.run(debug=False, port=5000)
