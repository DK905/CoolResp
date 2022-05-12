r"""Файловая база данных API и функции взаимодействия с ней.

"""

# # # Импорт общих модулей
import os

# # # Импорт модулей работы со временем
from datetime import datetime
import pytz

# # # Импорт модулей для работы сервера
from flask_sqlalchemy import SQLAlchemy
from cr_component.api import defaults as api_def


# Объявление экземпляра класса для взаимодействия с базой данных
db = SQLAlchemy()


def now_in_prefer_timezone():
    """ Функция получения актуального времени для выбранного часового пояса """

    return datetime.now(pytz.timezone(api_def.TIMEZONE))


""" Таблица исходных файлов """


class FileInfo(db.Model):
    """ Класс-таблица для хранения данных о файлах на сервере """

    # Название таблицы задаётся через __tablename__
    __tablename__ = 'files'
    # Столбец для хранения хеша файла. Первичный (хеши разных файлов не повторяются*), строка на 64 символа
    # *совпадение хешей крайне маловероятно, но возможно
    hash = db.Column(db.String(64), primary_key=True)
    # Столбец для хранения первоначального названия файла (для удобства)
    name = db.Column(db.Text(), nullable=False)
    # Столбец для хранения расширения файла. Обязательный, строка на 4 символа
    extension = db.Column(db.String(4), nullable=False)
    # Столбец для хранения даты загрузки файла
    uploaded_at = db.Column(db.DateTime(timezone=True),
                            default=now_in_prefer_timezone(),
                            nullable=False)
    # Столбец для хранения даты последнего взаимодействия с файлом. Обязательный, дата в выбранном часовом поясе
    modified_at = db.Column(db.DateTime(timezone=True),
                            default=now_in_prefer_timezone(),
                            onupdate=now_in_prefer_timezone(),
                            nullable=False)


def get_info_from_files(file_hash, as_dict=True):
    """ Функция получения из БД информации об исходном файле """

    # Запрос к таблице исходных файлов. Результат - запись с hash=file_hash
    query_result = FileInfo.query.get(file_hash)
    # Если в БД был нужный хеш, и результат нужно представить в формате словаря
    if query_result and as_dict:
        # Выделить из объекта запроса нужные данные в словарь
        return {
            'hash': query_result.hash,
            'name': query_result.name,
            'extension': query_result.extension,
            'uploaded_at': query_result.uploaded_at,
            'modified_at': query_result.modified_at,
        }
    else:
        # Вернуть объект запроса (может быть None, что нужно для условной обработки)
        return query_result


def add_info_to_files(file_hash, file_name, file_extension):
    """ Функция добавления информации о загруженном файле в БД таблицу хешей исходных файлов """

    # Информация вносится только в случае отсутствия хеша в БД
    if not get_info_from_files(file_hash):
        file_info = FileInfo(hash=file_hash,
                             name=file_name,
                             extension=file_extension)
        db.session.add(file_info)
        db.session.commit()


def update_time_in_files(file_hash):
    """ Функция обновления времени последнего взаимодействия с исходным файлом """

    # Получить из БД информацию о хеше исходного файла в формате объекта записи
    file_info = get_info_from_files(file_hash, as_dict=False)
    if file_info:
        # Если в БД содержится запрашиваемый хеш, обновить время взаимодействия
        file_info.modified_at = now_in_prefer_timezone()
        db.session.add(file_info)
        db.session.commit()


def delete_info_from_files(file_hash):
    """ Функция удаления информации об исходном файле из БД таблицы хешей исходных файлов """

    # Получить из БД информацию о хеше, но в формате объекта записи
    file_info = get_info_from_files(file_hash, as_dict=False)
    if file_info:
        # Если в БД содержится запрашиваемый хеш, удалить его запись
        db.session.delete(file_info)
        db.session.commit()
        delete_info_from_json(file_hash)


""" Таблица JSON-результатов """


class JSONInfo(db.Model):
    """ Класс-таблица для хранения данных о расписании в формате JSON-файлов на сервере """

    # Название таблицы задаётся через __tablename__
    __tablename__ = 'json'
    # Столбец для хранения хеша JSON-файла. Первичный (хеши разных файлов не повторяются*), строка на 64 символа
    # *совпадение хешей крайне маловероятно, но возможно
    hash = db.Column(db.String(64), primary_key=True)
    # Столбец для хранения ссылки на хеш исходного файла
    hash_root = db.Column(db.String(64), db.ForeignKey(FileInfo.hash), nullable=False)
    # Столбец для хранения названия листа, откуда было взято расписание
    sheet = db.Column(db.String(40), nullable=False)
    # Столбец для хранения названия группы, для которой парсилось расписание
    group = db.Column(db.String(40), nullable=False)


def get_info_from_json(json_hash):
    """ Функция """

    query_result = JSONInfo.query.get(json_hash)
    return query_result


def get_info_from_json_on_file(json_hash='', file_hash='', sheet='', group=''):
    """ Функция получения из БД информации о JSON-результате парсинга по вторичным параметрам """

    # Если запись используется в других запросах
    if json_hash:
        return get_info_from_json(json_hash)

    # Если нужно получить информацию, осуществить запрос с последовательной фильтрацией
    if file_hash:
        query_result = JSONInfo.query.filter(JSONInfo.hash_root == file_hash)
    else:
        query_result = JSONInfo.query.all()
    if sheet:
        query_result = JSONInfo.query.filter(JSONInfo.sheet == sheet)
    if group:
        query_result = JSONInfo.query.filter(JSONInfo.group == group)

    result = [{'hash': q.hash, 'hash_root': q.hash_root,
               'sheet': q.sheet, 'group': q.group} for q in query_result]
    return result


def add_info_to_json(json_hash, root_hash, sheet, group):
    """ Функция добавления в БД информации о JSON-результате парсинга """

    # Информация вносится только в случае отсутствия хеша в БД
    if not get_info_from_json(json_hash):
        json_info = JSONInfo(hash=json_hash,
                             hash_root=root_hash,
                             sheet=sheet,
                             group=group)
        db.session.add(json_info)
        db.session.commit()


def delete_info_from_json(file_hash, is_file=True):
    """ Функция удаления информации об исходном файле из БД таблицы хешей исходных файлов """

    jsons = []
    if is_file:
        json_info = JSONInfo.query.filter(JSONInfo.hash_root == file_hash).all()
        for json_hash in json_info:
            js_info = delete_info_from_json(json_hash, False)
            jsons.append(js_info.hash)
        return jsons
    else:
        json_info = get_info_from_json_on_file(file_hash.hash)
        if json_info:
            db.session.delete(json_info)
            db.session.commit()
            return json_info


# Если запуск производится из данного файла, то это тестовый запуск
if __name__ == '__main__':
    from flask import Flask
    app = Flask(__name__)

    # Путь к файлу с базой данных
    base_dir = os.path.abspath(os.path.dirname(__file__))
    app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///' + os.path.join(base_dir, 'Temporary', 'database.db')
    # Индикатор отслеживания модификации
    app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

    db.app = app
    db.init_app(app)
    db.create_all()

    test_hash = '19e390e4fd80a2b80f448e87289087af308324307b6314e84e3085c1e27383b4'
    print(get_info_from_files(test_hash))

    delete_info_from_files(test_hash)
    print(get_info_from_files(test_hash))

    add_info_to_files(test_hash, test_hash, 'xlsx')
    print(get_info_from_files(test_hash))

    update_time_in_files(test_hash)
    print(get_info_from_files(test_hash))

    # print(FileInfo.query.all())
