r"""Файловая база данных API и функции взаимодействия с ней.

"""

# # # Импорт общих модулей
import os

# # # Импорт модулей работы со временем
from datetime import datetime
import pytz

# # # Импорт модулей для работы сервера
from flask_sqlalchemy import SQLAlchemy
from CoolRespProject.modules_api import api_defaults as api_def


# Объявление экземпляра класса для взаимодействия с базой данных
db = SQLAlchemy()


def now_in_prefer_timezone():
    """ Функция получения актуального времени для выбранного часового пояса """

    return datetime.now(pytz.timezone(api_def.TIMEZONE))


class FileInfo(db.Model):
    """ Класс-таблица для хранения данных о файлах на сервере """

    # Название таблицы задаётся через __tablename__
    __tablename__ = 'files'
    # Столбец для хранения хеша файла. Первичный (хеши разных файлов не повторяются*), строка на 64 символа
    # *совпадение хешей крайне маловероятно, но возможно
    hash = db.Column(db.String(64), primary_key=True)
    # Столбец для хранения расширения файла. Обязательный, строка на 4 символа
    extension = db.Column(db.String(4), nullable=False)
    # Столбец для хранения даты последнего взаимодействия с файлом. Обязательный, дата в выбранном часовом поясе
    modified_at = db.Column(db.DateTime(timezone=True),
                            default=now_in_prefer_timezone(),
                            onupdate=now_in_prefer_timezone(),
                            nullable=False)


def get_info(file_hash, as_dict=True):
    """ Функция получения из БД информации о файле """

    # Запрос к базе данных, в формате ORM (Object-Relational Mapping). Результат - запись с hash=file_hash
    query_result = FileInfo.query.get(file_hash)
    # Если в БД был нужный хеш, и результат нужно представить в формате словаря
    if query_result and as_dict:
        # Выделить из объекта запроса нужные данные в словарь
        return {'hash': query_result.hash,
                'extension': query_result.extension,
                'modified_at': query_result.modified_at}
    else:
        # Вернуть объект запроса (даже если это None)
        return query_result


def add_info(file_hash, file_extension):
    """ Функция добавления информации о файле в базу данных """

    # Если в БД нет информации о хеше обрабатываемого файла
    if not get_info(file_hash):
        # Создать объект, имитирующий запись в БД
        file_info = FileInfo(hash=file_hash,
                             extension=file_extension)
        # Внести в БД созданный объект
        db.session.add(file_info)
        # Подтвердить завершение транзакции к БД
        db.session.commit()


def update_time(file_hash):
    """ Функция обновления в БД времени последнего взаимодействия с файлом """

    # Получить из БД информацию о хеше, но в формате объекта записи
    file_info = get_info(file_hash, False)
    if file_info:
        # Если в БД содержится запрашиваемый хеш, обновить время взаимодействия
        file_info.modified_at = now_in_prefer_timezone()
        # Внести изменения в БД
        db.session.add(file_info)
        # Подтвердить завершение транзацкии к БД
        db.session.commit()


def delete_info(file_hash):
    """ Функция удаления из БД информации о файле """

    # Получить из БД информацию о хеше, но в формате объекта записи
    file_info = get_info(file_hash, False)
    if file_info:
        # Если в БД содержится запрашиваемый хеш, удалить его запись
        db.session.delete(file_info)
        # Подтвердить завершение транзацкии к БД
        db.session.commit()


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
    print(get_info(test_hash))

    delete_info(test_hash)
    print(get_info(test_hash))

    add_info(test_hash, 'xlsx')
    print(get_info(test_hash))

    update_time(test_hash)
    print(get_info(test_hash))

    # print(FileInfo.query.all())
