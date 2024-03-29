r"""Хранение константных данных сервера API

"""


# Максимальный размер файла (в килобайтах)
MAX_SIZE = 200 * 8 * 1024  # MAX * размерность

# Время хранения файла на сервере (N часов от времени последнего взаимодействия с файлом)
FILE_LIFETIME = 6 * 60 * 60  # 6 часов - оптимальный срок жизни файла

# Словарь расширений файлов
EXTENSIONS = {'d0cf11': 'xls',   # HEX-представление байтов расширения XLS
              '504b03': 'xlsx',  # HEX-представление байтов расширения XLSX
              '404': 'unknown'   # Произвольная строка для обозначения всех других расширений
              }

# Используемый часовой пояс
TIMEZONE = 'Asia/Yekaterinburg'
