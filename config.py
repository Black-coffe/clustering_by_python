# config.py
DB_CONFIG = {
    'host': 'localhost',
    'user': 'root',  # Заміни на свій логін MySQL
    'password': '',  # Заміни на свій пароль
    'database': 'clustering'  # Назва твоєї бази даних
}

# Інші параметри
BRIGHT_DATA_USERNAME = ''
BRIGHT_DATA_PASSWORD = ''

# API токен для получения результатов из Bright Data
BRIGHT_DATA_API_TOKEN = ''  # Добавьте здесь свой токен API от Bright Data

DATA_DIR = 'data'  # Шлях до папки з даними

# Константа домена поиска
SEARCH_DOMAIN = 'google.com.ua'  # Например, google.com или google.com.ua

# Константа страны или региона (двухбуквенный код, параметр gl)
COUNTRY = 'ua'  # Например, 'us' для США, 'ua' для Украины

# Константа языка поиска (двухбуквенный код, параметр hl)
LANGUAGE = 'uk'  # Например, 'en' для английского, 'uk' для украинского

# Константа количества страниц поиска (по умолчанию 1 страница)
PAGES = 1

# Константа количества результатов на одну страницу (по умолчанию 10)
RESULTS_PER_PAGE = 10

# Дополнительные параметры (по желанию)
# Например, тип поиска (tbm): 'isch' для изображений, 'shop' для шопинга и т.д.
# SEARCH_TYPE = 'isch'