import tkinter as tk
from tkinter import messagebox, ttk
import mysql.connector
from config import DB_CONFIG, BRIGHT_DATA_USERNAME, BRIGHT_DATA_PASSWORD

# Пробуем импортировать API токен, но не обязательно
try:
    from config import BRIGHT_DATA_API_TOKEN
except ImportError:
    BRIGHT_DATA_API_TOKEN = None
    print("[WARNING] BRIGHT_DATA_API_TOKEN не найден в config.py. Режим получения response_ID отключен.")
import json
import aiohttp
import asyncio
from datetime import datetime
from urllib.parse import urlencode

# --- Глобальные переменные для статистики ---
processed_count = 0  # Количество обработанных запросов
total_keywords = 0  # Всего запросов для обработки
start_time = None  # Время старта сбора
successful_api_calls = 0  # Успешные запросы к API
failed_api_calls = 0  # Неудавшиеся запросы к API


# --- Подключение к базе данных ---
def get_db_connection():
    conn = mysql.connector.connect(**DB_CONFIG)
    return conn


# --- Получение количества необработанных запросов ---
def get_unprocessed_keywords_count():
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT COUNT(*) FROM keywords k
        LEFT JOIN queries q ON k.query = q.query
        WHERE q.query IS NULL
    """)
    count = cursor.fetchone()[0]
    conn.close()
    return count


# --- Получение списка необработанных ключевых слов ---
def get_unprocessed_keywords(limit):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute(f"""
        SELECT k.id, k.query FROM keywords k
        LEFT JOIN queries q ON k.query = q.query
        WHERE q.query IS NULL
        LIMIT {limit}
    """)
    keywords = cursor.fetchall()
    conn.close()
    return keywords


# --- Преобразование ISO timestamp в MySQL datetime ---
def format_timestamp(timestamp_str):
    if not timestamp_str:
        return None
    try:
        dt = datetime.strptime(timestamp_str, '%Y-%m-%dT%H:%M:%S.%fZ')
        return dt.strftime('%Y-%m-%d %H:%M:%S')
    except ValueError:
        try:
            dt = datetime.strptime(timestamp_str, '%Y-%m-%dT%H:%M:%SZ')
            return dt.strftime('%Y-%m-%d %H:%M:%S')
        except ValueError:
            return None


# --- Создание таблиц ---
def create_tables():
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS keywords (
            id INT AUTO_INCREMENT PRIMARY KEY,
            query VARCHAR(255) UNIQUE,
            frequency INT,
            prefix VARCHAR(50)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS queries (
            id INT AUTO_INCREMENT PRIMARY KEY,
            query VARCHAR(255),
            timestamp DATETIME,
            search_engine VARCHAR(50),
            language VARCHAR(10),
            country VARCHAR(50),
            country_code VARCHAR(5),
            location VARCHAR(100),
            results_area VARCHAR(100),
            gl VARCHAR(5),
            mobile BOOLEAN,
            search_type VARCHAR(20)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS organic_results (
            id INT AUTO_INCREMENT PRIMARY KEY,
            query_id INT,
            position INT,
            global_rank INT,
            url TEXT,
            display_url TEXT,
            title VARCHAR(255),
            description TEXT,
            last_modified_date VARCHAR(50),
            has_image BOOLEAN,
            image_data LONGTEXT,
            FOREIGN KEY (query_id) REFERENCES queries(id)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS navigation_links (
            id INT AUTO_INCREMENT PRIMARY KEY,
            query_id INT,
            title VARCHAR(100),
            href TEXT,
            FOREIGN KEY (query_id) REFERENCES queries(id)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS related_queries (
            id INT AUTO_INCREMENT PRIMARY KEY,
            query_id INT,
            related_query VARCHAR(255),
            FOREIGN KEY (query_id) REFERENCES queries(id)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS people_also_ask (
            id INT AUTO_INCREMENT PRIMARY KEY,
            query_id INT,
            question VARCHAR(255),
            answer_source TEXT,
            answer_link TEXT,
            FOREIGN KEY (query_id) REFERENCES queries(id)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS knowledge (
            id INT AUTO_INCREMENT PRIMARY KEY,
            query_id INT,
            description TEXT,
            facts TEXT,
            FOREIGN KEY (query_id) REFERENCES queries(id)
        )
    """)

    conn.commit()
    conn.close()


# --- Обновление структуры таблицы (если уже существует) ---
def update_table_schema():
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("ALTER TABLE people_also_ask MODIFY COLUMN answer_source TEXT")
        cursor.execute("ALTER TABLE people_also_ask MODIFY COLUMN answer_link TEXT")
        cursor.execute("ALTER TABLE organic_results MODIFY COLUMN url TEXT")
        cursor.execute("ALTER TABLE organic_results MODIFY COLUMN display_url TEXT")
        conn.commit()
    except mysql.connector.Error as e:
        print(f"Ошибка при изменении структуры таблиц: {e}")
    conn.close()


# --- Семафор для ограничения количества одновременных запросов ---
MAX_CONCURRENT_REQUESTS = 5  # Максимальное количество параллельных запросов


# --- Асинхронная отправка запроса к Bright Data SERP API ---
async def get_serp_json(session, semaphore, keyword, keyword_id, country_code="ua", language="uk", num_results=10,
                        retries=2):
    global successful_api_calls, failed_api_calls

    # Формируем параметры запроса
    params = {
        "q": keyword,
        "gl": country_code,
        "hl": language,
        "num": num_results,
        "brd_json": 1,  # Bright Data JSON формат
        "location": "Ukraine"
    }

    # Добавляем UULE только если страна - Украина
    if country_code.lower() == "ua":
        params["uule"] = "w+CAIQICIHVWtyYWluZQ"

    search_url = f"https://www.google.com/search?{urlencode(params)}"
    proxy_url = f"http://{BRIGHT_DATA_USERNAME}:{BRIGHT_DATA_PASSWORD}@brd.superproxy.io:33335"

    for attempt in range(retries + 1):
        try:
            # Используем семафор для ограничения количества запросов
            async with semaphore:
                print(f"[INFO] Отправка запроса для '{keyword}', country={country_code}, lang={language}")
                async with session.get(search_url, proxy=proxy_url, ssl=False, timeout=60) as response:
                    if response.status == 200:
                        response_data = await response.text()
                        try:
                            data = json.loads(response_data)
                            if "response_ID" in data:  # Проверяем, вернул ли API response_ID
                                successful_api_calls += 1
                                return {"response_ID": data["response_ID"], "keyword_id": keyword_id,
                                        "keyword": keyword}
                            else:  # Если результат пришел сразу
                                successful_api_calls += 1
                                return {"data": data, "keyword_id": keyword_id, "keyword": keyword}
                        except json.JSONDecodeError as e:
                            if attempt < retries:
                                print(
                                    f"[RETRY] Ошибка декодирования JSON для '{keyword}', попытка {attempt + 1}/{retries + 1}: {e}")
                                await asyncio.sleep(2)  # Пауза перед следующей попыткой
                                continue
                            failed_api_calls += 1
                            print(f"[ERROR] Ошибка декодирования JSON для '{keyword}': {e}")
                            print(f"[DEBUG] Первые 200 символов ответа: {response_data[:200]}")
                            return None
                    else:
                        if attempt < retries:
                            print(
                                f"[RETRY] Ошибка API: статус {response.status} для '{keyword}', попытка {attempt + 1}/{retries + 1}")
                            await asyncio.sleep(2)  # Пауза перед следующей попыткой
                            continue
                        failed_api_calls += 1
                        print(f"[WARNING] Ошибка API: статус {response.status} для '{keyword}'")
                        return None
        except aiohttp.ClientError as e:
            if attempt < retries:
                print(
                    f"[RETRY] Ошибка HTTP при запросе к API для '{keyword}', попытка {attempt + 1}/{retries + 1}: {e}")
                await asyncio.sleep(2)  # Пауза перед следующей попыткой
                continue
            failed_api_calls += 1
            print(f"[WARNING] Ошибка HTTP при запросе к API для '{keyword}': {e}")
            return None
        except asyncio.TimeoutError:
            if attempt < retries:
                print(f"[RETRY] Таймаут при запросе к API для '{keyword}', попытка {attempt + 1}/{retries + 1}")
                await asyncio.sleep(2)  # Пауза перед следующей попыткой
                continue
            failed_api_calls += 1
            print(f"[WARNING] Таймаут при запросе к API для '{keyword}'")
            return None
        except Exception as e:
            if attempt < retries:
                print(
                    f"[RETRY] Неизвестная ошибка при запросе к API для '{keyword}', попытка {attempt + 1}/{retries + 1}: {type(e).__name__}: {e}")
                await asyncio.sleep(2)  # Пауза перед следующей попыткой
                continue
            failed_api_calls += 1
            print(f"[WARNING] Неизвестная ошибка при запросе к API для '{keyword}': {type(e).__name__}: {e}")
            return None


# --- Получение результатов по response_ID ---
async def get_result(session, response_id, keyword, max_attempts=5, delay=5):
    # Проверяем, доступен ли API токен
    if not BRIGHT_DATA_API_TOKEN:
        print(
            f"[ERROR] Невозможно получить результат для {response_id} ('{keyword}') - API токен не настроен в config.py")
        return None

    for attempt in range(max_attempts):
        results_url = f"https://api.brightdata.com/serp/results?response_id={response_id}"
        headers = {"Authorization": f"Bearer {BRIGHT_DATA_API_TOKEN}"}
        try:
            async with session.get(results_url, headers=headers, ssl=False) as response:
                if response.status == 200:
                    result = await response.json()
                    return result
                elif response.status == 202:  # Результат еще не готов
                    print(
                        f"[INFO] Результат для {response_id} ('{keyword}') еще не готов, попытка {attempt + 1}/{max_attempts}")
                    await asyncio.sleep(delay)
                else:
                    print(f"[ERROR] Ошибка для {response_id} ('{keyword}'): {response.status}")
                    return None
        except Exception as e:
            print(f"[ERROR] Ошибка для {response_id} ('{keyword}'): {e}")
            return None
    print(f"[ERROR] Не удалось получить результат для {response_id} ('{keyword}') после {max_attempts} попыток")
    return None


# --- Парсинг ответа API ---
def parse_serp_response(response):
    if not response:
        return {}
    parsed_data = {
        'general': response.get('general', {}),
        'organic': response.get('organic', []),
        'navigation': response.get('navigation', []),
        'related': response.get('related', []),
        'people_also_ask': response.get('people_also_ask', []),
        'knowledge': response.get('knowledge', {})
    }
    return parsed_data


# --- Сохранение данных в базу ---
def store_serp_data(keyword_id, keyword, data):
    conn = get_db_connection()
    cursor = conn.cursor()

    try:
        # Убедимся, что в general есть query, или используем keyword
        general = data.get('general', {})
        if not general.get('query'):
            general['query'] = keyword

        cursor.execute("""
            INSERT INTO queries (
                query, timestamp, search_engine, language, country, 
                country_code, location, results_area, gl, mobile, search_type
            )
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        """, (
            general.get('query'),
            format_timestamp(general.get('timestamp')),
            general.get('search_engine'),
            general.get('language'),
            general.get('country'),
            general.get('country_code'),
            general.get('location'),
            general.get('results_area'),
            general.get('gl'),
            general.get('mobile', False),
            general.get('search_type')
        ))
        query_id = cursor.lastrowid

        organic_results = data.get('organic', [])
        print(
            f"[DEBUG] Для запроса ID={query_id} (из keyword_id={keyword_id}) получено органических результатов: {len(organic_results)}")

        for item in organic_results:
            position_val = item.get('rank')
            url_val = item.get('link')
            display_val = item.get('display_link')
            global_rank_val = item.get('global_rank')
            has_image = bool(item.get('image'))
            cursor.execute("""
                INSERT INTO organic_results (
                    query_id, position, global_rank, url, display_url, title, description, 
                    last_modified_date, has_image, image_data
                )
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            """, (
                query_id,
                position_val,
                global_rank_val,
                url_val,
                display_val,
                item.get('title'),
                item.get('description'),
                item.get('last_modified_date'),
                has_image,
                json.dumps(item.get('image')) if has_image else None
            ))

        for item in data.get('navigation', []):
            cursor.execute("""
                INSERT INTO navigation_links (query_id, title, href)
                VALUES (%s, %s, %s)
            """, (
                query_id,
                item.get('title'),
                item.get('href')
            ))

        for item in data.get('related', []):
            cursor.execute("""
                INSERT INTO related_queries (query_id, related_query)
                VALUES (%s, %s)
            """, (query_id, item.get('text')))

        for item in data.get('people_also_ask', []):
            cursor.execute("""
                INSERT INTO people_also_ask (query_id, question, answer_source, answer_link)
                VALUES (%s, %s, %s, %s)
            """, (
                query_id,
                item.get('question'),
                item.get('answer_source'),
                item.get('answer_link')
            ))

        knowledge = data.get('knowledge', {})
        if knowledge:
            cursor.execute("""
                INSERT INTO knowledge (query_id, description, facts)
                VALUES (%s, %s, %s)
            """, (
                query_id,
                knowledge.get('description'),
                json.dumps(knowledge.get('facts')) if knowledge.get('facts') else None
            ))

        conn.commit()
        return query_id
    except Exception as e:
        print(f"[ERROR] Ошибка при сохранении данных для '{keyword}': {e}")
        conn.rollback()
        return None
    finally:
        conn.close()


# --- Асинхронная обработка одного ключевого слова ---
async def process_keyword(session, semaphore, keyword_tuple, country_code, language_code, num_results):
    keyword_id, keyword_text = keyword_tuple

    response = await get_serp_json(
        session,
        semaphore,
        keyword_text,
        keyword_id,
        country_code=country_code,
        language=language_code,
        num_results=num_results
    )

    return response


# --- Разделение списка на части ---
def chunk_list(lst, chunk_size):
    """Разделяет список на части указанного размера"""
    return [lst[i:i + chunk_size] for i in range(0, len(lst), chunk_size)]


# --- Главная функция сбора данных ---
async def async_start_collection():
    global processed_count, total_keywords, start_time, successful_api_calls, failed_api_calls

    try:
        domain = combo_domain.get()
        country_code = combo_country.get().split()[0]
        language_code = combo_language.get().split()[0]
        results_per_page = int(entry_results.get())
        batch_size = int(entry_batch.get())  # Размер пакета запросов

        create_tables()
        update_table_schema()

        limit = int(entry_limit.get())
        keywords = get_unprocessed_keywords(limit)
        total_keywords = len(keywords)
        processed_count = 0
        successful_api_calls = 0
        failed_api_calls = 0
        start_time = datetime.now()

        # Создаем семафор для ограничения количества одновременных запросов
        semaphore = asyncio.Semaphore(MAX_CONCURRENT_REQUESTS)

        # Проверяем режим работы - с API токеном или без
        api_mode = BRIGHT_DATA_API_TOKEN is not None
        if not api_mode:
            print("[INFO] Работа в режиме прямого прокси без API токена.")
        else:
            print("[INFO] Работа в режиме API с получением response_ID.")

        # Разделяем все ключевые слова на пакеты для обработки
        keyword_batches = chunk_list(keywords, batch_size)
        total_batches = len(keyword_batches)

        # Обработка ключевых слов по пакетам
        async with aiohttp.ClientSession() as session:
            for batch_index, keyword_batch in enumerate(keyword_batches):
                print(f"[INFO] Обработка пакета {batch_index + 1}/{total_batches} ({len(keyword_batch)} ключевых слов)")

                # Список для хранения ответов с response_ID и результатов для прямой обработки
                response_id_list = []
                direct_results = []

                # Отправка запросов для текущего пакета
                tasks = []
                for keyword in keyword_batch:
                    task = process_keyword(session, semaphore, keyword, country_code, language_code, results_per_page)
                    tasks.append(task)

                results = await asyncio.gather(*tasks)

                for result in results:
                    if result:
                        if "response_ID" in result and api_mode:
                            response_id_list.append(result)
                            print(f"[INFO] Получен response_ID: {result['response_ID']} для '{result['keyword']}'")
                        elif "data" in result:
                            direct_results.append(result)
                            print(f"[INFO] Получен прямой результат для '{result['keyword']}'")
                        elif "response_ID" in result and not api_mode:
                            # Если получили response_ID, но у нас нет API токена
                            print(
                                f"[WARNING] Получен response_ID для '{result['keyword']}', но API токен не настроен. Запрос пропущен.")
                            failed_api_calls += 1

                # Обработка прямых результатов
                for result in direct_results:
                    data = parse_serp_response(result["data"])
                    query_id = store_serp_data(result["keyword_id"], result["keyword"], data)
                    processed_count += 1
                    if query_id:
                        print(f"[INFO] Обработан прямой результат для '{result['keyword']}', query_id={query_id}")
                    else:
                        print(f"[WARNING] Не удалось сохранить прямой результат для '{result['keyword']}'")

                    # Обновление прогресса
                    progress = (processed_count / total_keywords) * 100
                    update_progress(progress)

                # Вторая фаза: получение и обработка результатов по response_ID
                if response_id_list:
                    print(f"[INFO] Получение результатов по {len(response_id_list)} response_ID")

                    for response_info in response_id_list:
                        response_id = response_info["response_ID"]
                        keyword_id = response_info["keyword_id"]
                        keyword = response_info["keyword"]

                        print(f"[INFO] Получение результата для response_ID: {response_id} (keyword: '{keyword}')")
                        result = await get_result(session, response_id, keyword)

                        if result:
                            data = parse_serp_response(result)
                            query_id = store_serp_data(keyword_id, keyword, data)
                            if query_id:
                                print(f"[INFO] Обработан результат по response_ID для '{keyword}', query_id={query_id}")
                            else:
                                print(f"[WARNING] Не удалось сохранить результат по response_ID для '{keyword}'")
                        else:
                            print(f"[WARNING] Не удалось получить результат по response_ID для '{keyword}'")
                            failed_api_calls += 1

                        processed_count += 1

                        # Обновление прогресса
                        progress = (processed_count / total_keywords) * 100
                        update_progress(progress)

                # Небольшая пауза между пакетами для "охлаждения" соединения
                if batch_index < total_batches - 1:
                    pause_seconds = 5
                    print(f"[INFO] Пауза между пакетами: {pause_seconds} секунд...")
                    await asyncio.sleep(pause_seconds)

        # Завершение сбора
        end_time = datetime.now()
        duration = (end_time - start_time).total_seconds()

        summary = (
            f"Сбор данных завершен!\n"
            f"Обработано: {processed_count}/{total_keywords} запросов\n"
            f"Успешных запросов к API: {successful_api_calls}\n"
            f"Неудачных запросов: {failed_api_calls}\n"
            f"Время выполнения: {duration:.2f} сек"
        )

        messagebox.showinfo("Успех", summary)

    except Exception as e:
        error_msg = f"Произошла ошибка: {str(e)}\n\nДеталей: {type(e).__name__}"
        messagebox.showerror("Ошибка", error_msg)
        print(f"[ERROR] {error_msg}")


# --- Запуск асинхронной функции из синхронного контекста ---
def start_collection():
    asyncio.run(async_start_collection())


# --- Обновление прогресс-бара ---
def update_progress(value):
    progress_bar["value"] = value
    progress_percent.set(f"{value:.1f}%")
    root.update_idletasks()


# --- Очистка журнала ---
def clear_log():
    log_text.delete(1.0, tk.END)
    print("[INFO] Журнал очищен")


# --- Обновление счетчика ключевых слов ---
def update_keywords_count():
    try:
        count = get_unprocessed_keywords_count()
        keywords_count_var.set(f"Необработанных запросов: {count}")
        root.after(5000, update_keywords_count)  # Обновляем каждые 5 секунд
    except Exception as e:
        keywords_count_var.set(f"Ошибка при получении количества запросов: {e}")
        root.after(10000, update_keywords_count)  # При ошибке повторяем через 10 секунд

# --- Создание GUI ---
root = tk.Tk()
root.title("Сбор данных SERP через Bright Data")
root.geometry("800x700")
root.configure(bg="#f0f0f0")

style = ttk.Style()
style.configure("TLabel", background="#f0f0f0", font=("Arial", 10))
style.configure("TButton", font=("Arial", 10, "bold"))

# Верхняя рамка с статистикой
stats_frame = tk.Frame(root, bg="#f0f0f0")
stats_frame.pack(pady=5, padx=10, fill=tk.X)

keywords_count_var = tk.StringVar()
keywords_count_var.set("Загрузка...")
label_count = tk.Label(stats_frame, textvariable=keywords_count_var, font=("Arial", 12, "bold"), bg="#f0f0f0")
label_count.pack(pady=5)

# Рамка с настройками
settings_frame = tk.LabelFrame(root, text="Настройки запросов", bg="#f0f0f0", font=("Arial", 10, "bold"))
settings_frame.pack(pady=5, padx=10, fill=tk.X)

# Домен
label_domain = tk.Label(settings_frame, text="Домен Google:", bg="#f0f0f0")
label_domain.grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
combo_domain = ttk.Combobox(settings_frame, values=[
    "google.com", "google.com.ua", "google.pl", "google.ru", "google.by",
    "google.kz", "google.de", "google.fr", "google.it", "google.es"
], state="readonly", width=30)
combo_domain.set("google.com.ua")
combo_domain.grid(row=0, column=1, padx=5, pady=2, sticky=tk.W)

# Страна
label_country = tk.Label(settings_frame, text="Страна:", bg="#f0f0f0")
label_country.grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
combo_country = ttk.Combobox(settings_frame, values=[
    "ua (Украина)", "us (США)", "pl (Польша)", "ru (Россия)", "by (Беларусь)",
    "kz (Казахстан)", "de (Германия)", "fr (Франция)", "it (Италия)", "es (Испания)"
], state="readonly", width=30)
combo_country.set("ua (Украина)")
combo_country.grid(row=1, column=1, padx=5, pady=2, sticky=tk.W)

# Язык
label_language = tk.Label(settings_frame, text="Язык поиска:", bg="#f0f0f0")
label_language.grid(row=2, column=0, sticky=tk.W, padx=5, pady=2)
combo_language = ttk.Combobox(settings_frame, values=[
    "uk (украинский)", "en (английский)", "pl (польский)", "ru (русский)",
    "be (белорусский)", "kk (казахский)", "de (немецкий)", "fr (французский)",
    "it (итальянский)", "es (испанский)"
], state="readonly", width=30)
combo_language.set("uk (украинский)")
combo_language.grid(row=2, column=1, padx=5, pady=2, sticky=tk.W)

# Результаты на страницу
label_results = tk.Label(settings_frame, text="Результатов на страницу:", bg="#f0f0f0")
label_results.grid(row=3, column=0, sticky=tk.W, padx=5, pady=2)
entry_results = tk.Entry(settings_frame, width=10)
entry_results.insert(0, "10")
entry_results.grid(row=3, column=1, padx=5, pady=2, sticky=tk.W)

# Тип поиска
label_search_type = tk.Label(settings_frame, text="Тип поиска:", bg="#f0f0f0")
label_search_type.grid(row=4, column=0, sticky=tk.W, padx=5, pady=2)
combo_search_type = ttk.Combobox(settings_frame, values=[
    "" " (текст)", "isch (изображения)", "map (карта)", "nws (новости)", "vid (видео)"
], state="readonly", width=30)
combo_search_type.set("" " (текст)")
combo_search_type.grid(row=4, column=1, padx=5, pady=2, sticky=tk.W)

# Рамка с дополнительными настройками
advanced_frame = tk.LabelFrame(root, text="Оптимизация сбора", bg="#f0f0f0", font=("Arial", 10, "bold"))
advanced_frame.pack(pady=5, padx=10, fill=tk.X)

# Лимит запросов
label_limit = tk.Label(advanced_frame, text="Количество запросов для обработки:", bg="#f0f0f0")
label_limit.grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
entry_limit = tk.Entry(advanced_frame, width=10)
entry_limit.insert(0, "10")
entry_limit.grid(row=0, column=1, padx=5, pady=2, sticky=tk.W)

# Размер пакета запросов
label_batch = tk.Label(advanced_frame, text="Размер пакета запросов:", bg="#f0f0f0")
label_batch.grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
entry_batch = tk.Entry(advanced_frame, width=10)
entry_batch.insert(0, "5")
entry_batch.grid(row=1, column=1, padx=5, pady=2, sticky=tk.W)
tk.Label(advanced_frame, text="(Рекомендуется 5-10 для стабильной работы)", bg="#f0f0f0", fg="gray").grid(row=1,
                                                                                                          column=2,
                                                                                                          padx=5,
                                                                                                          pady=2,
                                                                                                          sticky=tk.W)

# Прогресс-бар
progress_frame = tk.LabelFrame(root, text="Прогресс сбора", bg="#f0f0f0", font=("Arial", 10, "bold"))
progress_frame.pack(pady=10, padx=10, fill=tk.X)

progress_percent = tk.StringVar()
progress_percent.set("0.0%")

progress_bar = ttk.Progressbar(progress_frame, orient="horizontal", length=300, mode="determinate")
progress_bar.pack(padx=10, pady=10, fill=tk.X)

progress_info_frame = tk.Frame(progress_frame, bg="#f0f0f0")
progress_info_frame.pack(fill=tk.X, padx=10, pady=5)

progress_label = tk.Label(progress_info_frame, text="Выполнено:", bg="#f0f0f0")
progress_label.pack(side=tk.LEFT, padx=5)

progress_value_label = tk.Label(progress_info_frame, textvariable=progress_percent, bg="#f0f0f0", width=8, anchor="w")
progress_value_label.pack(side=tk.LEFT, padx=5)

# Журнал событий
log_frame = tk.LabelFrame(root, text="Журнал событий", bg="#f0f0f0", font=("Arial", 10, "bold"))
log_frame.pack(pady=5, padx=10, fill=tk.BOTH, expand=True)

log_text = tk.Text(log_frame, height=8, width=70, font=("Consolas", 9))
log_text.pack(padx=5, pady=5, fill=tk.BOTH, expand=True)

log_scrollbar = ttk.Scrollbar(log_text, command=log_text.yview)
log_text.configure(yscrollcommand=log_scrollbar.set)
log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

# Переопределяем print для записи в журнал
original_print = print


def log_print(*args, **kwargs):
    message = " ".join(map(str, args))
    original_print(*args, **kwargs)
    if log_text:
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        log_text.see(tk.END)  # Прокрутка к последней строке
        root.update_idletasks()  # Обновление UI


print = log_print

# Кнопки управления
control_frame = tk.Frame(root, bg="#f0f0f0")
control_frame.pack(pady=10, padx=10, fill=tk.X)

button_start = tk.Button(control_frame, text="Начать сбор", command=start_collection, bg="#4CAF50", fg="white",
                         font=("Arial", 11, "bold"), padx=20, pady=5)
button_start.pack(side=tk.LEFT, padx=5)

button_stop = tk.Button(control_frame, text="Остановить", command=lambda: None, bg="#f44336", fg="white",
                        font=("Arial", 11, "bold"), padx=20, pady=5, state=tk.DISABLED)
button_stop.pack(side=tk.LEFT, padx=5)

button_exit = tk.Button(control_frame, text="Выход", command=root.quit, bg="#607D8B", fg="white",
                        font=("Arial", 11, "bold"), padx=20, pady=5)
button_exit.pack(side=tk.RIGHT, padx=5)

# --- Добавление кнопки очистки журнала ---
button_clear_log = tk.Button(log_frame, text="Очистить журнал", command=clear_log, bg="#9E9E9E", fg="white",
                           font=("Arial", 9), padx=10)
button_clear_log.pack(side=tk.BOTTOM, anchor=tk.SE, padx=5, pady=5)


# --- Инициализация интерфейса ---
def init_ui():
    # Выводим информацию о версии
    print("[INFO] SERP Collector v1.0 - Bright Data SERP API")
    print("[INFO] Инициализация интерфейса...")

    # Проверяем конфигурацию
    if BRIGHT_DATA_API_TOKEN:
        print("[INFO] API токен Bright Data найден, доступен режим с response_ID")
    else:
        print("[WARNING] API токен Bright Data не найден, доступен только режим прямого прокси")

    # Запускаем регулярное обновление количества запросов
    update_keywords_count()


# Инициализация интерфейса
init_ui()

# Запуск главного цикла
root.mainloop()