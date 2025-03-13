import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, simpledialog
import pandas as pd
import mysql.connector
from config import DB_CONFIG  # Импорт настроек базы данных
import re

# Глобальная переменная для хранения данных
df = None

# Функция для логирования
def log_message(message, color='green'):
    log_text.config(state=tk.NORMAL)
    log_text.insert(tk.END, message + '\n', color)
    log_text.config(state=tk.DISABLED)
    log_text.see(tk.END)

# Подключение к базе данных
def connect_to_db():
    try:
        return mysql.connector.connect(**DB_CONFIG)
    except mysql.connector.Error as e:
        log_message(f"Ошибка подключения к базе данных: {e}", 'red')
        raise

# Создание таблицы, если ее нет
def create_table_if_not_exists():
    conn = connect_to_db()
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS keywords (
            id INT AUTO_INCREMENT PRIMARY KEY,
            query VARCHAR(255) UNIQUE,
            frequency INT,
            prefix VARCHAR(50)
        )
    """)
    conn.commit()
    conn.close()
    log_message("Таблица 'keywords' проверена/создана.")

# Очистка таблицы
def clear_table():
    conn = connect_to_db()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM keywords")
    conn.commit()
    conn.close()
    log_message("Таблица 'keywords' очищена.")

# Вставка данных в таблицу
def insert_data(df, prefix):
    conn = connect_to_db()
    cursor = conn.cursor()
    for _, row in df.iterrows():
        try:
            cursor.execute("""
                INSERT INTO keywords (query, frequency, prefix)
                VALUES (%s, %s, %s)
            """, (row['query'], row['frequency'], prefix))
        except mysql.connector.IntegrityError:
            log_message(f"Дубликат запроса '{row['query']}' пропущен.", 'yellow')
    conn.commit()
    conn.close()
    log_message(f"Добавлено {len(df)} записей с префиксом '{prefix}'.")

# Проверка, является ли строка заголовком
def is_header(row):
    # Список слов, характерных для заголовков
    header_keywords = ['запрос', 'ключевое слово', 'частотность', 'frequency', 'query']
    # Если первая колонка содержит текстовые метки или вторая не число — это заголовок
    return any(keyword in str(row[0]).lower() for keyword in header_keywords) or not str(row[1]).isdigit()

# Загрузка файла с определением и удалением заголовка
def load_file():
    global df
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")])
    if file_path:
        try:
            if file_path.endswith('.csv'):
                # Читаем первую строку для проверки
                with open(file_path, 'r', encoding='utf-8') as f:
                    first_line = f.readline().strip().split(',')
                if is_header(first_line):
                    df = pd.read_csv(file_path, header=0, names=['query', 'frequency'])
                    log_message("Обнаружен и удалён заголовок в CSV файле.")
                else:
                    df = pd.read_csv(file_path, header=None, names=['query', 'frequency'])
                    log_message("Заголовок в CSV файле не обнаружен.")
            else:
                # Для Excel файлов
                df_temp = pd.read_excel(file_path, header=None)
                if is_header(df_temp.iloc[0]):
                    df = pd.read_excel(file_path, header=0, names=['query', 'frequency'])
                    log_message("Обнаружен и удалён заголовок в Excel файле.")
                else:
                    df = pd.read_excel(file_path, header=None, names=['query', 'frequency'])
                    log_message("Заголовок в Excel файле не обнаружен.")

            log_message(f"Файл '{file_path}' загружен. Строк: {len(df)}")
            check_button.config(state=tk.NORMAL)  # Активируем кнопку "Проверить файл"
            save_button.config(state=tk.DISABLED)  # Отключаем кнопку "Сохранить данные"
        except Exception as e:
            log_message(f"Ошибка при загрузке файла: {e}", 'red')
    else:
        log_message("Файл не выбран.", 'yellow')

# Проверка файла
def check_file():
    global df
    try:
        # Проверка структуры: две колонки
        if df.shape[1] != 2:
            raise ValueError("Файл должен содержать ровно две колонки: запросы и частотность.")

        # Проверка типов данных
        if not pd.api.types.is_string_dtype(df['query']):
            raise ValueError("Колонка 'query' должна содержать строки.")
        if not pd.api.types.is_integer_dtype(df['frequency']):
            raise ValueError("Колонка 'frequency' должна содержать целые числа.")

        # Проверка на пропуски
        if df.isnull().any().any():
            raise ValueError("Файл содержит пропущенные значения.")

        # Проверка на отрицательные или дробные значения
        if (df['frequency'] < 0).any():
            raise ValueError("Колонка 'frequency' содержит отрицательные значения.")

        # Удаление дубликатов
        initial_len = len(df)
        df.drop_duplicates(subset='query', keep='first', inplace=True)
        log_message(f"Удалено дубликатов: {initial_len - len(df)}. Осталось строк: {len(df)}")

        # Очистка запросов
        df['query'] = df['query'].str.lower().str.strip()  # Нижний регистр и удаление пробелов
        df['query'] = df['query'].apply(lambda x: re.sub(r'[^\w\s]', '', x))  # Удаление спецсимволов

        # Дополнительные проверки
        if (df['query'].str.len() < 1).any():
            raise ValueError("После очистки обнаружены пустые запросы.")
        if (df['query'].str.len() > 255).any():
            raise ValueError("Некоторые запросы превышают максимальную длину 255 символов.")

        log_message("Запросы очищены: нижний регистр, удалены пробелы и спецсимволы.")
        save_button.config(state=tk.NORMAL)
    except Exception as e:
        log_message(f"Ошибка при проверке файла: {e}", 'red')

# Сохранение данных
def save_data():
    try:
        create_table_if_not_exists()
        conn = connect_to_db()
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM keywords")
        count = cursor.fetchone()[0]
        conn.close()

        prefix = simpledialog.askstring("Префикс", "Введите префикс для нового набора данных (обязательно):")
        if not prefix:
            log_message("Префикс не указан. Данные не сохранены.", 'yellow')
            return

        if count > 0:
            choice = messagebox.askquestion("Таблица не пуста", "Таблица содержит данные. Удалить все данные или добавить новые?")
            if choice == 'yes':
                clear_table()
                insert_data(df, prefix)
            else:
                insert_data(df, prefix)
        else:
            insert_data(df, prefix)
    except Exception as e:
        log_message(f"Ошибка при сохранении данных: {e}", 'red')

# Создание GUI
root = tk.Tk()
root.title("Модуль загрузки семантики")
root.geometry("600x400")

# Кнопки
load_button = tk.Button(root, text="Загрузить файл", command=load_file)
load_button.pack(pady=10)

check_button = tk.Button(root, text="Проверить файл", command=check_file, state=tk.DISABLED)
check_button.pack(pady=10)

save_button = tk.Button(root, text="Сохранить данные", command=save_data, state=tk.DISABLED)
save_button.pack(pady=10)

# Текстовое поле для логов
log_text = scrolledtext.ScrolledText(root, height=10, bg='black', fg='green', state=tk.DISABLED)
log_text.pack(pady=10, fill=tk.BOTH, expand=True)
log_text.tag_config('green', foreground='green')
log_text.tag_config('red', foreground='red')
log_text.tag_config('yellow', foreground='yellow')

root.mainloop()