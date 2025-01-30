import telebot
import pandas as pd
from io import BytesIO
import os
from telebot import types
import re

TOKEN = '7681175404:AAHQIGRxJ1FTXxvlnycEHco4jCrIbJtNu_Q'  # Замените на свой токен
bot = telebot.TeleBot(TOKEN)

# Путь к папке для сохранения файлов (измените на нужный)
FILE_DIRECTORY = 'bot_data'
#  Путь к файлу (не изменяется, так как файл сохраняется в указанную выше папку)
FILE_PATH = os.path.join(FILE_DIRECTORY, 'data.xlsx')

# Флаг, чтобы отслеживать, загружен ли файл
file_loaded = False

# Функция для поиска столбца с нечетким соответствием
def find_column(df_columns, target_column, required=False, max_mismatches=2):
    target_column = target_column.lower()
    best_match = None
    min_mismatches = float('inf')

    for column in df_columns:
      column_lower = column.lower()
      mismatches = sum(c1 != c2 for c1, c2 in zip(target_column, column_lower))
      mismatches += abs(len(target_column) - len(column_lower))
      
      if mismatches <= max_mismatches and mismatches < min_mismatches :
         min_mismatches = mismatches
         best_match = column

    if required and not best_match:
        raise ValueError(f"Ошибка: не найден столбец похожий на '{target_column}'. Доступные столбцы: {df_columns}")

    return best_match


# Функция для расчета процента проверки
def calculate_homework_status_v1(file_path):
    try:
        # Читаем CSV или Excel, pandas сам определит тип файла
        try:
            df = pd.read_csv(file_path, sep=',')
        except:
            df = pd.read_excel(file_path)
        # Приводим названия столбцов к нижнему регистру и удаляем пробелы в начале и конце
        df.columns = df.columns.str.lower().str.strip()
        
        # Ищем нужные столбцы
        teacher_column = find_column(df.columns, 'фио преподавателя', True, max_mismatches=3)
        checked_columns = []
        for col_name in ['unnamed: 5', 'unnamed: 10', 'unnamed: 15']:
             column = find_column(df.columns, col_name, max_mismatches=3)
             if column:
                  checked_columns.append(column)
        if not checked_columns:
            raise ValueError(f"Ошибка: не найден ни один столбец с проверкой, похожий на 'unnamed: 5, unnamed: 10, unnamed: 15'. Доступные столбцы: {df.columns.tolist()}")
        # Преобразуем данные в числовой формат, некорректные значения заменяем на NaN
        for col in checked_columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
        # Удаляем строки, где хотя бы в одном из столбцов 'Проверено' нет числового значения (NaN)
        df.dropna(subset=checked_columns, inplace=True)
        # Если после удаления NaN таблица пуста, возвращаем сообщение об ошибке
        if df.empty:
            return "Ошибка: Столбцы с проверкой не содержат числовых значений."
        # Считаем общее количество проверенных работ по каждому преподавателю
        df['Всего проверено'] = df[checked_columns].sum(axis=1)
        # Определяем максимальное количество проверенных заданий
        max_homework = df['Всего проверено'].max()
        # Если максимальное значение равно 0, значит, проверенных заданий нет
        if max_homework == 0:
            return "Ошибка: Максимальное значение в столбце 'Всего проверено' равно 0. Проверенные задания отсутствуют."
        # Рассчитываем процент проверки относительно максимального значения
        df['Процент проверки'] = (df['Всего проверено'] / max_homework) * 100
        # Фильтруем педагогов с процентом проверки ниже 75% и исключаем строки без ФИО
        low_check = df[(df['Процент проверки'] < 75) & (df[teacher_column].notna())]
        # Формируем список для вывода
        result_v1 = []
        for _, row in low_check.iterrows():
            teacher_name = row[teacher_column]
            percentage = row['Процент проверки']
            result_v1.append(f"Уважаемый(ая) {teacher_name}, ваш процент проверки домашних заданий составляет {percentage:.2f}%. Пожалуйста, уделите внимание проверке.")
        return result_v1 or ["Нет преподавателей с низким процентом проверки."]
    except ValueError as ve:
        return str(ve)
    except Exception as e:
        return f"Ошибка при обработке файла: {e}"

# Функция для анализа успеваемости студентов
def analyze_student_grades(file_path):
    try:
        # Читаем CSV или Excel, pandas сам определит тип файла
        try:
            df = pd.read_csv(file_path, sep=',')
        except:
            df = pd.read_excel(file_path)
        # Приводим названия столбцов к нижнему регистру и удаляем пробелы
        df.columns = df.columns.str.lower().str.strip()

        # Ищем нужные столбцы
        student_column = find_column(df.columns, 'фио', True, max_mismatches=3)
        homework_column = find_column(df.columns, 'homework', False, max_mismatches=3)
        classroom_column = find_column(df.columns, 'classroom', False, max_mismatches=3)
        exam_column = find_column(df.columns, 'average score', False, max_mismatches=3)


        # Преобразуем баллы в числовой формат, некорректные значения заменяем на NaN
        for col in [homework_column, classroom_column, exam_column]:
          if col:
            df[col] = pd.to_numeric(df[col], errors='coerce')

        df.dropna(subset=[homework_column, classroom_column, exam_column] , how='all', inplace=True)


        if df.empty:
            return "Ошибка: Нет данных для анализа или не найдены необходимые столбцы с баллами."

        # Проверка на то, что хотя бы один столбец с оценками найден
        if not any([homework_column, classroom_column, exam_column]):
            return "Ошибка: не найден ни один из столбцов 'homework', 'classroom' или 'average score'. Проверьте файл."
       
       # Вычисляем средний балл, учитывая только найденные столбцы
        grade_columns = [col for col in [homework_column, classroom_column, exam_column] if col]
        df['average_grade'] = df[grade_columns].mean(axis=1)

        # Фильтруем студентов с средним баллом ниже 4 (из 12 баллов)
        low_grades = df[df['average_grade'] < 4]

        result_list = []
        for _, row in low_grades.iterrows():
            student_name = row[student_column]
            average_grade = row['average_grade']
            result_list.append(f"У студента {student_name} средний балл ({average_grade:.2f}) ниже 4. Рекомендуется уделить внимание успеваемости.")
        return result_list or ["Нет студентов со средним баллом ниже 4."]

    except ValueError as ve:
        return str(ve)
    except Exception as e:
        return f"Ошибка при обработке файла: {e}"

# Функция для создания директории, если ее нет
def create_directory_if_not_exists(directory):
    if not os.path.exists(directory):
        os.makedirs(directory)


# Обработчик команды /start
@bot.message_handler(commands=['start'])
def start(message):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=2)
    item1 = types.KeyboardButton("Загрузить файл")
    item2 = types.KeyboardButton("Показать данные")
    item3 = types.KeyboardButton("Анализ успеваемости")
    markup.add(item1, item2, item3)
    bot.send_message(message.chat.id, "Привет! Я бот для обработки Excel/CSV файлов. Выберите действие:", reply_markup=markup)
    

# Обработчик текстовых сообщений
@bot.message_handler(func=lambda message: message.text == "Загрузить файл")
def load_file_button(message):
    bot.send_message(message.chat.id, "Отправьте файл Excel (.xlsx, .xls) или CSV.")
    bot.register_next_step_handler(message, handle_document)

@bot.message_handler(func=lambda message: message.text == "Показать данные")
def show_data_button(message):
      show_data(message)

@bot.message_handler(func=lambda message: message.text == "Анализ успеваемости")
def analyze_grades_button(message):
      show_grades(message)

# Обработчик получения файлов
def handle_document(message):
    global file_loaded
    try:
        if not message.document:
            bot.send_message(message.chat.id, "Пожалуйста, отправьте файл как документ.")
            return

        # Создаем директорию, если ее нет
        create_directory_if_not_exists(FILE_DIRECTORY)

        file_id = message.document.file_id
        file_info = bot.get_file(file_id)
        file = bot.download_file(file_info.file_path)

        with open(FILE_PATH, 'wb') as f:
            f.write(file)
        file_loaded = True
        bot.send_message(message.chat.id, "Файл успешно загружен и сохранен!")
    except Exception as e:
        bot.send_message(message.chat.id, f"Произошла ошибка при загрузке файла: {e}")


# Обработчик команды /showdata
@bot.message_handler(commands=['showdata'])
def show_data(message):
    global file_loaded
    if not file_loaded:
        bot.send_message(message.chat.id, "Сначала загрузите файл.")
        return

    try:
        result = calculate_homework_status_v1(FILE_PATH)
        if isinstance(result, list):
            for msg in result:
                if msg.strip():
                    bot.send_message(message.chat.id, msg)
        else:
            if result.strip():
                bot.send_message(message.chat.id, result)
    except Exception as e:
        bot.send_message(message.chat.id, f"Произошла ошибка при обработке данных: {e}")

# Обработчик команды /showgrades
@bot.message_handler(commands=['showgrades'])
def show_grades(message):
    global file_loaded
    if not file_loaded:
        bot.send_message(message.chat.id, "Сначала загрузите файл.")
        return
    try:
        result = analyze_student_grades(FILE_PATH)
        if isinstance(result, list):
            for msg in result:
                if msg.strip():
                    bot.send_message(message.chat.id, msg)
        else:
            if result.strip():
                bot.send_message(message.chat.id, result)
    except Exception as e:
        bot.send_message(message.chat.id, f"Произошла ошибка при обработке данных: {e}")
        
if __name__ == '__main__':
    bot.polling(none_stop=True)