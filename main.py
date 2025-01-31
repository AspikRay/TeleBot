import telebot
import pandas as pd
import os
from telebot import types

TOKEN = '7681175404:AAHQIGRxJ1FTXxvlnycEHco4jCrIbJtNu_Q'
bot = telebot.TeleBot(TOKEN)

FILE_DIRECTORY = 'bot_data'
FILE_PATH = os.path.join(FILE_DIRECTORY, 'data.xlsx')
file_loaded = False


def find_column(df_columns, target_column, required=False, max_mismatches=2):
    target_column = target_column.lower()
    best_match = None
    min_mismatches = float('inf')

    for column in df_columns:
        column_lower = column.lower()
        mismatches = sum(c1 != c2 for c1, c2 in zip(target_column, column_lower))
        mismatches += abs(len(target_column) - len(column_lower))

        if mismatches <= max_mismatches and mismatches < min_mismatches:
            min_mismatches = mismatches
            best_match = column

    if required and not best_match:
        raise ValueError(f"Не нашел столбец '{target_column}'. Доступные: {df_columns}")
    return best_match


def calculate_homework_status_v1(file_path):
    try:
        try:
            df = pd.read_csv(file_path, sep=',')
        except:
            df = pd.read_excel(file_path)

        df.columns = df.columns.str.lower().str.strip()
        teacher_column = find_column(df.columns, 'фио преподавателя', True, max_mismatches=3)
        checked_columns = []
        for col_name in ['unnamed: 5', 'unnamed: 10', 'unnamed: 15']:
            column = find_column(df.columns, col_name, max_mismatches=3)
            if column:
                checked_columns.append(column)
        if not checked_columns:
            raise ValueError(f"Не нашел столбцы с проверкой 'unnamed: 5, unnamed: 10, unnamed: 15'. Доступные: {df.columns.tolist()}")
        for col in checked_columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
        df.dropna(subset=checked_columns, inplace=True)
        if df.empty:
            return "Проверка пустая. Нет числовых значений."
        df['Всего проверено'] = df[checked_columns].sum(axis=1)
        max_homework = df['Всего проверено'].max()
        if max_homework == 0:
            return "Нет проверенных заданий."
        df['Процент проверки'] = (df['Всего проверено'] / max_homework) * 100
        low_check = df[(df['Процент проверки'] < 75) & (df[teacher_column].notna())]
        result_v1 = []
        for _, row in low_check.iterrows():
            teacher_name = row[teacher_column]
            percentage = row['Процент проверки']
            result_v1.append(f"Преподаватель {teacher_name}, процент проверки {percentage:.2f}%. Нужно проверить.")
        return result_v1 or ["Нет преподавателей с низким процентом."]
    except ValueError as ve:
        return str(ve)
    except Exception as e:
        return f"Ошибка обработки: {e}"

def analyze_student_grades(file_path):
    try:
        try:
            df = pd.read_csv(file_path, sep=',')
        except:
            df = pd.read_excel(file_path)
        df.columns = df.columns.str.lower().str.strip()
        student_column = find_column(df.columns, 'фио', True, max_mismatches=3)
        homework_column = find_column(df.columns, 'homework', False, max_mismatches=3)
        classroom_column = find_column(df.columns, 'classroom', False, max_mismatches=3)
        exam_column = find_column(df.columns, 'average score', False, max_mismatches=3)
        for col in [homework_column, classroom_column, exam_column]:
            if col:
                df[col] = pd.to_numeric(df[col], errors='coerce')
        df.dropna(subset=[homework_column, classroom_column, exam_column], how='all', inplace=True)
        if df.empty:
            return "Нет данных для анализа или не найдены столбцы с баллами."
        if not any([homework_column, classroom_column, exam_column]):
            return "Не нашел столбцы 'homework', 'classroom' или 'average score'."
        grade_columns = [col for col in [homework_column, classroom_column, exam_column] if col]
        df['average_grade'] = df[grade_columns].mean(axis=1)
        low_grades = df[df['average_grade'] < 4]
        result_list = []
        for _, row in low_grades.iterrows():
            student_name = row[student_column]
            average_grade = row['average_grade']
            result_list.append(f"Студент {student_name}, средний балл {average_grade:.2f}. Балл низкий.")
        return result_list or ["Нет студентов с низким баллом."]
    except ValueError as ve:
        return str(ve)
    except Exception as e:
        return f"Ошибка обработки: {e}"


def create_directory_if_not_exists(directory):
    if not os.path.exists(directory):
        os.makedirs(directory)


def save_analysis_to_file(analysis_results, file_type, file_directory):
    create_directory_if_not_exists(file_directory)
    if file_type == "grades":
        file_name = os.path.join(file_directory, "student_grades_analysis.txt")
    elif file_type == "homework":
        file_name = os.path.join(file_directory, "homework_check_analysis.txt")
    else:
        file_name = os.path.join(file_directory, "analysis_results.txt")

    try:
        with open(file_name, 'w', encoding='utf-8') as f:
            for item in analysis_results:
                f.write(item + '\n')
    except Exception as e:
        return f"Ошибка записи: {e}"
    return file_name

@bot.message_handler(commands=['start'])
def start(message):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=2)
    item1 = types.KeyboardButton("Загрузить файл")
    item2 = types.KeyboardButton("Показать данные о домашних заданиях")
    item3 = types.KeyboardButton("Показать данные о отчете по студентам")
    markup.add(item1, item2, item3)
    bot.send_message(message.chat.id, "Привет! Что будем делать?", reply_markup=markup)

@bot.message_handler(func=lambda message: message.text == "Загрузить файл")
def load_file_button(message):
    bot.send_message(message.chat.id, "Отправь файл Excel (.xlsx, .xls) или CSV.")
    bot.register_next_step_handler(message, handle_document)

@bot.message_handler(func=lambda message: message.text == "Показать данные о домашних заданиях")
def show_data_button(message):
      show_data(message)

@bot.message_handler(func=lambda message: message.text == "Показать данные о отчете по студентам")
def analyze_grades_button(message):
      show_grades(message)


def handle_document(message):
    global file_loaded
    try:
        if not message.document:
            bot.send_message(message.chat.id, "Файл документом.")
            return
        create_directory_if_not_exists(FILE_DIRECTORY)
        file_id = message.document.file_id
        file_info = bot.get_file(file_id)
        file = bot.download_file(file_info.file_path)
        with open(FILE_PATH, 'wb') as f:
            f.write(file)
        file_loaded = True
        bot.send_message(message.chat.id, "Файл загружен!")
    except Exception as e:
        bot.send_message(message.chat.id, f"Ошибка загрузки: {e}")

@bot.message_handler(commands=['showdata'])
def show_data(message):
    global file_loaded
    if not file_loaded:
        bot.send_message(message.chat.id, "Сначала загрузи файл.")
        return
    try:
        result = calculate_homework_status_v1(FILE_PATH)
        if isinstance(result, list):
            file_path = save_analysis_to_file(result, "homework", FILE_DIRECTORY)
            with open(file_path, 'rb') as f:
               bot.send_document(message.chat.id, f)
        else:
            if result.strip():
                bot.send_message(message.chat.id, result)
    except Exception as e:
        bot.send_message(message.chat.id, f"Ошибка обработки: {e}")

@bot.message_handler(commands=['showgrades'])
def show_grades(message):
    global file_loaded
    if not file_loaded:
        bot.send_message(message.chat.id, "Сначала загрузи файл.")
        return
    try:
        result = analyze_student_grades(FILE_PATH)
        if isinstance(result, list):
            file_path = save_analysis_to_file(result, "grades", FILE_DIRECTORY)
            with open(file_path, 'rb') as f:
              bot.send_document(message.chat.id, f)
        else:
            if result.strip():
                bot.send_message(message.chat.id, result)
    except Exception as e:
        bot.send_message(message.chat.id, f"Ошибка обработки: {e}")


if __name__ == '__main__':
    bot.polling(none_stop=True)
