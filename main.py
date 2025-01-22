import telebot
import pandas as pd
from io import BytesIO
import re

TOKEN = '7681175404:AAHQIGRxJ1FTXxvlnycEHco4jCrIbJtNu_Q'

bot = telebot.TeleBot(TOKEN)

# Функция для расчета процента проверки (Вариант 1)
def calculate_homework_status_v1(file_content, file_type):
    try:
        if file_type == 'csv':
            # Читаем CSV файл из BytesIO
            df = pd.read_csv(file_content)
        elif file_type == 'excel':
            df = pd.read_excel(file_content)

        # Проверяем наличие нужных столбцов
        required_columns = ['ФИО преподавателя', 'Unnamed: 5', 'Unnamed: 10', 'Unnamed: 15']
        if not all(col in df.columns for col in required_columns):
            return f"Ошибка: В файле отсутствуют нужные столбцы. Доступные столбцы: {df.columns.tolist()}"

        # Преобразуем данные в числовой формат, некорректные значения заменяем на NaN
        for col in ['Unnamed: 5', 'Unnamed: 10', 'Unnamed: 15']:
            df[col] = pd.to_numeric(df[col], errors='coerce')

        # Удаляем строки, где хотя бы в одном из столбцов 'Проверено' нет числового значения (NaN)
        df.dropna(subset=['Unnamed: 5', 'Unnamed: 10', 'Unnamed: 15'], inplace=True)

        # Если после удаления NaN таблица пуста, возвращаем сообщение об ошибке
        if df.empty:
            return "Ошибка: Столбцы 'Проверено' не содержат числовых значений."

        # Считаем общее количество проверенных работ по каждому преподавателю
        df['Всего проверено'] = df['Unnamed: 5'] + df['Unnamed: 10'] + df['Unnamed: 15']

        # Определяем максимальное количество проверенных заданий
        max_homework = df['Всего проверено'].max()

        # Если максимальное значение равно 0, значит, проверенных заданий нет
        if max_homework == 0:
            return "Ошибка: Максимальное значение в столбце 'Всего проверено' равно 0. Проверенные задания отсутствуют."

        # Рассчитываем процент проверки относительно максимального значения
        df['Процент проверки'] = (df['Всего проверено'] / max_homework) * 100

        # Фильтруем педагогов с процентом проверки ниже 75% и исключаем строки без ФИО
        low_check = df[(df['Процент проверки'] < 75) & (df['ФИО преподавателя'].notna())]

        # Формируем список для вывода
        result_v1 = []
        for _, row in low_check.iterrows():
            teacher_name = row['ФИО преподавателя']
            percentage = row['Процент проверки']
            result_v1.append(f"Уважаемый(ая) {teacher_name}, ваш процент проверки домашних заданий составляет {percentage:.2f}%. Пожалуйста, уделите внимание проверке.")

        return result_v1

    except Exception as e:
        return f"Ошибка при обработке файла (Вариант 1): {e}"

# Функция для проверки тем уроков (Вариант 2)
def check_lesson_topics_v2(file_content):
    try:
        df = pd.read_excel(file_content)  # Читаем как Excel

        result_v2 = []
        for index, row in df.iterrows():
            try:
                teacher_name = row['ФИО преподавателя']
                lesson_topic = row['Тема урока']

                # Исправленное регулярное выражение
                if not re.match(r"Урок №\d+\.\s*Тема:\s*$", lesson_topic):
                    result_v2.append(f"Уважаемый(ая) {teacher_name}, пожалуйста, проверьте правильность заполнения темы урока '{lesson_topic}'. Ожидаемый формат: 'Урок №. Тема:'")

            except KeyError as e:
                print(f"Ошибка в строке {index}: Не найден ключ {e}")
                continue
            except Exception as e:
                print(f"Ошибка в строке {index}: {e}")
                continue

        return result_v2

    except Exception as e:
        return f"Ошибка при обработке файла (Вариант 2): {e}"

# Обработчик команды /start
@bot.message_handler(commands=['start'])
def start(message):
    bot.send_message(message.chat.id, "Привет! Отправь мне Excel файл, и я обработаю его.")

# Обработчик получения файлов
@bot.message_handler(content_types=['document'])
def handle_document(message):
    try:
        # Получаем файл
        file_id = message.document.file_id
        file_info = bot.get_file(file_id)
        file = bot.download_file(file_info.file_path)
        file_name = message.document.file_name

        # Определяем формат файла
        if file_name.endswith('.xlsx') or file_name.endswith('.xls'):
            # Вариант 1: Расчет процента проверки
            result_v1 = calculate_homework_status_v1(BytesIO(file), 'excel')

            # Вариант 2: Проверка тем уроков
            result_v2 = check_lesson_topics_v2(BytesIO(file))

        elif file_name.endswith('.csv'):
            # Вариант 1: Расчет процента проверки
            result_v1 = calculate_homework_status_v1(BytesIO(file), 'csv')

            # Вариант 2: Проверка тем уроков (менее вероятно, но можно добавить)
            result_v2 = "Обработка тем уроков для CSV не реализована в этом примере"

        else:
            bot.send_message(message.chat.id, "Неподдерживаемый формат файла. Пожалуйста, отправьте файл в формате .xlsx, .xls или .csv.")
            return

        # Отправляем результаты Варианта 1
        if isinstance(result_v1, str):
            bot.send_message(message.chat.id, f"Результат проверки процента:\n{result_v1}")
        elif result_v1:  # Проверяем, что список не пустой
            bot.send_message(message.chat.id, "Результат проверки процента:")
            for alert_text in result_v1:
                bot.send_message(message.chat.id, alert_text)
        else:
            bot.send_message(message.chat.id, "Нет преподавателей с низким процентом проверки.")

        # Отправляем результаты Варианта 2
        if isinstance(result_v2, str):
            bot.send_message(message.chat.id, f"Результат проверки тем уроков:\n{result_v2}")
        elif result_v2:
            bot.send_message(message.chat.id, "Результат проверки тем уроков:")
            for message_text in result_v2:
                bot.send_message(message.chat.id, message_text)
        else:
            bot.send_message(message.chat.id, "Все темы уроков заполнены корректно (или не найдены).")

    except Exception as e:
        bot.send_message(message.chat.id, f"Произошла ошибка: {e}")

if __name__ == '__main__':
    bot.polling(none_stop=True)