import telebot
import pandas as pd
from io import BytesIO

# Ваш токен для бота
TOKEN = '7681175404:AAHQIGRxJ1FTXxvlnycEHco4jCrIbJtNu_Q'

bot = telebot.TeleBot(TOKEN)

# Функция для расчета процента проверки
def calculate_homework_status(file):
    """
    Рассчитывает процент проверки только на основе столбца 'Проверено' и автоматически определяет максимальное количество проверенных заданий.
    """
    try:
        # Читаем Excel файл
        df = pd.read_excel(file, sheet_name=0)
        print("Первые строки файла:\n", df.head())
        print("Столбцы:", df.columns.tolist())

        # Проверяем наличие нужных столбцов
        required_columns = ['ФИО преподавателя', 'Unnamed: 5']
        if not all(col in df.columns for col in required_columns):
            return f"Ошибка: В файле отсутствуют нужные столбцы. Доступные столбцы: {df.columns.tolist()}"

        # Переименовываем столбцы для удобства
        df.rename(columns={'Unnamed: 5': 'Проверено'}, inplace=True)

        # Преобразуем данные в числовой формат, некорректные значения заменяем на 0
        df['Проверено'] = pd.to_numeric(df['Проверено'], errors='coerce').fillna(0)

        # Определяем максимальное количество проверенных заданий (максимальное значение в столбце 'Проверено')
        max_homework = df['Проверено'].max()

        # Рассчитываем процент проверки относительно максимального значения
        df['Процент проверки'] = (df['Проверено'] / max_homework) * 100

        # Фильтруем педагогов с процентом проверки ниже 75% и исключаем строки без ФИО
        low_check = df[(df['Процент проверки'] < 75) & (df['ФИО преподавателя'].notna())]

        # Формируем список для автоматических сообщений
        alerts = []
        for _, row in low_check.iterrows():
            alerts.append(f"Уважаемый(ая) {row['ФИО преподавателя']}, ваш процент проверки домашних заданий составляет {row['Процент проверки']:.2f}%. Пожалуйста, уделите внимание проверке.")

        return alerts

    except Exception as e:
        return f"Ошибка при обработке файла: {e}"

# Обработчик команды /start
@bot.message_handler(commands=['start'])
def start(message):
    bot.send_message(message.chat.id, "Привет! Отправь мне Excel файл с домашними заданиями, и я посчитаю процент проверенных заданий. Если у педагогов процент проверки ниже 75%, они получат уведомление.")

# Обработчик получения файлов
@bot.message_handler(content_types=['document'])
def handle_document(message):
    try:
        # Получаем файл
        file_id = message.document.file_id
        file_info = bot.get_file(file_id)
        file = bot.download_file(file_info.file_path)

        # Проверяем статус домашних заданий
        result = calculate_homework_status(BytesIO(file))

        if isinstance(result, str):  # Если возникла ошибка
            bot.send_message(message.chat.id, result)
        else:
            # Отправляем сообщения педагогам
            for alert in result:
                bot.send_message(message.chat.id, alert)

            bot.send_message(message.chat.id, "Сообщения отправлены педагогам с низким процентом проверки.")

    except Exception as e:
        bot.send_message(message.chat.id, f"Произошла ошибка: {e}")

# Запуск бота
if __name__ == '__main__':
    bot.polling(none_stop=True)
