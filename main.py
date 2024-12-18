import json
import pandas as pd
import telebot
from telebot import types
import os
from io import BytesIO

# TOKEN вашего бота
token = "7902339805:AAHABOA6Q-gYhzB6fl5CZEdHgiWWx2zwFBM"

bot = telebot.TeleBot(token)

# Хранилища состояния
user_states = {}
user_data = {}

# Состояния
STATE_NONE = "NONE"
STATE_EXCEL_TO_JSON_WAIT_FILE = "EXCEL_TO_JSON_WAIT_FILE"
STATE_EXCEL_TO_JSON_CHOOSE_SHEET = "EXCEL_TO_JSON_CHOOSE_SHEET"
STATE_EXCEL_TO_JSON_WAIT_COLUMN = "EXCEL_TO_JSON_WAIT_COLUMN"
STATE_EXCEL_TO_JSON_WAIT_START_ROW = "EXCEL_TO_JSON_WAIT_START_ROW"
STATE_JSON_TO_EXCEL_WAIT_FILE = "JSON_TO_EXCEL_WAIT_FILE"

def show_main_menu(user_id):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    item1 = types.KeyboardButton('Из Excel в Json')
    item2 = types.KeyboardButton('Из Json в Excel')
    markup.add(item1, item2)
    bot.send_message(user_id, "Выберите действие:", reply_markup=markup)

def show_return_menu_button(user_id, text="Вы можете вернуться в меню"):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    item_start = types.KeyboardButton('Вернуться в меню')
    markup.add(item_start)
    bot.send_message(user_id, text, reply_markup=markup)

@bot.message_handler(commands=['start'])
def start(message):
    user_id = message.chat.id
    user_states[user_id] = STATE_NONE
    user_data[user_id] = {}
    show_main_menu(user_id)


@bot.message_handler(content_types=['text'])
def bot_message(message):
    user_id = message.chat.id
    text = message.text.strip()
    state = user_states.get(user_id, STATE_NONE)

    # Обработка нажатия "Вернуться в меню"
    if text == "Вернуться в меню":
        # Возвращаем в начальное состояние
        user_states[user_id] = STATE_NONE
        user_data[user_id] = {}
        show_main_menu(user_id)
        return

    if state == STATE_NONE:
        if text == 'Из Excel в Json':
            user_states[user_id] = STATE_EXCEL_TO_JSON_WAIT_FILE
            # Скрываем предыдущее меню и показываем кнопку "Вернуться в меню"
            bot.send_message(user_id, "Отправь мне Excel файл (формат .xlsx).", reply_markup=types.ReplyKeyboardRemove())
            show_return_menu_button(user_id, "Ожидаю Excel файл...")
        elif text == 'Из Json в Excel':
            user_states[user_id] = STATE_JSON_TO_EXCEL_WAIT_FILE
            bot.send_message(user_id, "Отправь мне JSON файл.", reply_markup=types.ReplyKeyboardRemove())
            show_return_menu_button(user_id, "Ожидаю JSON файл...")
        else:
            bot.send_message(user_id, "Пожалуйста, выберите один из вариантов меню.")

    elif state == STATE_EXCEL_TO_JSON_WAIT_FILE:
        # Ожидаем файл Excel, пользователь отправил текст
        bot.send_message(user_id, "Отправьте файл Excel (.xlsx).")

    elif state == STATE_EXCEL_TO_JSON_CHOOSE_SHEET:
        # Ожидаем нажатия кнопки-инлайн с листом, пользователь отправил текст
        bot.send_message(user_id, "Выберите лист из предложенных кнопок.")

    elif state == STATE_EXCEL_TO_JSON_WAIT_COLUMN:
        # Ожидаем выбор столбца кнопками, пользователь отправил текст
        bot.send_message(user_id, "Выберите столбец из предложенных кнопок.")

    elif state == STATE_EXCEL_TO_JSON_WAIT_START_ROW:
        # Здесь пользователь должен ввести номер строки
        if not text.isdigit():
            bot.send_message(user_id, "Пожалуйста, введите корректный номер строки (число).")
            return

        start_row = int(text)
        user_data[user_id]['start_row'] = start_row

        # Преобразование в JSON
        excel_bytes = user_data[user_id]['excel_bytes']
        chosen_sheet = user_data[user_id]['chosen_sheet']
        chosen_column = user_data[user_id]['chosen_column']

        try:
            df = pd.read_excel(BytesIO(excel_bytes), sheet_name=chosen_sheet, header=None)
            col_index = ord(chosen_column) - ord('A')
            selected_data = df.iloc[start_row-1:, col_index].dropna()

            data_list = selected_data.tolist()
            json_data = json.dumps(data_list, ensure_ascii=False, indent=4)

            # Отправляем JSON как файл из памяти
            output_bytes = BytesIO(json_data.encode('utf-8'))
            output_bytes.seek(0)
            bot.send_document(user_id, output_bytes, visible_file_name='converted.json', caption="Вот ваш JSON файл")

            show_return_menu_button(user_id, "Готово!")
            user_states[user_id] = STATE_NONE
            user_data[user_id] = {}
            show_main_menu(user_id)

        except Exception as e:
            bot.send_message(user_id, f"Ошибка при обработке: {e}")
            show_return_menu_button(user_id, "Возникла ошибка, вы можете вернуться в меню.")

        user_states[user_id] = STATE_NONE
        user_data[user_id] = {}

    elif state == STATE_JSON_TO_EXCEL_WAIT_FILE:
        # Ожидаем файл JSON, пользователь отправил текст
        bot.send_message(user_id, "Отправьте файл JSON.")


@bot.message_handler(content_types=['document'])
def handle_docs(message):
    user_id = message.chat.id
    state = user_states.get(user_id, STATE_NONE)

    doc = message.document
    file_info = bot.get_file(doc.file_id)
    downloaded_file = bot.download_file(file_info.file_path)

    if state == STATE_EXCEL_TO_JSON_WAIT_FILE:
        # Проверяем, что это Excel
        if not doc.file_name.lower().endswith('.xlsx'):
            bot.send_message(user_id, "Пожалуйста, отправьте файл в формате .xlsx")
            return

        # Сохраняем файл в память
        try:
            # Пробуем открыть его через pandas
            xls = pd.ExcelFile(BytesIO(downloaded_file))
            sheets = xls.sheet_names
            user_data[user_id]['excel_bytes'] = downloaded_file
        except Exception as e:
            bot.send_message(user_id, f"Ошибка при чтении Excel файла: {e}")
            return

        # Предлагаем выбрать лист
        markup = types.InlineKeyboardMarkup()
        for sheet in sheets:
            markup.add(types.InlineKeyboardButton(text=sheet, callback_data=f"sheet:{sheet}"))
        bot.send_message(user_id, "Выберите лист:", reply_markup=markup)

        user_states[user_id] = STATE_EXCEL_TO_JSON_CHOOSE_SHEET

    elif state == STATE_JSON_TO_EXCEL_WAIT_FILE:
        # Проверяем, что это JSON
        if not doc.file_name.lower().endswith('.json'):
            bot.send_message(user_id, "Пожалуйста, отправьте файл в формате .json")
            return

        try:
            data = json.loads(downloaded_file.decode('utf-8'))
        except Exception as e:
            bot.send_message(user_id, f"Ошибка при чтении JSON: {e}")
            return

        # Конвертируем JSON в Excel
        try:
            if isinstance(data, dict):
                data = list(data.items())
            elif not isinstance(data, list):
                data = [data]

            df = pd.DataFrame(data)
            df_str = df.astype(str)
            one_col = df_str.apply(lambda x: ', '.join(x), axis=1)
            final_df = pd.DataFrame(one_col, columns=["Data"])

            output_io = BytesIO()
            final_df.to_excel(output_io, index=False)
            output_io.seek(0)

            bot.send_document(user_id, output_io, visible_file_name='converted.xlsx', caption="Вот ваш Excel файл")

            show_return_menu_button(user_id, "Готово!")
            user_states[user_id] = STATE_NONE
            user_data[user_id] = {}
            show_main_menu(user_id)

            user_states[user_id] = STATE_NONE
            user_data[user_id] = {}

        except Exception as e:
            bot.send_message(user_id, f"Ошибка при обработке JSON: {e}")
            show_return_menu_button(user_id, "Возникла ошибка, вы можете вернуться в меню.")
            user_states[user_id] = STATE_NONE
            user_data[user_id] = {}


@bot.callback_query_handler(func=lambda call: call.data.startswith("sheet:"))
def callback_choose_sheet(call):
    user_id = call.message.chat.id
    state = user_states.get(user_id, STATE_NONE)

    if state == STATE_EXCEL_TO_JSON_CHOOSE_SHEET:
        chosen_sheet = call.data.split("sheet:")[1]
        user_data[user_id]['chosen_sheet'] = chosen_sheet

        bot.edit_message_text(chat_id=user_id, message_id=call.message.message_id,
                              text=f"Вы выбрали лист: {chosen_sheet}")

        excel_bytes = user_data[user_id]['excel_bytes']
        df = pd.read_excel(BytesIO(excel_bytes), sheet_name=chosen_sheet, header=None)

        filled_columns = []
        for col_idx in range(df.shape[1]):
            col_series = df.iloc[:, col_idx]
            if col_series.notna().any():
                col_letter = chr(ord('A') + col_idx)
                filled_columns.append(col_letter)

        if not filled_columns:
            bot.send_message(user_id, "В этом листе нет заполненных столбцов.")
            show_return_menu_button(user_id, "Вы можете вернуться в меню")
            # Вернёмся к начальному состоянию
            user_states[user_id] = STATE_NONE
            user_data[user_id] = {}
            return

        markup = types.InlineKeyboardMarkup()
        for col_letter in filled_columns:
            markup.add(types.InlineKeyboardButton(text=col_letter, callback_data=f"column:{col_letter}"))

        bot.send_message(user_id, "Выберите заполненный столбец для преобразования:", reply_markup=markup)
        user_states[user_id] = STATE_EXCEL_TO_JSON_WAIT_COLUMN


@bot.callback_query_handler(func=lambda call: call.data.startswith("column:"))
def callback_choose_column(call):
    user_id = call.message.chat.id
    state = user_states.get(user_id, STATE_NONE)

    if state == STATE_EXCEL_TO_JSON_WAIT_COLUMN:
        chosen_column = call.data.split("column:")[1]
        user_data[user_id]['chosen_column'] = chosen_column

        bot.edit_message_text(chat_id=user_id, message_id=call.message.message_id,
                              text=f"Вы выбрали столбец: {chosen_column}")

        bot.send_message(user_id, "С какой строки начинаем сбор данных?")
        user_states[user_id] = STATE_EXCEL_TO_JSON_WAIT_START_ROW


bot.polling(non_stop=True)
