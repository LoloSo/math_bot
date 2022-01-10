import random

import telegram

# from settings import TOKEN
from telegram import Update, ReplyKeyboardMarkup, ReplyKeyboardRemove
from telegram.ext import Updater, MessageHandler, Filters, CommandHandler, CallbackContext, CallbackQueryHandler
from openpyxl import load_workbook
from telegram import InlineKeyboardButton, InlineKeyboardMarkup
from sys import exc_info
from traceback import extract_tb
import theory_text

wb = load_workbook('database.xlsx')
vendor = wb['users']


def main():
    updater = Updater(token='1684008354:AAGtETeI2tgXd-AM_6AQvJmh43aTSYjkXxE', use_context=True)
    dispatcher = updater.dispatcher

    handler = MessageHandler(Filters.command, do_command)

    # start_handler = CommandHandler('start', do_start)
    # help_handler = CommandHandler('help', do_help)
    keyboard_handler = MessageHandler(Filters.text, keyboard_value)
    # photo_handler = MessageHandler(Filters.photo, do_photo)

    dispatcher.add_handler(handler)
    dispatcher.add_handler(keyboard_handler)
    # dispatcher.add_handler(photo_handler)

    updater.start_polling()
    updater.idle()


def do_photo(update: Update, context):
    print('a')
    # idp = update.message.photo.file_id
    # print(idp)
    update.message.reply_photo(open('4bedf30c2410ab76334d86f35aaf689c (1).png', 'rb'))


def do_start(update, context):
    keyboard = [
        ["Теория"],
        ["Практика"],
        ['Прочее']
    ]

    update.message.reply_text(text='Добро пожаловать, выберите, что бы вы хотели сделать.',
                              reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True, resize_keyboard=True))


def menu(update, context):
    keyboard = [
        ["Теория"],
        ["Практика"]
    ]

    update.message.reply_text(text='Выберите, что бы вы хотели сделать.',
                              reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True, resize_keyboard=True))

    """Функция для работы с кейбордом и темами"""


def keyboard_value(update: Update, context):
    text = update.message.text

    try:

        if 'command' in context.user_data:
            if update.message.text != 'Сообщить о проблеме' and context.user_data['command'] == '/addven':
                do_words(update, context)

        if text == 'Прочее':
            keyboard = [
                ['Сообщить о ошибке'], ['Предложения'],
                ['Назад']
            ]

            update.message.reply_text(text='Добро пожаловать, выберите, что бы вы хотели сделать.',
                                      reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True,
                                                                       resize_keyboard=True))

        if text == 'Назад':
            context.user_data['sort'] = None
            context.user_data['practice'] = None
            menu(update=update, context=context)
            return

        if text == 'Сообщить о проблеме' or 'other_options' in context.user_data and context.user_data["other_options"] == 'Сообщить о проблеме':
            do_command(update=update, context=context)
            context.user_data["other_options"] = 'Сообщить о проблеме'
            return

        if text == 'Предложения' or 'other_options' in context.user_data and context.user_data["other_options"] == 'Предложения':
            do_command(update=update, context=context)
            context.user_data["other_options"] = 'Предложения'
            return


        """Выбор темы для изучения"""

        if text == "Теория" or text == 'Назад' and context.user_data['back'] == 'назад теория':
            context.user_data['sort'] = 'Теория'
            keyboard = [
                ["Умножения двоичных чисел"],
                ["Возведение в квадрат"],
                ["Квадратные уравнения"],
                ["Назад"]
            ]
            update.message.reply_text(text='Выберите тему',
                                      reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True,
                                                                       resize_keyboard=True))

            """Выбор темы для практики"""

        elif text == "Практика":
            context.user_data['sort'] = 'Практика'
            keyboard = [
                ["Умножения двоичных чисел"],
                ["Возведение в квадрат"],
                ["Квадратные уравнения"],
                ["Назад"]
            ]
            update.message.reply_text(text='Выберите тему',
                                      reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True,
                                                                       resize_keyboard=True))

        check_practice = 'practice' in context.user_data

        if 'sort' in context.user_data:

            """Блок кода для теории"""

            if update.message.text != 'Теория' and context.user_data['sort'] == 'Теория':

                """Блок умножения двочиных чисел"""

                if text == "Далее" and context.user_data[context.user_data['sort']] == 'Умножения двоичных чисел[2]':
                    context.user_data[context.user_data['sort']] = 'Умножения двоичных чисел[3]'
                    context.user_data['back'] = 'назад теория'
                    context.user_data['practice'] = 'Практика умножение двоичных чисел'
                    context.user_data['sort'] = 'Практика'

                    keyboard = [
                        ["Назад"],
                        ['Практика по теме']
                    ]
                    update.message.reply_text(text=theory_text.THEORY1_3,
                                              reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True,
                                                                               resize_keyboard=True))

                if text == "Далее" and context.user_data[context.user_data['sort']] == 'Умножения двоичных чисел[1]':
                    context.user_data[context.user_data['sort']] = 'Умножения двоичных чисел[2]'

                    keyboard = [
                        ["Далее"]
                    ]
                    update.message.reply_text(text=theory_text.THEORY1_2_1,
                                              reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True,
                                                                               resize_keyboard=True))
                    update.message.reply_text(text=theory_text.THEORY1_2_2,
                                              reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True,
                                                                               resize_keyboard=True))

                if text == "Умножения двоичных чисел":
                    context.user_data['sort'] = 'Теория'
                    context.user_data[context.user_data['sort']] = 'Умножения двоичных чисел[1]'
                    context.user_data['back'] = 'назад теория'

                    keyboard = [
                        ["Далее"],
                        ['Назад']
                    ]
                    update.message.reply_text(text=theory_text.THEORY1_1,
                                              reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True,
                                                                               resize_keyboard=True))

                    """Блок возведения в квадрат"""

                if text == "Далее" and context.user_data[context.user_data['sort']] == 'Возведение в квадрат[5]':
                    context.user_data[context.user_data['sort']] = 'Возведение в квадрат[6]'
                    context.user_data['back'] = 'назад теория'
                    context.user_data['practice'] = 'Практика возведение в квадрат'
                    context.user_data['sort'] = 'Практика'

                    keyboard = [
                        ["Назад"],
                        ['Практика по теме']
                    ]
                    update.message.reply_text(text=theory_text.THEORY2_6,
                                              reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True,
                                                                               resize_keyboard=True))
                    update.message.reply_text(text=theory_text.THEORY2_6_1,
                                              reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True,
                                                                               resize_keyboard=True))
                    update.message.reply_text(text=theory_text.THEORY2_6_2,
                                              reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True,
                                                                               resize_keyboard=True))
                    update.message.reply_text(text=theory_text.THEORY2_6_3,
                                              reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True,
                                                                               resize_keyboard=True))

                if text == "Далее" and context.user_data[context.user_data['sort']] == 'Возведение в квадрат[4]':
                    context.user_data[context.user_data['sort']] = 'Возведение в квадрат[5]'

                    keyboard = [
                        ["Далее"]
                    ]
                    update.message.reply_text(text=theory_text.THEORY2_5,
                                              reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True,
                                                                               resize_keyboard=True))
                    update.message.reply_text(text=theory_text.THEORY2_5_1,
                                              reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True,
                                                                               resize_keyboard=True))
                    update.message.reply_text(text=theory_text.THEORY2_5_2,
                                              reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True,
                                                                               resize_keyboard=True))
                    update.message.reply_text(text=theory_text.THEORY2_5_3,
                                              reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True,
                                                                               resize_keyboard=True))

                if text == "Далее" and context.user_data[context.user_data['sort']] == 'Возведение в квадрат[3]':
                    context.user_data[context.user_data['sort']] = 'Возведение в квадрат[4]'

                    keyboard = [
                        ["Далее"]
                    ]
                    update.message.reply_text(text=theory_text.THEORY2_4,
                                              reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True,
                                                                               resize_keyboard=True))
                    update.message.reply_text(text=theory_text.THEORY2_4_2,
                                              reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True,
                                                                               resize_keyboard=True))
                    update.message.reply_text(text=theory_text.THEORY2_4_3,
                                              reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True,
                                                                               resize_keyboard=True))
                    update.message.reply_text(text=theory_text.THEORY2_4_4,
                                              reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True,
                                                                               resize_keyboard=True))
                if text == "Далее" and context.user_data[context.user_data['sort']] == 'Возведение в квадрат[2]':
                    context.user_data[context.user_data['sort']] = 'Возведение в квадрат[3]'

                    keyboard = [
                        ["Далее"]
                    ]
                    update.message.reply_text(text=theory_text.THEORY2_3,
                                              reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True,
                                                                               resize_keyboard=True))
                    update.message.reply_text(text=theory_text.THEORY2_3_2,
                                              reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True,
                                                                               resize_keyboard=True))
                if text == "Далее" and context.user_data[context.user_data['sort']] == 'Возведение в квадрат[1]':
                    context.user_data[context.user_data['sort']] = 'Возведение в квадрат[2]'

                    keyboard = [
                        ["Далее"]
                    ]
                    update.message.reply_text(text=theory_text.THEORY2_2,
                                              reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True,
                                                                               resize_keyboard=True))

                if text == "Возведение в квадрат":
                    context.user_data['sort'] = 'Теория'
                    context.user_data[context.user_data['sort']] = 'Возведение в квадрат[1]'
                    context.user_data['back'] = 'назад теория'

                    keyboard = [
                        ["Далее"],
                        ['Назад']
                    ]

                    update.message.reply_photo(open('4bedf30c2410ab76334d86f35aaf689c (1).png', 'rb'))

                    update.message.reply_text(text=theory_text.THEORY2_1,
                                              reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True,
                                                                               resize_keyboard=True))

                if text == "Далее" and context.user_data[context.user_data['sort']] == 'Квадратные уравнения[2]':
                    context.user_data[context.user_data['sort']] = 'Квадратные уравнения[3]'
                    context.user_data['back'] = 'назад теория'
                    context.user_data['practice'] = 'Практика квадратные уравнения'
                    context.user_data['sort'] = 'Практика'

                    keyboard = [
                        ["Назад"],
                        ['Практика по теме']
                    ]
                    update.message.reply_text(text=theory_text.THEORY3_3,
                                              reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True,
                                                                               resize_keyboard=True))
                    update.message.reply_text(text=theory_text.THEORY3_3_1,
                                              reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True,
                                                                               resize_keyboard=True))

                    update.message.reply_photo(open('slide-9.jpg', 'rb'))

                    update.message.reply_text(text=theory_text.THEORY3_3_3,
                                              reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True,
                                                                               resize_keyboard=True))
                    update.message.reply_text(text=theory_text.THEORY3_3_4,
                                              reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True,
                                                                               resize_keyboard=True))
                    update.message.reply_text(text=theory_text.THEORY3_3_5,
                                              reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True,
                                                                               resize_keyboard=True))
                    update.message.reply_text(text=theory_text.THEORY3_3_6,
                                              reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True,
                                                                               resize_keyboard=True))
                    update.message.reply_text(text=theory_text.THEORY3_3_7,
                                              reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True,
                                                                               resize_keyboard=True))

                if text == "Далее" and context.user_data[context.user_data['sort']] == 'Квадратные уравнения[1]':
                    context.user_data[context.user_data['sort']] = 'Квадратные уравнения[2]'

                    keyboard = [
                        ["Далее"]
                    ]
                    update.message.reply_text(text=theory_text.THEORY3_2,
                                              reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True,
                                                                               resize_keyboard=True))
                    update.message.reply_text(text=theory_text.THEORY3_2_1,
                                              reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True,
                                                                               resize_keyboard=True))
                    update.message.reply_photo(open('diskriminant-i-korni-kvadratnogo-uravneniya.png', 'rb'))
                    update.message.reply_text(text=theory_text.THEORY3_2_3,
                                              reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True,
                                                                               resize_keyboard=True))
                    update.message.reply_text(text=theory_text.THEORY3_2_4,
                                              reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True,
                                                                               resize_keyboard=True))

                if text == "Квадратные уравнения":
                    context.user_data['sort'] = 'Теория'
                    context.user_data[context.user_data['sort']] = 'Квадратные уравнения[1]'
                    context.user_data['back'] = 'назад теория'

                    keyboard = [
                        ["Далее"],
                        ['Назад']
                    ]
                    update.message.reply_text(text=theory_text.THEORY3_1,
                                              reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True,
                                                                               resize_keyboard=True))

                    update.message.reply_text(text=theory_text.THEORY3_1_1,
                                              reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True,
                                                                               resize_keyboard=True))

                """"Блок кода для практики"""

            elif update.message.text != 'Практика' and context.user_data['sort'] == 'Практика' or \
                    check_practice and context.user_data['practice'] == f'Практика{1 or 2 or 3}':

                if text == 'Назад':
                    do_start(update=update, context=context)
                    return

                """Практика по умножению двоичных чисел"""

                if text == "Умножения двоичных чисел" or text == 'Практика по теме' and context.user_data[
                    'practice'] == 'Практика умножение двоичных чисел' or check_practice and context.user_data[
                    'practice'] == 'Практика1':
                    keyboard = [
                        ["Назад"]
                    ]
                    try:
                        if check_practice and context.user_data['practice'] == 'Практика1':
                            if text != str(context.user_data['first_number'] * context.user_data['second_number']):
                                update.message.reply_text(text=f'Неправильно, попробуйте ещё раз!',
                                                          reply_markup=ReplyKeyboardMarkup(keyboard,
                                                                                           one_time_keyboard=True,
                                                                                           resize_keyboard=True))
                                return
                            else:
                                update.message.reply_text(text=f'Правильно, двигаемся дальше!',
                                                          reply_markup=ReplyKeyboardMarkup(keyboard,
                                                                                           one_time_keyboard=True,
                                                                                           resize_keyboard=True))
                        else:
                            update.message.reply_text(text=f'Итак, давай поупражняемся!',
                                                      reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True,
                                                                                       resize_keyboard=True))
                    except Exception as ex:
                        print('\033[31mLine: ', extract_tb(exc_info()[2])[0][1], '\nException: ', ex)

                    context.user_data['sort'] = 'Практика'
                    context.user_data['practice'] = 'Практика1'

                    first_number = context.user_data['first_number'] = random.randint(10, 99)
                    second_number = context.user_data['second_number'] = random.randint(10, 99)
                    print(f'\033[32m{first_number * second_number}')

                    update.message.reply_text(text=f'{first_number} * {second_number}',
                                              reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True,
                                                                               resize_keyboard=True))

                    """Практика по возведению в квадрат"""

                elif text == "Возведение в квадрат" or text == 'Практика по теме' and context.user_data[
                    'practice'] == 'Практика возведение в квадрат' or check_practice and context.user_data[
                    'practice'] == 'Практика2':

                    keyboard = [
                        ["Назад"]
                    ]
                    try:
                        if check_practice and context.user_data['practice'] == 'Практика2':
                            if text != str(context.user_data['first_number'] * context.user_data['first_number']):
                                update.message.reply_text(text=f'Неправильно, попробуйте ещё раз!',
                                                          reply_markup=ReplyKeyboardMarkup(keyboard,
                                                                                           one_time_keyboard=True,
                                                                                           resize_keyboard=True))
                                return
                            else:
                                update.message.reply_text(text=f'Правильно, двигаемся дальше!',
                                                          reply_markup=ReplyKeyboardMarkup(keyboard,
                                                                                           one_time_keyboard=True,
                                                                                           resize_keyboard=True))
                        else:
                            update.message.reply_text(text=f'Итак, давай поупражняемся!',
                                                      reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True,
                                                                                       resize_keyboard=True))
                    except Exception as ex:
                        print('\033[31mLine: ', extract_tb(exc_info()[2])[0][1], '\nException: ', ex)

                    context.user_data['sort'] = 'Практика'
                    context.user_data['practice'] = 'Практика2'

                    first_number = context.user_data['first_number'] = random.randint(10, 99)
                    print(f'\033[32m{first_number * first_number}')

                    update.message.reply_text(text=f'Квадрат числа: {first_number}',
                                              reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True,
                                                                               resize_keyboard=True))

                    """Практика по квадратным уравнениям"""

                elif text == "Квадратные уравнения" or text == 'Практика по теме' and context.user_data[
                    'practice'] == 'Практика квадратные уравнения' or check_practice and context.user_data[
                    'practice'] == 'Практика3':

                    keyboard = [
                        ["Назад"]
                    ]
                    try:
                        if check_practice and context.user_data['practice'] == 'Практика3':
                            answer1 = f"{context.user_data['x1']} {context.user_data['x2']}"
                            answer2 = f"{context.user_data['x2']} {context.user_data['x1']}"
                            if text != answer1 and text != answer2:
                                update.message.reply_text(text=f'Неправильно, попробуйте ещё раз!',
                                                          reply_markup=ReplyKeyboardMarkup(keyboard,
                                                                                           one_time_keyboard=True,
                                                                                           resize_keyboard=True))
                                return
                            else:
                                update.message.reply_text(text=f'Правильно, двигаемся дальше!',
                                                          reply_markup=ReplyKeyboardMarkup(keyboard,
                                                                                           one_time_keyboard=True,
                                                                                           resize_keyboard=True))
                        else:
                            update.message.reply_text(text=
                                                      f'Итак, давай поупражняемся! В ответ запиши корни через пробел в любом порядке',
                                                      reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True,
                                                                                       resize_keyboard=True))

                    except Exception as ex:
                        print('\033[31mLine: ', extract_tb(exc_info()[2])[0][1], '\nException: ', ex)

                    context.user_data['sort'] = 'Практика'
                    context.user_data['practice'] = 'Практика3'

                    """Генератор создания квадратного уравнения"""

                    while True:
                        A_coefficient = random.randint(1, 2)
                        B_coefficient = random.randint(-25, 25)
                        C_coefficient = random.randint(-40, 40)
                        D = B_coefficient ** 2 - 4 * A_coefficient * C_coefficient
                        sq_D = D ** 0.5
                        if D > 0 and sq_D - int(sq_D) == 0.0:
                            x1_pre = (-1 * B_coefficient + sq_D) / (2 * A_coefficient)
                            x2_pre = (-1 * B_coefficient - sq_D) / (2 * A_coefficient)
                            if abs(x1_pre) - abs(int(x1_pre)) == 0.0 and abs(x2_pre) - abs(int(x2_pre)) == 0.0:
                                x1 = context.user_data['x1'] = int(x1_pre)
                                x2 = context.user_data['x2'] = int(x2_pre)
                                print(f'\033[32m{x1, x2}')
                                break

                    mes_text = f'Найдите корни уравнения: {A_coefficient if A_coefficient == 2 else ""}x^2 ' \
                               f'{"-" if B_coefficient < 0 else "+"} {abs(B_coefficient)}x {"-" if C_coefficient < 0 else "+"}' \
                               f' {abs(C_coefficient)}'
                    update.message.reply_text(text=mes_text,
                                              reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True,
                                                                               resize_keyboard=True))

    except Exception as ex:
        print('\033[31mLine: ', extract_tb(exc_info()[2])[0][1], '\nException: ', ex)

        """Функция обработчика команд"""


def do_command(update, context: CallbackContext):
    if update.message.text == '/start':
        do_start(update, context)
        return

    if update.message.text == '/addven':
        context.user_data['command'] = update.message.text
        update.message.reply_text(text='Введите артикул:')

    if update.message.text == 'Сообщить о проблеме':
        context.user_data['command'] = '/addven'
        update.message.reply_text(text='Опишите вашу проблему, а так же ваши действия, которые привели вас к ошибке:')

    if update.message.text == 'Предложения':
        context.user_data['command'] = '/addkey'
        update.message.reply_text(text='Напишите ваше предложение в улучшении бота:')

    elif update.message.text == '/delven':
        context.user_data['command'] = update.message.text
        update.message.reply_text(text='Введите артикул для удаления:')

    elif update.message.text == '/addkey':
        context.user_data['command'] = update.message.text
        update.message.reply_text(text='Введите ключевое слово:')

    elif update.message.text == '/delkey':
        context.user_data['command'] = update.message.text
        update.message.reply_text(text='Введите ключевое слово для удаления:')

    if 'command' in context.user_data:

        if update.message.text != '/delven' and context.user_data['command'] == '/delven':
            do_words(update, context)

        elif update.message.text != '/addven' and update.message.text != 'Сообщить о проблеме' and context.user_data[
            'command'] == '/addven':
            do_words(update, context)

        elif update.message.text != '/addkey' and update.message.text != 'Предложения' and context.user_data[
            'command'] == '/addkey':
            do_words(update, context)

        elif update.message.text != '/delkey' and context.user_data['command'] == '/delkey':
            do_words(update, context)

    else:
        update.message.reply_text(text='что?...')

    """Диспетчер команд"""


def do_words(update: Update, context: CallbackContext):
    if context.user_data['command'] == '/addven':
        do_add_vendor(update, context)

    elif context.user_data['command'] == '/delven':
        do_delete_vendor(update, context)

    elif context.user_data['command'] == '/addkey':
        do_add_keyword(update, context)

    elif context.user_data['command'] == '/delkey':
        do_delete_keyword(update, context)
    else:
        update.message.reply_text(text='sry...')

    context.user_data['command'] = None
    return context.user_data['command']


def do_add_vendor(update: Update, context):
    for i in range(2, 100):

        if vendor.cell(column=1, row=i).value is None:
            vendor.cell(column=1, row=i).value = update.message.text
            vendor.cell(column=2,
                        row=i).value = f'{update.message.from_user.first_name} {update.message.from_user.last_name}'

            update.message.reply_text(text='Большое спасибо, мы это исправим!')
            break

    return wb.save('database.xlsx')


def do_delete_vendor(update: Update, context):
    for i in range(2, 100):

        if vendor.cell(column=1, row=i).value == update.message.text:
            vendor.cell(column=1, row=i).value = None
            update.message.reply_text(text='Готово!')
            break

    return wb.save('database.xlsx')


def do_add_keyword(update: Update, context):
    for i in range(2, 100):

        if vendor.cell(column=3, row=i).value is None:
            vendor.cell(column=3, row=i).value = update.message.text
            vendor.cell(column=4,
                        row=i).value = f'{update.message.from_user.first_name} {update.message.from_user.last_name}'
            update.message.reply_text(text='Большое вам спасибо!')
            break

    return wb.save('database.xlsx')


def do_delete_keyword(update: Update, context):
    for i in range(2, 100):

        if vendor.cell(column=2, row=i).value == update.message.text:
            vendor.cell(column=2, row=i).value = None
            update.message.reply_text(text='Готово!')
            break

    return wb.save('database.xlsx')


# if __name__ == '__main__':
main()
