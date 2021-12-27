from telegram import Update, ReplyKeyboardMarkup, ReplyKeyboardRemove
from telegram.ext import Updater, MessageHandler, Filters, CommandHandler, CallbackContext, CallbackQueryHandler
from openpyxl import load_workbook
from telegram import InlineKeyboardButton, InlineKeyboardMarkup

wb = load_workbook('database.xlsx')
vendor = wb['users']
token = '1684008354:AAGtETeI2tgXd-AM_6AQvJmh43aTSYjkXxE'


def main():

    updater = Updater(token=token, use_context=True)
    dispatcher = updater.dispatcher

    handler = MessageHandler(Filters.command, do_command)

    # start_handler = CommandHandler('start', do_start)
    # help_handler = CommandHandler('help', do_help)
    keyboard_handler = MessageHandler(Filters.text, keyboard_value)

    dispatcher.add_handler(handler)
    dispatcher.add_handler(keyboard_handler)

    updater.start_polling()
    updater.idle()

def do_start(update, context):

    keyboard = [
        ["Теория"],
        ["Практика"]
    ]

    update.message.reply_text(text='Добро пожаловать, выберите, что бы вы хотели сделать.',
                              reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True, resize_keyboard=True))

def keyboard_value(update, context):
    text = update.message.text

    if text == "Теория":
        context.user_data['sort'] = 'Теория'
        keyboard = [
            ["Умножения двоичных чисел"],
            ["Возведение в квадрат"],
            ["Квадратные уравнения"]
        ]
        update.message.reply_text(text='Выберите тему',reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True, resize_keyboard=True))

    elif text == "Практика":
        context.user_data['sort'] = 'Практика'
        keyboard = [
            ["Умножения двоичных чисел"],
            ["Возведение в квадрат"],
            ["Квадратные уравнения"]
        ]
        update.message.reply_text(text='Выберите тему',reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True, resize_keyboard=True))

    if 'sort' in context.user_data:

        if update.message.text != 'Теория' and context.user_data['command'] == 'Теория':

            if text == "Умножения двоичных чисел":

                context.user_data['sort'] = 'Теория'
                keyboard = [
                    ["Далее"]
                ]
                update.message.reply_text(text='Выберите тему',
                                          reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True,
                                                                           resize_keyboard=True))

            elif text == "Возведение в квадрат":

                context.user_data['sort'] = 'Практика'
                keyboard = [
                    ["Умножения двоичных чисел"],
                    ["Возведение в квадрат"],
                    ["Квадратные уравнения"]
                ]
                update.message.reply_text(text='Выберите тему',
                                          reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True,
                                                                           resize_keyboard=True))

            elif text == "Квадратные уравнения":

                context.user_data['sort'] = 'Практика'
                keyboard = [
                    ["Умножения двоичных чисел"],
                    ["Возведение в квадрат"],
                    ["Квадратные уравнения"]
                ]
                update.message.reply_text(text='Выберите тему',
                                          reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True,
                                                                           resize_keyboard=True))

        elif update.message.text != 'Практика' and context.user_data['command'] == 'Практика':

            if text == "Умножения двоичных чисел":

                context.user_data['sort'] = 'Теория'
                keyboard = [
                    ["Умножения двоичных чисел"],
                    ["Возведение в квадрат"],
                    ["Квадратные уравнения"]
                ]
                update.message.reply_text(text='Выберите тему',
                                          reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True,
                                                                           resize_keyboard=True))

            elif text == "Возведение в квадрат":

                context.user_data['sort'] = 'Практика'
                keyboard = [
                    ["Умножения двоичных чисел"],
                    ["Возведение в квадрат"],
                    ["Квадратные уравнения"]
                ]
                update.message.reply_text(text='Выберите тему',
                                          reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True,
                                                                           resize_keyboard=True))

            elif text == "Квадратные уравнения":

                context.user_data['sort'] = 'Практика'
                keyboard = [
                    ["Умножения двоичных чисел"],
                    ["Возведение в квадрат"],
                    ["Квадратные уравнения"]
                ]
                update.message.reply_text(text='Выберите тему',
                                          reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True,
                                                                           resize_keyboard=True))


def do_command(update, context: CallbackContext):
    if update.message.text == '/start':
        do_start(update, context)
        return

    if update.message.text == '/addven':
        context.user_data['command'] = update.message.text
        update.message.reply_text(text='Введите артикул:')

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

        elif update.message.text != '/addven' and context.user_data['command'] == '/addven':
            do_words(update, context)

        elif update.message.text != '/addkey' and context.user_data['command'] == '/addkey':
            do_words(update, context)

        elif update.message.text != '/delkey' and context.user_data['command'] == '/delkey':
            do_words(update, context)

    else:
        update.message.reply_text(text='что?...')


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
            update.message.reply_text(text='Готово!')
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

        if vendor.cell(column=2, row=i).value is None:
            vendor.cell(column=2, row=i).value = update.message.text
            update.message.reply_text(text='Готово!')
            break

    return wb.save('database.xlsx')


def do_delete_keyword(update: Update, context):

    for i in range(2, 100):

        if vendor.cell(column=2, row=i).value == update.message.text:
            vendor.cell(column=2, row=i).value = None
            update.message.reply_text(text='Готово!')
            break

    return wb.save('database.xlsx')


main()