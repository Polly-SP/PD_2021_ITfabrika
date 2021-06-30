import telebot
import openpyxl


#bot = telebot.TeleBot('1773022199:AAFICnJg_0DLHYhNrQchvq4cy4v7fqjqc6k')
bot = telebot.TeleBot('1713340125:AAHqPMOqd8_uJPaabxrV7VE95Zg50UzzT8I')
bot.remove_webhook()


file = openpyxl.open("goods4bot.xlsx", read_only=True)
sheet = file.active


@bot.message_handler(commands=['start'])
def process_start(message):
    a = telebot.types.ReplyKeyboardRemove()
    bot.send_message(message.from_user.id, "1. Для поиска товара по названию выбери пункт в меню ниже", reply_markup=a)
    keyboard = telebot.types.ReplyKeyboardMarkup(True)
    keyboard.row('Ламинат')
    keyboard.row('Линолеум')
    msg = bot.send_message(message.chat.id, text= "2. Для поиска товара по артикулу введи артикул соответствующего товара.", reply_markup=keyboard)


@bot.message_handler(content_types=['text'])
def step1(message):
    for row in range(1, sheet.max_row + 1):
        if str(message.text) == str(sheet[row][3].value):
            bot.send_message(message.from_user.id, f"Название: {sheet[row][0].value} \n "
                                                   f"Коллекция: {sheet[row][1].value} \n "
                                                   f"Тон/Рулон: {sheet[row][2].value} \n "
                                                   f"Артикул: {sheet[row][3].value} \n "
                                                   f"Штрихкод: {sheet[row][4].value}")
            bot.send_photo(message.from_user.id,   f"{sheet[row][5].value}")

    menu1 = telebot.types.InlineKeyboardMarkup()
    menu1.add(telebot.types.InlineKeyboardButton(text='Floorwood', callback_data='Floorwood'))
    menu1.add(telebot.types.InlineKeyboardButton(text='Balterio', callback_data='Balterio'))
    menu1.add(telebot.types.InlineKeyboardButton(text='Vitality', callback_data='Vitality'))
    menu1.add(telebot.types.InlineKeyboardButton(text='Faus', callback_data='Faus'))
    if message.text == 'Ламинат':
        msg = bot.send_message(message.chat.id, text='Выберите коллекцию', reply_markup=menu1)

    menu2 = telebot.types.InlineKeyboardMarkup()
    menu2.add(telebot.types.InlineKeyboardButton(text='Venturi Spark/001 Ravena', callback_data='Venturi Spark/001 Ravena'))
    menu2.add(telebot.types.InlineKeyboardButton(text='Samson Spark/3 Samson', callback_data='Samson Spark/3 Samson'))
    if message.text == 'Линолеум':
        msg = bot.send_message(message.chat.id, text='Выберите коллекцию', reply_markup=menu2)


@bot.callback_query_handler(func=lambda call: True)
def step2(call):
    menu3 = telebot.types.InlineKeyboardMarkup()
    menu3.add(telebot.types.InlineKeyboardButton(text='Floorwood Estet', callback_data='Floorwood Estet'))
    menu3.add(telebot.types.InlineKeyboardButton(text='Floorwood Epica', callback_data='Floorwood Epica'))

    menu4 = telebot.types.InlineKeyboardMarkup()
    menu4.add(telebot.types.InlineKeyboardButton(text='Faus Elegance', callback_data='Faus Elegance'))

    menu5 = telebot.types.InlineKeyboardMarkup()
    menu5.add(telebot.types.InlineKeyboardButton(text='Vitality Deluxe', callback_data='Vitality Deluxe'))

    menu6 = telebot.types.InlineKeyboardMarkup()
    menu6.add(telebot.types.InlineKeyboardButton(text='Balterio Impressio', callback_data='Balterio Impressio'))
    menu6.add(telebot.types.InlineKeyboardButton(text='Balterio Fortissimo', callback_data='Balterio Fortissimo'))

    if call.data == 'Floorwood':
        msg = bot.send_message(call.message.chat.id, 'Выберете тип', reply_markup=menu3)

    elif call.data == 'Faus':
        msg = bot.send_message(call.message.chat.id, 'Выберете тип', reply_markup=menu4)

    elif call.data == 'Vitality':
        msg = bot.send_message(call.message.chat.id, 'Выберете тип', reply_markup=menu5)

    elif call.data == 'Balterio':
        msg = bot.send_message(call.message.chat.id, 'Выберете тип', reply_markup=menu6)

    menu7 = telebot.types.InlineKeyboardMarkup()
    menu7.add(telebot.types.InlineKeyboardButton(text='116,10 кв.м', callback_data='116,10 кв.м'))
    menu7.add(telebot.types.InlineKeyboardButton(text='119,34 кв.м', callback_data='119,34 кв.м'))
    menu7.add(telebot.types.InlineKeyboardButton(text='120,00 кв.м', callback_data='120,00 кв.м'))

    menu8 = telebot.types.InlineKeyboardMarkup()
    menu8.add(telebot.types.InlineKeyboardButton(text='105,30 кв.м', callback_data='105,30 кв.м'))
    menu8.add(telebot.types.InlineKeyboardButton(text='150,00 кв.м', callback_data='150,00 кв.м'))

    if call.data == 'Venturi Spark/001 Ravena':
        msg = bot.send_message(call.message.chat.id, 'Выберете размер', reply_markup=menu7)
    elif call.data == 'Samson Spark/3 Samson':
        msg = bot.send_message(call.message.chat.id, 'Выберете размер', reply_markup=menu8)

    menu9 = telebot.types.InlineKeyboardMarkup()
    menu9.add(telebot.types.InlineKeyboardButton(text='Дуб Бэкстер', callback_data='Дуб Бэкстер'))
    menu9.add(telebot.types.InlineKeyboardButton(text='Дуб Иберо грей', callback_data='Дуб Иберо грей'))
    menu9.add(telebot.types.InlineKeyboardButton(text='Дуб Савой', callback_data='Дуб Савой'))
    menu9.add(telebot.types.InlineKeyboardButton(text='Дуб Ленсингтон', callback_data='Дуб Ленсингтон'))
    menu9.add(telebot.types.InlineKeyboardButton(text='Дуб Энтони', callback_data='Дуб Энтони'))

    menu10 = telebot.types.InlineKeyboardMarkup()
    menu10.add(telebot.types.InlineKeyboardButton(text='Дуб Винсент', callback_data='Дуб Винсент'))
    menu10.add(telebot.types.InlineKeyboardButton(text='Дуб Ануари ', callback_data='Дуб Ануари'))
    menu10.add(telebot.types.InlineKeyboardButton(text='Дуб Грюйер', callback_data='Дуб Грюйер'))

    menu11 = telebot.types.InlineKeyboardMarkup()
    menu11.add(telebot.types.InlineKeyboardButton(text='Divino Oak', callback_data='Divino Oak'))
    menu11.add(telebot.types.InlineKeyboardButton(text='Romance Oak', callback_data='Romance Oak'))
    menu11.add(telebot.types.InlineKeyboardButton(text='Colonial Oak', callback_data='Colonial Oak'))

    menu12 = telebot.types.InlineKeyboardMarkup()
    menu12.add(telebot.types.InlineKeyboardButton(text='Дуб шато', callback_data='Дуб шато'))
    menu12.add(telebot.types.InlineKeyboardButton(text='Дуб Амбарный', callback_data='Дуб Амбарный'))
    menu12.add(telebot.types.InlineKeyboardButton(text='Дуб отбеленный', callback_data='Дуб отбеленный'))
    menu12.add(telebot.types.InlineKeyboardButton(text='Орех Селект', callback_data='Орех Селект'))
    menu12.add(telebot.types.InlineKeyboardButton(text='Дуб Лакированный натуральный',
                                                  callback_data='Дуб Лакированный натуральный'))
    menu12.add(telebot.types.InlineKeyboardButton(text='Дуб белый промасленный', callback_data='Дуб белый промасленный'))
    menu12.add(telebot.types.InlineKeyboardButton(text='Дуб Песчаный', callback_data='Дуб Песчаный'))
    menu12.add(telebot.types.InlineKeyboardButton(text='Дуб Золотой закат', callback_data='Дуб Золотой закат'))
    menu12.add(telebot.types.InlineKeyboardButton(text='Дуб серовато-дымчатый', callback_data='Дуб серовато-дымчатый'))
    menu12.add(telebot.types.InlineKeyboardButton(text='Дуб Шамо', callback_data='Дуб Шамо'))

    menu13 = telebot.types.InlineKeyboardMarkup()
    menu13.add(telebot.types.InlineKeyboardButton(text='Дуб Wadi Rum', callback_data='Дуб Wadi Rum'))
    menu13.add(telebot.types.InlineKeyboardButton(text='Дуб Фраппучино', callback_data='Дуб Фраппучино'))
    menu13.add(telebot.types.InlineKeyboardButton(text='Дуб Гарда', callback_data='Дуб Гарда'))
    menu13.add(telebot.types.InlineKeyboardButton(text='Дуб Коричнево-Дымчатый', callback_data='Дуб Коричнево-Дымчатый'))
    menu13.add(telebot.types.InlineKeyboardButton(text='Дуб с Подпалиной', callback_data='Дуб с Подпалиной'))
    menu13.add(telebot.types.InlineKeyboardButton(text='Дуб Каспий', callback_data='Дуб Каспий'))
    menu13.add(telebot.types.InlineKeyboardButton(text='Дуб Саванна', callback_data='Дуб Саванна'))
    menu13.add(telebot.types.InlineKeyboardButton(text='Дуб Слоновая кость', callback_data='Дуб Слоновая кость'))

    menu14 = telebot.types.InlineKeyboardMarkup()
    menu14.add(telebot.types.InlineKeyboardButton(text='Дуб Фуджи', callback_data='Дуб Фуджи'))
    menu14.add(telebot.types.InlineKeyboardButton(text='Дуб Килиманджаро', callback_data='Дуб Килиманджаро'))
    menu14.add(telebot.types.InlineKeyboardButton(text='Дуб Гималайя', callback_data='Дуб Гималайя'))
    menu14.add(telebot.types.InlineKeyboardButton(text='Дуб Этна', callback_data='Дуб Этна'))

    if call.data == 'Floorwood Estet':
        msg = bot.send_message(call.message.chat.id, 'Выберете материал', reply_markup=menu9)

    elif call.data == 'Floorwood Epica':
        msg = bot.send_message(call.message.chat.id, 'Выберете материал', reply_markup=menu10)

    elif call.data == 'Faus Elegance':
        msg = bot.send_message(call.message.chat.id, 'Выберете материал', reply_markup=menu11)

    elif call.data == 'Vitality Deluxe':
        msg = bot.send_message(call.message.chat.id, 'Выберете материал', reply_markup=menu12)

    elif call.data == 'Balterio Impressio':
        msg = bot.send_message(call.message.chat.id, 'Выберете материал', reply_markup=menu13)

    elif call.data == 'Balterio Fortissimo':
        msg = bot.send_message(call.message.chat.id, 'Выберете материал', reply_markup=menu14)

    keyboard = telebot.types.ReplyKeyboardMarkup(True)
    for row in range(1, sheet.max_row + 1):
        if sheet[row][2].value.count(call.data):
            bot.send_message(call.from_user.id,
                             f"Название: {sheet[row][0].value} \n "
                             f"Коллекция: {sheet[row][1].value} \n "
                             f"Тон/Рулон: {sheet[row][2].value} \n "
                             f"Артикул: {sheet[row][3].value} \n "
                             f"Штрихкод: {sheet[row][4].value}")
            bot.send_photo(call.from_user.id,   f"{sheet[row][5].value}")


bot.polling(none_stop=True)