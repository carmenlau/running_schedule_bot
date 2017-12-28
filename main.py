import os
import telebot
from telebot import types
from data import districts as districts_data

bot = telebot.TeleBot(os.environ.get('API_KEY'))
checking_dict = {}

@bot.message_handler(commands=['start', 'help'])
def send_welcome(message):
    bot.reply_to(message, """Support commands:
    /check - Start checking
    """)

@bot.message_handler(commands=['check'])
def start_checking(message):
    chat_id = message.chat.id
    checking_dict[chat_id] = {}

    markup = types.ReplyKeyboardMarkup(one_time_keyboard=True)
    markup.add(*[d['district'] for d in districts_data])
    msg = bot.reply_to(message, '選擇地區', reply_markup=markup)
    bot.register_next_step_handler(msg, process_district)


def process_district(message):
    # save district
    district = message.text
    chat_id = message.chat.id
    checking_dict[message.chat.id]['district'] = district

    # ask areas
    places = next((d['places'] for d in districts_data if d['district'] == district), None)
    if places is None:
        raise Exception()

    # handle no places
    if not len(places):
        places.append(district)

    markup = types.ReplyKeyboardMarkup(one_time_keyboard=True)
    markup.add(*places)
    msg = bot.reply_to(message, '選擇運動場', reply_markup=markup)
    bot.register_next_step_handler(msg, process_area)

def process_area(message):
    bot.reply_to(message, 'Test: ' + message.text)

bot.polling()
