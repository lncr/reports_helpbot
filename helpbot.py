import os
import telebot
import constants
import openpyxl
from telebot import types
from flask import Flask, request

from utils import create_lang_inlines

TOKEN = constants.TOKEN
bot = telebot.TeleBot(TOKEN, parse_mode='HTML')
server = Flask(__name__)


main_menu = types.InlineKeyboardButton(text='Вернуться в главное меню', callback_data='main_menu')
langs = types.InlineKeyboardButton(text='Языки программирования ', callback_data='languages')
ages = types.InlineKeyboardButton(text='Ограничения по возрасту', callback_data='ages')
ort = types.InlineKeyboardButton(text='Нужен ли ОРТ?', callback_data='ort')
wanna_langs = types.InlineKeyboardButton(text='Каким языкам программирования смогу научиться?', callback_data='languages')
python = types.InlineKeyboardButton(text='Python', callback_data='python')
javascript = types.InlineKeyboardButton(text='JavaScript', callback_data='javascript')
education_program = types.InlineKeyboardButton(text='Программа обучения', callback_data='education_program')
education_duration = types.InlineKeyboardButton(text='Длительность обучения', callback_data='education_duration')
prev_menu = types.InlineKeyboardButton(text='Назад', callback_data='previous')
contract_group = types.InlineKeyboardButton(text='Контрактные групы', callback_data='contract_group')
budget_group = types.InlineKeyboardButton(text='Бюджетные группы', callback_data='budget_group')
documents = types.InlineKeyboardButton(text='Документы  для поступления', callback_data='documents')
laptop = types.InlineKeyboardButton(text='Характеристики ноутбука', callback_data='laptop')
dorm = types.InlineKeyboardButton(text='Общежитие', callback_data='dorm')


CONTENT = {
	'languages': (constants.LANGUAGES, [[python], [javascript], [main_menu]]),
	'ages': (constants.AGES, [[python], [javascript], [main_menu]]),
	'ort': (constants.ORT, [[python], [javascript], [main_menu]]),
	'main_menu': ('Что тебя интересует?', [[langs], [ages], [ort]]),
	'javascript': (constants.AGE_BASED, create_lang_inlines('javascript')),
	'python': (constants.AGE_BASED, create_lang_inlines('python')),
	'class_9_javascript': (constants.CLASS_9_JAVASCRIPT, [[education_program], [education_duration], [main_menu]]),
	'class_11_javascript': (constants.CLASS_11_JAVASCRIPT, [[education_program], [education_duration], [main_menu]]),
	'class_9_python': (constants.CLASS_9_PYTHON, [[education_program], [education_duration], [main_menu]]),
	'class_11_python': (constants.CLASS_11_PYTHON, [[education_program], [education_duration], [main_menu]]),
	'education_program': (constants.EDUCATION_PROGRAM, [[main_menu]]),
	'education_duration': (constants.EDUCATION_DURATION, [[contract_group], [budget_group], [main_menu]]),
	'contract_group': (constants.CONTRACT_GROUP, [[documents], [laptop], [dorm], [main_menu]]),
	'budget_group': (constants.BUDGET_GROUP, [[documents], [laptop], [dorm], [main_menu]]),
	'documents': (constants.DOCUMENTS, [[main_menu]]),
	'laptop': (constants.LAPTOP, [[main_menu]]),
	'dorm': (constants.DORM, [[main_menu]])
}
history = []
user_dict = {}


def write_to_xlsx(user):
	wb = openpyxl.load_workbook(filename='stats.xlsx')
	ws = wb.worksheets[0]
	max_row = ws.max_row
	for i in range(2, ws.max_row+1):
		if ws.cell(row=i, column=1).value == user.telegram_id:
			ws[f'A{i}'] = Noneuser.telegram_id
			ws[f'B{i}'] = user.telegram_username
			ws[f'C{i}'] = user.name
			ws[f'D{i}'] = user.credential
			break
	else:
		ws[f'A{max_row+1}'] = user.telegram_id
		ws[f'B{max_row+1}'] = user.telegram_username
		ws[f'C{max_row+1}'] = user.name
		ws[f'D{max_row+1}'] = user.credential
	wb.save(filename='stats.xlsx')


class User:
	def __init__(self, name, telegram_username=None, telegram_id=None):
		self.name = name
		self.telegram_username = telegram_username
		self.telegram_id = telegram_id
		self.credential = None


def process_name_step(message):
	chat_id = message.chat.id
	name = message.text
	user = User(name, message.from_user.username, message.from_user.id)
	user_dict[chat_id] = user
	msg = bot.reply_to(message, 'Оставь свой номер или почту, отправим уведомление о старте набора в группы')
	bot.register_next_step_handler(msg, process_credential_step)


def process_credential_step(message):
	chat_id = message.chat.id
	credential = message.text
	user = user_dict[chat_id]
	user.credential = credential
	write_to_xlsx(user)
	reply_markup = types.InlineKeyboardMarkup(CONTENT['main_menu'][1])
	bot.send_message(message.chat.id, CONTENT['main_menu'][0], reply_markup=reply_markup)


@bot.message_handler(commands=['start'])
def send_welcome(message):
	msg = bot.reply_to(message, constants.WELCOME_MSG)
	bot.register_next_step_handler(msg, process_name_step)


@bot.message_handler(commands=['loadstats'])
def load_stats(message):
	doc = open('stats.xlsx', 'rb')
	bot.send_document(message.chat.id, doc)


@bot.callback_query_handler(func=lambda call:True)
def answer(call):

	global history

	if call.data == 'previous':
		history.pop()
		content = CONTENT[history[-1]] if history else CONTENT['main_menu']
		keyboards = content[1]

		if history and not len(keyboards[-1]) > 1:
			keyboards[-1].insert(0, prev_menu)

		reply_markup = types.InlineKeyboardMarkup(keyboards)
		bot.send_message(call.message.chat.id, content[0], reply_markup=reply_markup)

	elif call.data == 'main_menu':
		history = []
		reply_markup = types.InlineKeyboardMarkup(CONTENT['main_menu'][1])
		
		bot.send_message(call.message.chat.id, CONTENT['main_menu'][0], reply_markup=reply_markup)

	else:
		content = CONTENT[call.data]
		text = content[0]
		keyboards = content[1]

		if history and not len(keyboards[-1]) > 1:
			keyboards[-1].insert(0, prev_menu)

		reply_markup = types.InlineKeyboardMarkup(keyboards)
		history.append(call.data)
		bot.send_message(call.message.chat.id, text, reply_markup=reply_markup)

	bot.answer_callback_query(call.id)

@server.route('/' + TOKEN, methods=['POST'])
def getMessage():
	bot.process_new_updates([types.Update.de_json(request.stream.read().decode('utf-8'))])
	return "!", 200

@server.route('/')
def webhook():
	bot.remove_webhook()
	bot.set_webhook(url='https://sheltered-plains-90885.herokuapp.com/'+TOKEN)
	return "!", 200

if __name__ == '__main__':
	server.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
