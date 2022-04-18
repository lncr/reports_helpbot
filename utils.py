from telebot import types


def create_lang_inlines(language):
    assert language in ['python', 'javascript']

    class_9 = types.InlineKeyboardButton(text='9 класс', callback_data=f'class_9_{language}')
    class_11 = types.InlineKeyboardButton(text='11 класс', callback_data=f'class_11_{language}')
    main_menu = types.InlineKeyboardButton(text='Вернуться в главное меню', callback_data='main_menu')
    return [[class_9], [class_11], [main_menu]]
