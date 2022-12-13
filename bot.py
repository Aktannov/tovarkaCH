from aiogram import Dispatcher, types, executor, Bot
from openpyxl import load_workbook, Workbook
from config import load_config
from data import data_by_id, data_to_notifications, dataXlsx
from connect_db import subscriber_exists, add_subscriber, update_subscription, subscriber_status, get_subscriptions, get_len, add_len
import asyncio
import logging
import keyboards as kb


wbSearch = load_workbook('data.xlsx')
sheet = wbSearch.active


logging.basicConfig(level=logging.INFO)

# Подключаемся к боту
token = load_config().get('Token')
bot = Bot(token=token)
dp = Dispatcher(bot)


@dp.message_handler(commands=['start'])
async def start(message: types.Message):
    await message.reply('Выберите язык!', reply_markup=kb.greet_kb)



# Обрабатываем стартовое сообщение
@dp.message_handler(commands=['KGZ🇰🇬', 'RU🇷🇺'])
async def len(message: types.Message):
    # print(message.text)
    if message.text == '/KGZ🇰🇬':
        if not get_len(message.from_user.id):
            print('yes')
            add_len(message.from_user.id, 'kgz')
        await message.answer(text='Саламатсызбы!\nТоварларыңызды көзөмөлдөө үчүн жазыңыз /info сиздин жеке '
                                'номериңиз\nКатуу форматта болуш керек GX-0000 же 0000!!!\nЭгер буюмуңуз жөнүндө эскертмелерди алгыңыз келсе, анда /subscribe GX-0000')
    elif message.text == '/RU🇷🇺':
        if not get_len(message.from_user.id):
            add_len(message.from_user.id, 'ru')
        await message.answer(text='Здравствуйте!\nДля отслеживания ваших товаров введите /info ваш персональный '
                                'номер\nСтрого в фортмате GX-0000 или 0000!!!\nВЕсли хотите получать уведомления о вашем товар, то /subscribe GX-0000')



# Обрабатываем
@dp.message_handler(commands='info')
async def info(message: types.Message):
    # RU
    if get_len(message.from_user.id)[0][0] == 'ru':
        id = str(message.text[6:])
        if id.__len__() == 7 and id[:3] == 'GX-':
            """Приводим все возможные виды GX кода"""
            ex1 = id
            ex2 = id[3:]
            ex3 = id[4:]
            ex4 = id[5:]

            await message.answer(text='Подождите минутку...')

            data = asyncio.create_task(data_by_id(ex1, ex2, ex3, ex4))
            await data


            """Проводим проверку и отправляем результат"""
            if not data.result()[0]:
                await message.answer(text='Товаров с таким кодом нет!')
            else:
                for i in range((data.result().__len__())):
                    await message.answer(text=data.result()[i])
        elif id.__len__() == 4:
            """Приводим все возможные виды GX кода"""
            ex1 = 'GX-' + id
            ex2 = id
            ex3 = id[1:]
            ex4 = id[2:]


            await message.answer(text='Подождите минутку...')

            data = asyncio.create_task(data_by_id(ex1, ex2, ex3, ex4))

            await data


            """Проводим проверку и отправляем результат"""
            if not data.result()[0]:
                await message.answer(text='Товаров с таким кодом нет!')
            else:
                for i in range((data.result().__len__())):
                    await message.answer(text=data.result()[i])
        else:
            await message.answer(text='Строго в фортмате GX-0000 или 0000!!!')

    # KGZ
    elif get_len(message.from_user.id)[0][0] == 'kgz':
        id = message.text[6:]
        if len(id) == 7 and id[:3] == 'GX-':
            """Приводим все возможные виды GX кода"""
            ex1 = id
            ex2 = id[3:]
            ex3 = id[4:]
            ex4 = id[5:]

            await message.answer(text='Бир мүнөт күтөө туруңуз...')

            data = asyncio.create_task(data_by_id(ex1, ex2, ex3, ex4))

            await data
            print(len(data.result()))

            """Проводим проверку и отправляем результат"""
            if not data.result()[0]:
                await message.answer(text='Мындай коду бар товарлар жок!')
            else:
                for i in range(len(data.result())):
                    await message.answer(text=data.result()[i])
        elif len(id) == 4:
            """Приводим все возможные виды GX кода"""
            ex1 = 'GX-' + id
            ex2 = id
            ex3 = id[1:]
            ex4 = id[2:]
            print(ex3, ex4)

            await message.answer(text='Бир мүнөт күтөө туруңуз...')

            data = asyncio.create_task(data_by_id(ex1, ex2, ex3, ex4))

            await data
            print(len(data.result()))

            """Проводим проверку и отправляем результат"""
            if not data.result()[0]:
                await message.answer(text='Мындай коду бар товарлар жок!')
            else:
                for i in range(len(data.result())):
                    await message.answer(text=data.result()[i])
        else:
            await message.answer(text='Катуу форматта болуш керек GX-0000 же 0000!!!')



# Обрабатываем активацию подписки
@dp.message_handler(commands=['subscribe'])
async def subscribe(message):
    if get_len(message.from_user.id) == 'ru':
        if (str(message.text)).__len__() != 18:
            await message.answer('Запрос должен быть формата /subscribe GX-0000')
        gx = message.text[11:]
        if not subscriber_exists(message.from_user.id, gx):
            """Если пользователя нет в базе данных"""
            add_subscriber(message.from_user.id, gx)
            await message.answer('Вы подписались на уведомления!\nЕсли товар придет, вы сразу же об этом узнаете!')
        elif subscriber_exists(message.from_user.id, gx) and not subscriber_status(message.from_user.id):
            """Если пользователь есть в базе данных и неактивен"""
            update_subscription(message.from_user.id, True)
            await message.answer('Вы снова подписались на уведомления!\nЕсли товар придет, вы сразу же об этом узнаете!')
        else:
            """Если пользователь есть в базе данных и активен"""
            await message.answer('Вы уже подписаны!')

            
    elif get_len(message.from_user.id) == 'kgz':
        if (str(message.text)).__len__() != 18:
            await message.answer('Катуу форматта болуш керек /subscribe GX-0000')
        gx = message.text[11:]
        if not subscriber_exists(message.from_user.id, gx):
            """Если пользователя нет в базе данных"""
            add_subscriber(message.from_user.id, gx)
            await message.answer('Сиз эскертмелерге жазылдыңыз!\nЭгерде продукт келсе, сиз бул жөнүндө дароо билесиз!')
        elif subscriber_exists(message.from_user.id, gx) and not subscriber_status(message.from_user.id):
            """Если пользователь есть в базе данных и неактивен"""
            update_subscription(message.from_user.id, True)
            await message.answer('Сиз эскертмелерге жазылдыңыз кайта!\nЭгерде продукт келсе, сиз бул жөнүндө дароо билесиз!')
        else:
            """Если пользователь есть в базе данных и активен"""
            await message.answer('Сиз эскертмелерге уже жазылдыңыз!')


# Обрабатываем деактивацию подписки
@dp.message_handler(commands=['unsubscribe'])
async def unsubscribe(message):
    if get_len(message.from_user.id) == 'ru':
        if (str(message.text)).__len__() != 20:
            await message.answer('Запрос должен быть формата /unsubscribe GX-0000')
        gx = message.text[13:]
        if subscriber_exists(message.from_user.id, gx) and subscriber_status(message.from_user.id):
            """Если пользователь есть в базе данных и активен"""
            update_subscription(message.from_user.id, False)
            await message.answer('Вы успешно отписались от уведомлений!')
        else:
            """Если пользователя нет в базе данных или он неактивен"""
            await message.answer('Вы и так не подписаны!')

    elif get_len(message.from_user.id) == 'kgz':
        if (str(message.text)).__len__() != 20:
            await message.answer('Катуу форматта болуш керек /unsubscribe GX-0000')
        gx = message.text[13:]
        if subscriber_exists(message.from_user.id, gx) and subscriber_status(message.from_user.id):
            """Если пользователь есть в базе данных и активен"""
            update_subscription(message.from_user.id, False)
            await message.answer('Эскертмелерди ийгиликтүү жокко чыгардыңыз!')
        else:
            """Если пользователя нет в базе данных или он неактивен"""
            await message.answer('Сиз ансыз деле жазылбайсыз!')        


# Обрабатываем отправку новой таблицы
@dp.message_handler(content_types=['document'])
async def get_new_table(message):
    if get_len(message.from_user.id) == 'ru':
        file = await bot.get_file(message.document.file_id)
        await bot.download_file(file.file_path, 'data.xlsx')
        await message.answer('Вы успешно загрузили новую таблицу!')
    elif get_len(message.from_user.id) == 'kgz':
        file = await bot.get_file(message.document.file_id)
        await bot.download_file(file.file_path, 'data.xlsx')
        await message.answer('Жаңы таблицаны ийгиликтүү жүктөдүңүз!')


async def notifications(delay):
    while True:
        await asyncio.sleep(delay)
        data = data_to_notifications()
        str_to_send = ''
        for obj in get_subscriptions():
            for tup in data:
                gx = tup[0]
                if type(tup[0]) == int:
                    gx = str(tup[0])
                if obj[2] == gx or obj[2] == '0'+gx or obj[2] == '00'+gx or obj[2] == 'GX-00'+gx or obj[2] == 'ML-0'+gx:
                    str_to_send += f'Ваш товар с треком {tup[1]} пришел!\n'
            await bot.send_message(obj[1], str_to_send)

@dp.message_handler(commands=['test'])
async def tester(message: types.Message):
    if get_len(message.from_user.id)[0][0] == 'ru':
        id = str(message.text[6:])
        if id.__len__() == 7 and id[:3] == 'GX-':
            """Приводим все возможные виды GX кода"""
            ex1 = id
            ex2 = id[3:]
            ex3 = id[4:]
            ex4 = id[5:]

            await message.answer(text='Подождите минутку...')

            asyncio.create_task(dataXlsx(ex1, ex2, ex3, ex4))
            fail = open('to_data.wlsx', 'rb')
            await message.answer_document(message.chat.id, document=fail)
        elif id.__len__() == 4:
            """Приводим все возможные виды GX кода"""
            ex1 = 'GX-' + id
            ex2 = id
            ex3 = id[1:]
            ex4 = id[2:]


            await message.answer(text='Подождите минутку...')

            asyncio.create_task(dataXlsx(ex1, ex2, ex3, ex4))

            fail = open('to_data.wlsx', 'rb')
            await message.answer_document(message.chat.id, document=fail)
        else:
            await message.answer(text='Строго в фортмате GX-0000 или 0000!!!')

    # KGZ
    elif get_len(message.from_user.id)[0][0] == 'kgz':
        id = message.text[6:]
        if len(id) == 7 and id[:3] == 'GX-':
            """Приводим все возможные виды GX кода"""
            ex1 = id
            ex2 = id[3:]
            ex3 = id[4:]
            ex4 = id[5:]

            await message.answer(text='Бир мүнөт күтөө туруңуз...')

            asyncio.create_task(dataXlsx(ex1, ex2, ex3, ex4))

            fail = open('to_data.wlsx', 'rb')
            await message.answer_document(message.chat.id, document=fail)

        elif len(id) == 4:
            """Приводим все возможные виды GX кода"""
            ex1 = 'GX-' + id
            ex2 = id
            ex3 = id[1:]
            ex4 = id[2:]
            print(ex3, ex4)

            await message.answer(text='Бир мүнөт күтөө туруңуз...')

            asyncio.create_task(dataXlsx(ex1, ex2, ex3, ex4))

            fail = open('to_data.wlsx', 'rb')
            await message.answer_document(message.chat.id, document=fail)

        else:
            await message.answer(text='Катуу форматта болуш керек GX-0000 же 0000!!!')



if __name__ == '__main__':
    executor.start_polling(dp)
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    loop.create_task(notifications(43200))

