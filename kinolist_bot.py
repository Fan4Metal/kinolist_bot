import logging
import shutil
import os
import argparse_ru
import argparse
from random import choice

from aiogram import Bot, Dispatcher, executor, types
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from aiogram.dispatcher.filters.state import State, StatesGroup
import aiogram.utils.markdown as fmt
from docx2pdf import convert
from kinolist_lib import *
import config

VER = '0.4.1'
TELEGRAM_API_TOKEN = config.TELEGRAM_API_TOKEN
KINOPOISK_API_TOKEN = config.KINOPOISK_API_TOKEN


parser = argparse.ArgumentParser(prog='Kinolist_Bot',
                                description='Телеграм бот для быстрого создания списков фильмов (@kinolist_one_bot)')
parser.add_argument("-ver", "--version", action="version", version=f"%(prog)s {VER}", help="выводит версию программы и завершает работу")
parser.add_argument("-l", "--log", action='store_true', help="включает запись лога в файл kinolist_bot.log")
parser.add_argument("--libre", action='store_true', help="конвертация docx в pdf с помощью Libre Office")
args = parser.parse_args()

# Configure logging
if args.log:
    logging.basicConfig(filename='kinolist_bot.log', level=logging.INFO,
                        format='[%(asctime)s]%(levelname)s:%(name)s:%(message)s',
                        datefmt='%d.%m.%Y %H:%M:%S')
else:
    logging.basicConfig(level=logging.INFO,
                        format='[%(asctime)s]%(levelname)s:%(name)s:%(message)s',
                        datefmt='%d.%m.%Y %H:%M:%S')
log = logging.getLogger("Bot")
log.info(f"Kinolist_Bot ver. {VER}, Kinolist_Lib ver. {LIB_VER}")

# Initialize bot and dispatcher
storage = MemoryStorage()
bot = Bot(token=TELEGRAM_API_TOKEN)
dp = Dispatcher(bot, storage=storage)


# States
class DocFormat(StatesGroup):
    pdf = State()
    docx = State()
    info = State()


@dp.message_handler(state='*', commands=['start', 'help'])
async def send_welcome(message: types.Message):
    """
    This handler will be called when user sends `/start` or `/help` command
    """
    log.info(f"Начало работы (chat_id: {message.chat.id})")
    if os.path.exists("./" + str(message.chat.id)):
        shutil.rmtree("./" + str(message.chat.id))
        log.info(f"Каталог очищен")
    await DocFormat.pdf.set()
    await message.reply("Привет, я Кinolist Bot!\nОтправьте мне список фильмов, и я пришлю его в формате pdf.")


@dp.message_handler(state='*', commands=['word', 'docx'])
async def send_welcome(message: types.Message):
    log.info(f"Переключение на отправку списков в формате docx (chat_id: {message.chat.id})")
    await DocFormat.docx.set()
    await message.reply("Ок, отправьте мне список фильмов, и я пришлю его в формате *docx*\.", parse_mode="MarkdownV2")


@dp.message_handler(state='*', commands=['pdf'])
async def send_welcome(message: types.Message):
    log.info(f"Переключение на отправку списков в формате pdf (chat_id: {message.chat.id})")
    await DocFormat.pdf.set()
    await message.reply("Ок, отправьте мне список фильмов, и я пришлю его в формате *pdf*\.", parse_mode="MarkdownV2")


@dp.message_handler(state='*', commands=['info'])
async def send_welcome(message: types.Message):
    log.info(f"Переключение на отправку списков в чат (chat_id: {message.chat.id})")
    await DocFormat.info.set()
    await message.reply("Ок, отправьте мне список фильмов, и я пришлю его описание в чат\.", parse_mode="MarkdownV2")


@dp.message_handler(state='*', commands=['lisa', 'Lisa'])
async def send_heart(message: types.Message):
    stickers = ["CAACAgIAAxkBAAEEjSZiZXLQqPDFY70qC0m9PPH2AAEJjfgAAjIAA-Sgzgd7_cFVbY2YfiQE",
                "CAACAgIAAxkBAAEE56JimeqS59Ey3lewiSQVUELCyRg36QACQQADspiaDm4PLeCw1KxAJAQ",
                "CAACAgIAAxkBAAEE56NimeqST5qTkoyZ4FH3v7RPWAxFkAACdRIAAgguuUgySuZE4jBZdiQE",
                "CAACAgIAAxkBAAEE56RimeqSFOSFCQfwqqB7syw5Pka9GwACDwkAAhhC7gjWQ00JSrFY0yQE",
                "CAACAgIAAxkBAAEE56VimeqSP8JYJPX5i3tzlO9URLECoAACqgAD5KDOB-HQd7qptDvWJAQ",
                "CAACAgIAAxkBAAEE56ZimeqSBIVcEBiVOJUq1z8lrPOehgACwAgAAhhC7ghAwahsXqNd9SQE",
                "CAACAgIAAxkBAAEE565imevqhC3H8qyDhAlKm_oGB7gcmAACSwMAArVx2gZu3ktViH-zcCQE"
    ]
    await message.reply_sticker(choice(stickers))
    log.info("Отправлен стикер")


@dp.message_handler(state=DocFormat.pdf)
async def reply(message: types.Message):
    if not is_api_ok(KINOPOISK_API_TOKEN):
        log.warning("API error.")
        await message.reply("Ой, что-то сломалось!((\n(API error)")
        return

    chat_id = str(message.chat.id)
    log.info(f"Начало создания списка для chat_id: {chat_id}")
    if os.path.isdir(chat_id):
        log.info(f"Папка {chat_id} обнаружена")
        await message.reply("Подождите, я все еще работаю!")
        return

    film_list = message.text.split('\n')
    film_list = list(filter(None, film_list))
    log.info("Запрос: " + ", ".join(film_list))

    kp_id = find_kp_id(film_list, KINOPOISK_API_TOKEN)
    film_codes = kp_id[0]
    film_not_found = kp_id[1]

    if len(film_not_found) > 0:
        log.info(f'Не найдено: {", ".join(film_not_found)}')
    if len(film_codes) == 0:
        await message.reply("Ой, ничего не найдено!")
        return

    full_films_list = get_full_film_list(film_codes, KINOPOISK_API_TOKEN)
    if len(full_films_list) < 1:
        await message.reply("Ни один фильм не найден!")
        return
    if args.libre:
        template_path = get_resource_path('template_libre.docx')
    else:
        template_path = get_resource_path('template.docx')
    try:
        doc = Document(template_path)
    except Exception:
        log.warning('Не найден шаблон "template.docx". Список не создан.')
        await message.reply("Ой, что-то сломалось!((")
        return

    if not os.path.isdir(chat_id):
        os.mkdir(chat_id)

    path_docx = chat_id + "/list.docx"
    try:
        write_all_films_to_docx(doc, full_films_list, path_docx)
    except:
        log.warning('Ошибка при записи файла docx')
        await message.reply("Ой, что-то сломалось!((")
        return
    path_pdf = chat_id + "/list.pdf"
    if args.libre:
        log.info("Конвертация docx в pdf через Libre Office")
        if docx_to_pdf_libre(path_docx) != 0:
            log.warning("Ошибка конвертации в pdf через Libre Office")
            await message.reply("Ой, что-то сломалось!((")
            return
    else:
        log.info("Конвертация docx в pdf через Microsoft Word")
        convert(path_docx, path_pdf)
    log.info(f'Файл "{path_pdf}" создан.')
    with open(path_pdf, 'rb') as pdf:
        if len(film_not_found) > 0:
            text = "Список готов!\n" + "Правда, вот эти фильмы не смог найти:\n" + "\n".join(film_not_found)
            await message.reply_document(pdf, caption=text)
        else:
            await message.reply_document(pdf, caption='Список готов!')
    log.info(f'Список отправлен в чат: {chat_id}')
    shutil.rmtree(chat_id)
    return


@dp.message_handler(state=DocFormat.docx)
async def reply(message: types.Message):
    if not is_api_ok(KINOPOISK_API_TOKEN):
        log.warning("API error.")
        await message.reply("Ой, что-то сломалось!((\n(API error)")
        return

    chat_id = str(message.chat.id)
    log.info(f"Начало создания списка для chat_id: {chat_id}")
    if os.path.isdir(chat_id):
        log.info(f"Папка {chat_id} обнаружена")
        await message.reply("Подождите, я все еще работаю!")
        return

    film_list = message.text.split('\n')
    film_list = list(filter(None, film_list))
    log.info("Запрос: " + ", ".join(film_list))

    kp_id = find_kp_id(film_list, KINOPOISK_API_TOKEN)
    film_codes = kp_id[0]
    film_not_found = kp_id[1]

    if len(film_not_found) > 0:
        log.info(f'Не найдено: {", ".join(film_not_found)}')
    if len(film_codes) == 0:
        await message.reply("Ой, ничего не найдено!")
        return

    full_films_list = get_full_film_list(film_codes, KINOPOISK_API_TOKEN)
    if len(full_films_list) < 1:
        await message.reply("Ни один фильм не найден!")
        return

    template_path = get_resource_path('template.docx')
    try:
        doc = Document(template_path)
    except Exception:
        log.warning('Не найден шаблон "template.docx". Список не создан.')
        await message.reply("Ой, что-то сломалось!((")
        return

    if not os.path.isdir(chat_id):
        os.mkdir(chat_id)

    path_docx = chat_id + "/list.docx"
    try:
        write_all_films_to_docx(doc, full_films_list, path_docx)
    except:
        log.warning('Ошибка при записи файла docx')
        await message.reply("Ой, что-то сломалось!((")
        return

    with open(path_docx, 'rb') as docx:
        if len(film_not_found) > 0:
            text = "Список готов!\n" + "Правда, вот эти фильмы не смог найти:\n" + "\n".join(film_not_found)
            await message.reply_document(docx, caption=text)
        else:
            await message.reply_document(docx, caption='Список готов!')
    log.info(f'Список отправлен в чат: {chat_id}')
    shutil.rmtree(chat_id)
    return


@dp.message_handler(state=DocFormat.info)
async def reply(message: types.Message):
    if not is_api_ok(KINOPOISK_API_TOKEN):
        log.warning("API error.")
        await message.reply("Ой, что-то сломалось!((\n(API error)")
        return

    chat_id = str(message.chat.id)
    log.info(f"Начало создания списка для chat_id: {chat_id}")

    film_list = message.text.split('\n')
    film_list = list(filter(None, film_list))
    log.info("Запрос: " + ", ".join(film_list))

    kp_id = find_kp_id(film_list, KINOPOISK_API_TOKEN)
    film_codes = kp_id[0]
    film_not_found = kp_id[1]

    if len(film_not_found) > 0:
        log.info(f'Не найдено: {", ".join(film_not_found)}')
    if len(film_codes) == 0:
        await message.reply("Ой, ничего не найдено!")
        return

    full_films_list = get_full_film_list(film_codes, KINOPOISK_API_TOKEN)
    if len(full_films_list) < 1:
        await message.reply("Ни один фильм не найден!")
        return

    for film in full_films_list:
        text = fmt.text(fmt.text(fmt.bold(f"{film[0]} ({film[1]}) - Кинопоиск {film[2]}")),
                        fmt.text(", ".join(film[3])),
                        fmt.text("Режиссер:" if len(film[7]) ==1 else "Режиссеры:", text_to_markdown(", ".join(film[7]))),
                        fmt.text(""),
                        fmt.text("В главных ролях:", fmt.underline(", ".join(film[8]))),
                        fmt.text(""),
                        fmt.text(text_to_markdown(film[4])),
                        sep="\n"
        )
        await message.reply_photo(film[6], caption=text, parse_mode="MarkdownV2")
    log.info(f'Информация о фильмах отправлена в чат: {chat_id}')
    return


if __name__ == '__main__':
    executor.start_polling(dp, skip_updates=True)
