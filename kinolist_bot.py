import logging
import shutil
import os

from aiogram import Bot, Dispatcher, executor, types
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from aiogram.dispatcher.filters.state import State, StatesGroup
from docx2pdf import convert
from kinolist_lib import *
import config

VER = '0.3.1'
TELEGRAM_API_TOKEN = config.TELEGRAM_API_TOKEN
KINOPOISK_API_TOKEN = config.KINOPOISK_API_TOKEN

# Configure logging
logging.basicConfig(level=logging.INFO,
                    format='[%(asctime)s]%(levelname)s:%(name)s:%(message)s',
                    datefmt='%d.%m.%Y %H:%M:%S')
log = logging.getLogger("Bot")


# Initialize bot and dispatcher
storage = MemoryStorage()
bot = Bot(token=TELEGRAM_API_TOKEN)
dp = Dispatcher(bot, storage=storage)


# States
class DocFormat(StatesGroup):
    pdf = State()
    docx = State()


@dp.message_handler(state='*', commands=['start', 'help'])
async def send_welcome(message: types.Message):
    """
    This handler will be called when user sends `/start` or `/help` command
    """
    log.info(f"Start (chat_id: {message.chat.id})")
    # Set state
    await DocFormat.pdf.set()
    await message.reply("Привет, я Кinolist Bot!\nОтправьте мне список фильмов, и я пришлю его в формате pdf.")


@dp.message_handler(state='*', commands=['word', 'docx'])
async def send_welcome(message: types.Message):
    log.info(f"Start request for docx document (chat_id: {message.chat.id})")
    await DocFormat.docx.set()
    await message.reply("Ок, отправьте мне список фильмов, и я пришлю его в формате *docx*\.", parse_mode="MarkdownV2")


@dp.message_handler(state='*', commands=['pdf'])
async def send_welcome(message: types.Message):
    log.info(f"Start request for docx document (chat_id: {message.chat.id})")
    await DocFormat.pdf.set()
    await message.reply("Ок, отправьте мне список фильмов, и я пришлю его в формате *pdf*\.", parse_mode="MarkdownV2")


@dp.message_handler(state='*', commands=['lisa', 'Lisa'])
async def send_heart(message: types.Message):
    await message.reply_sticker("CAACAgIAAxkBAAEEjSZiZXLQqPDFY70qC0m9PPH2AAEJjfgAAjIAA-Sgzgd7_cFVbY2YfiQE")
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


if __name__ == '__main__':
    executor.start_polling(dp, skip_updates=True)
