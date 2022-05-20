import argparse
import io
import logging
import os
import re
import sys
import time
from copy import deepcopy
from pathlib import Path
import textwrap

import requests
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from kinopoisk_unofficial.kinopoisk_api_client import KinopoiskApiClient
from kinopoisk_unofficial.request.films.film_request import FilmRequest
from kinopoisk_unofficial.request.staff.staff_request import StaffRequest
from PIL import Image
from tqdm import tqdm

# Configure logging
logging.basicConfig(level=logging.INFO,
                    format='[%(asctime)s]%(levelname)s:%(name)s:%(message)s',
                    datefmt='%d.%m.%Y %H:%M:%S')
log = logging.getLogger("Lib")


def is_api_ok(api):
    '''Проверка авторизации.'''
    try:
        api_client = KinopoiskApiClient(api)
        request = FilmRequest(328)
        api_client.films.send_film_request(request)
    except Exception:
        return False
    else:
        return True


def image_to_file(image):
    """Return `image` as PNG file-like object."""
    image_file = io.BytesIO()
    image.save(image_file, format="PNG")
    return image_file


def get_resource_path(relative_path):
    '''
    Определение пути для запуска из автономного exe файла.

    Pyinstaller cоздает временную папку, путь в _MEIPASS.
    '''
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def copy_table_after(table, paragraph):
    '''Копирование таблицы в указанный параграф.'''
    tbl, p = table._tbl, paragraph._p
    new_tbl = deepcopy(tbl)
    p.addnext(new_tbl)


def clone_first_table(document: Document, num):
    '''Клонирует первую таблицу в документе num раз.'''
    template = document.tables[0]
    paragraph = document.paragraphs[0]
    for i in range(num):
        copy_table_after(template, paragraph)
        paragraph = document.add_paragraph()


def find_kp_id(film_list, api):
    film_codes = []
    film_not_found = []
    for film in film_list:
        time.sleep(0.2)
        payload = {'keyword': film, 'page': 1}
        headers = {'X-API-KEY': api, 'Content-Type': 'application/json'}
        try:
            r = requests.get('https://kinopoiskapiunofficial.tech/api/v2.1/films/search-by-keyword',
                             headers=headers,
                             params=payload)
            if r.status_code == 200:
                resp_json = r.json()
                if resp_json['searchFilmsCountResult'] == 0:
                    log.info(f'{film} не найден')
                    film_not_found.append(film)
                    continue
                else:
                    id = resp_json['films'][0]['filmId']
                    year = resp_json['films'][0]['year']
                    if 'nameRu' in resp_json['films'][0]:
                        found_film = resp_json['films'][0]['nameRu']
                    else:
                        found_film = resp_json['films'][0]['nameEn']
                    log.info(f'Найден фильм: {found_film} ({year}), kinopoisk id: {id}')
                    film_codes.append(id)
            else:
                log.warning('Ошибка доступа к https://kinopoiskapiunofficial.tech')
                return
        except Exception as e:
            log.warning("Exeption:", str(e))
            log.info(f'{film} не найден (exeption)')
            film_not_found.append(film)
            continue
    result = []
    result.append(film_codes)
    result.append(film_not_found)
    return result


def get_film_info(film_code, api, shorten=False):
    '''
    Получение информации о фильме с помощью kinopoisk_api_client.

            Элементы списка:
                0 - название фильма на русском языке
                1 - год
                2 - рейтинг Кинопоиска
                3 - список стран
                4 - описание
                5 - ссылка на постер
                6 - имя файла без расширения
                7 - режиссер
             8:17 - 10 актеров
               18 - Постер размером 360x540 в формате PIL.Image.Image
    '''
    api_client = KinopoiskApiClient(api)
    request_staff = StaffRequest(film_code)
    response_staff = api_client.staff.send_staff_request(request_staff)
    staff_list = []
    if len(response_staff.items) >= 11:
        for i in range(0, 11):  # загружаем 11 персоналий (режиссер + 10 актеров)
            if response_staff.items[i].name_ru == '':
                staff_list.append(response_staff.items[i].name_en)
            else:
                staff_list.append(response_staff.items[i].name_ru)
    else:
        for i in range(0, len(response_staff.items)):
            if response_staff.items[i].name_ru == '':
                staff_list.append(response_staff.items[i].name_en)
            else:
                staff_list.append(response_staff.items[i].name_ru)
        for i in range(11 - len(response_staff.items)):
            staff_list.append("")
    request_film = FilmRequest(film_code)
    response_film = api_client.films.send_film_request(request_film)
    # с помощью регулярного выражения находим значение стран в кавычках ''
    countries = re.findall("'([^']*)'", str(response_film.film.countries))
    # имя файла
    if response_film.film.name_ru is not None:
        file_name = response_film.film.name_ru
        film_name = response_film.film.name_ru
    else:
        file_name = response_film.film.name_original
        film_name = response_film.film.name_original
    # очистка имени файла от запрещенных символов
    trtable = file_name.maketrans('', '', '\/:*?"<>')
    file_name = file_name.translate(trtable)

    # Сокращение описания фильма
    if shorten:
        description = response_film.film.description.replace("\n\n", " ")
        description = textwrap.shorten(description, 665, fix_sentence_endings=True,
                                       break_long_words=False, placeholder='...')
    else:
        description = response_film.film.description

    film_list = [
        film_name,
        response_film.film.year,
        response_film.film.rating_kinopoisk,
        countries,
        # response_film.film.description,
        description,
        response_film.film.poster_url,
        file_name
    ]
    result = film_list + staff_list
    # загрузка постера
    cover_url = response_film.film.poster_url
    cover = requests.get(cover_url, stream=True)
    if cover.status_code == 200:
        cover.raw.decode_content = True
        image = Image.open(cover.raw)
        width, height = image.size
        # обрезка до соотношения сторон 1x1.5
        if width > (height / 1.5):
            image = image.crop((((width - height / 1.5) / 2), 0, ((width - height / 1.5) / 2) + height / 1.5, height))
        image.thumbnail((360, 540))
        rgb_image = image.convert('RGB')  # Fix "OSError: cannot write mode RGBA as JPEG"
        result.append(rgb_image)
    else:
        result.append("")
    return result


def get_full_film_list(film_codes: list, api: str, shorten=False):
    """Загружает информацию о фильмах

    Args:
        film_codes (list): Список kinopoisk_id фильмов
        api (str): Kinopoisk API token
        shorten (boolean): Option to shorten movie descriptions
    Returns:
        list: Список с полной информацией о фильмах для записи в таблицу.
    """
    full_films_list = []
    for film_code in tqdm(film_codes, desc="Загрузка информации...   "):
        try:
            film_info = get_film_info(film_code, api, shorten)
            full_films_list.append(film_info)
        except Exception as e:
            print("Exeption:", str(e))
            # log.warning(f'{film_code} - ошибка')
        else:
            continue
    return full_films_list


def write_film_to_table(current_table, filminfo: list):
    '''Заполнение таблицы в файле docx.'''
    paragraph = current_table.cell(0, 1).paragraphs[0]  # название фильма + рейтинг
    run = paragraph.add_run(str(filminfo[0]) + ' - ' + 'Кинопоиск ' + str(filminfo[2]))
    run.font.name = 'Arial'
    run.font.size = Pt(11)
    run.font.bold = True

    paragraph = current_table.cell(1, 1).add_paragraph()  # год
    run = paragraph.add_run(str(filminfo[1]))
    run.font.name = 'Arial'
    run.font.size = Pt(10)

    paragraph = current_table.cell(1, 1).add_paragraph()  # страна
    run = paragraph.add_run(', '.join(filminfo[3]))
    run.font.name = 'Arial'
    run.font.size = Pt(10)

    paragraph = current_table.cell(1, 1).add_paragraph()  # режиссер
    run = paragraph.add_run('Режиссер: ' + filminfo[7])
    run.font.name = 'Arial'
    run.font.size = Pt(10)

    paragraph = current_table.cell(1, 1).add_paragraph()

    paragraph = current_table.cell(1, 1).add_paragraph()  # в главных ролях
    run = paragraph.add_run('В главных ролях: ')
    run.font.color.rgb = RGBColor(255, 102, 0)
    run.font.name = 'Arial'
    run.font.size = Pt(10)
    run = paragraph.add_run(', '.join(filminfo[8:18]))
    run.font.color.rgb = RGBColor(0, 0, 255)
    run.font.name = 'Arial'
    run.font.size = Pt(10)
    run.font.underline = True

    paragraph = current_table.cell(1, 1).add_paragraph()
    paragraph = current_table.cell(1, 1).add_paragraph()
    paragraph = current_table.cell(1, 1).add_paragraph()  # синопсис
    run = paragraph.add_run(filminfo[4])
    run.font.name = 'Arial'
    run.font.size = Pt(10)
    paragraph = current_table.cell(1, 1).add_paragraph()

    # запись постера в таблицу
    paragraph = current_table.cell(0, 0).paragraphs[1]
    run = paragraph.add_run()
    run.add_picture(image_to_file(filminfo[18]), width=Cm(7))


def write_all_films_to_docx(document, films: list, path: str):
    table_num = len(films)
    if table_num > 1:
        clone_first_table(document, table_num - 1)
    for i in tqdm(range(table_num), desc="Запись в таблицу...      "):
        current_table = document.tables[i]
        write_film_to_table(current_table, films[i])
    try:
        document.save(path)
        log.info(f'Файл "{path}" создан.')
    except PermissionError:
        log.warning(f"Ошибка! Нет доступа к файлу {path}. Список не сохранен.")
        raise Exception


def file_to_list(file: str):
    if os.path.isfile(file):
        with open(file, 'r', encoding="utf-8") as f:
            list = [x.rstrip() for x in f]
        return list
    else:
        print(f'Файл {file} не найден.')
        raise FileNotFoundError


def main():
    from config import KINOPOISK_API_TOKEN
    parser = argparse.ArgumentParser(prog='Kinolist Library')
    parser.add_argument("-f", "--file", nargs=1, help="list of films in .txt format")
    parser.add_argument("-m", "--movie", nargs="+", help="list of films")
    parser.add_argument("-o", "--output", nargs=1, help="output file name (list.docx by default)")
    parser.add_argument("-s", "--shorten", action='store_true', help="optionally shorten movie descriptions")
    args = parser.parse_args()

    if args.output:
        output = args.output[0]
        output_dir, output_file_name = os.path.split(output)
        Path(output_dir).mkdir(parents=True, exist_ok=True)
        _, ext = os.path.splitext(output_file_name)
        if ext != ".docx":
            print("Output file must have .docx extension.")
            return
    else:
        output = "list.docx"

    if args.file:
        list = file_to_list((args.file[0]))
        print(f"Запрос из {args.file[0]}: ", ", ".join(list))
        file_path = get_resource_path('template.docx')
        doc = Document(file_path)
        kp_codes = find_kp_id(list, KINOPOISK_API_TOKEN)
        if len(kp_codes[1]) != 0:
            print("Не найдено:", ", ".join(kp_codes[1]))
        full_list = get_full_film_list(kp_codes[0], KINOPOISK_API_TOKEN, args.shorten)
        write_all_films_to_docx(doc, full_list, output)

    elif args.movie:
        film = args.movie
        kp_codes = find_kp_id(film, KINOPOISK_API_TOKEN)
        if len(kp_codes[1]) != 0:
            print("Не найдено:", ", ".join(kp_codes[1]))
        full_list = get_full_film_list(kp_codes[0], KINOPOISK_API_TOKEN, args.shorten)
        template_path = get_resource_path('template.docx')
        doc = Document(template_path)
        write_all_films_to_docx(doc, full_list, output)


if __name__ == "__main__":
    main()
