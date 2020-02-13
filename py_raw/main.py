import yaml
import datetime
import requests
import pandas as pd
import logging.config

from bs4 import BeautifulSoup as bs


# чтение ymal-файла с настройками логирования, создание логгера
with open('config.yaml', 'r') as f:
    config = yaml.safe_load(f.read())
    logging.config.dictConfig(config)
logger = logging.getLogger(__name__)


def get_urls_from_excel():
    """
    Достаем из экселевского файла адрес, который будем парсить
    """
    instr_urls_list = []
    likar_urls_list = []
    try:
        name = "./../urls/urls.xlsx"
        urls_list = pd.read_excel(name)
        for i in range(len(urls_list['instr_urls'])):
            instr_urls_list.append(urls_list['instr_urls'][i])
        for i in range(len(urls_list['likar_urls'])):
            likar_urls_list.append(urls_list['likar_urls'][i])

        return instr_urls_list, likar_urls_list
    except FileExistsError:
        logger.error('some shit')
    except FileNotFoundError:
        logger.error('нет файла')


def price_cutter(item):
    '''убираем лишние пробелы и знак рубля из входящего текста, кроме запятых'''
    price = ''
    for i in item:
        if i.isdigit() == True:# or (i == ','):
            price += i
        if i == ',':
            price += '.'
            return float(price)
    return int(price)


def parsing_instrument(instr_url_list):
    """
    Парсинг инструмента. На входе урл, на выходе словарь, где название - ключ, значения цена и ссылка.
    При это сначаала надо проверить, если ли дополнительные сраницы в разделе и если есть, добавить их к списку на парсинг.
    """
    headers = {
        'accpet': '*/*',
        'user-agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/75.0.3770.142 Safari/537.36'
    }
    instr_items_dict = {}
    session = requests.Session()
    extra_links = ""
    instr_url_list_full = instr_url_list[::]
    for url in instr_url_list:
        request = session.get(url)
        if request.status_code == 200:
            print('подключился к ', url)
            big_html = requests.get(url)  # получаем доступ к урлу
            big_soup = bs(big_html.text, 'html.parser')  # парсим его
            if big_soup.select('#catalog-products__show-more'):  # если кнопка show-more есть на странице, грабим сожержимое
                # TODO прописать что делать, если кнопки show-more нет на странице
                print('нашел продолжение страницы')
                extra_links = extra_links + (big_soup.select('#catalog-products__show-more')[0].attrs['data-urls'])
            for i in extra_links.split('"'):
                if len(i) > 5:
                    instr_url_list_full.append("https://kirov.instrument.ms" + i)  # вот он чистый полный список нижних ссылок
        else:
            print('some trouble with ', url)

    for url in instr_url_list_full:  # последовательно проходим по списку чистых урлов
        request = session.get(url)  # коннектимся
        if request.status_code == 200:  # если удачно:
            raw_html = requests.get(url)
            instr_soup = bs(raw_html.text, 'html.parser')
            names_list_raw = instr_soup.select('.product-card__name')
            prices_list_raw = instr_soup.select('.product-card__price-value')
            links_list_raw = instr_soup.find_all('a', itemprop="name")

            instr_names_list_clear = []
            instr_prices_list_clear = []
            instr_links_list_clear = []

            for i in names_list_raw:
                instr_names_list_clear.append(i.getText().strip())

            for i in prices_list_raw:
                instr_prices_list_clear.append(price_cutter(i.getText().strip()))

            for i in links_list_raw:
                instr_links_list_clear.append("https://kirov.instrument.ms" + i.attrs['href'])

            for i in range(len(instr_names_list_clear)):
                instr_items_dict[instr_names_list_clear[i]] = []
                instr_items_dict[instr_names_list_clear[i]].append(instr_prices_list_clear[i])
                instr_items_dict[instr_names_list_clear[i]].append(instr_links_list_clear[i])
        else:
            print('some trouble with ', url)

    instr_items_df = pd.DataFrame.from_dict(instr_items_dict, orient='index')
    instr_items_df.reset_index(drop=False, inplace=True)
    writer = pd.ExcelWriter('./../xlsx/instrument.xls', engine='xlsxwriter')
    instr_items_df.to_excel(writer, sheet_name='main', index=False)

# DECOR
    ########  ########  ######   #######  ########
    ##     ## ##       ##    ## ##     ## ##     ##
    ##     ## ##       ##       ##     ## ##     ##
    ##     ## ######   ##       ##     ## ########
    ##     ## ##       ##       ##     ## ##   ##
    ##     ## ##       ##    ## ##     ## ##    ##
    ########  ########  ######   #######  ##     ##

    workbook = writer.book

    cell_format = workbook.add_format({
        'bold': True,
        'font_color': 'black',
        'align': 'center',
        'valign': 'center',
        'bg_color': '#ecf0f1'
    })

    for sheet in ['main']:
        worksheet = writer.sheets[sheet]
        worksheet.set_column('A:A', 60)
        worksheet.set_column('B:B', 10)
        worksheet.set_column('C:F', 20)
        worksheet.write('A1', 'название', cell_format)
        worksheet.write('B1', 'цена', cell_format)
        worksheet.write('C1', 'ссылка', cell_format)

        worksheet.freeze_panes(1, 0)

    writer.save()
    writer.close()


def parsing_likar(likar_url_list):
    """
    Получает список
    Заходит на ссылку в этом списке
    Находит количество страниц которые есть в этом разделе
    Добавляет в список страниц эти номер добавочных страниц
    Заходит на каждуй из этих страниц
    Собирает массив имен, массив цен, массив ссылок
    Пушит их в словарь, словарь в эксель
    """
    headers = {
        'accpet': '*/*',
        'user-agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/75.0.3770.142 Safari/537.36'
    }
    likar_items_dict = {}
    session = requests.Session()
    likar_url_list_full = likar_url_list[::]
    for url in likar_url_list:
        request = session.get(url)
        if request.status_code == 200:
            raw_html = requests.get(url)  # получаем доступ к урлу
            likar_soup = bs(raw_html.text, 'html.parser')  # парсим его
            if likar_soup.select('.nums'):  # если див с номерами дополнительных страниц есть на странице
                number_of_extra_pages = int(likar_soup.select('.nums a')[-1].text)
                for i in range(2, number_of_extra_pages+1):
                    likar_url_list_full.append(url + '?PAGEN_1=' + str(i))
        else:
            logger.error('some shit with url', url)

    likar_names_list = []
    likar_prices_list = []
    likar_links_list = []
    for url in likar_url_list_full:
        request = session.get(url)
        if request.status_code == 200:
            raw_html = requests.get(url)
            likar_soup = bs(raw_html.text, 'html.parser')
            likar_names_list_raw = likar_soup.select('.item-title span')
            likar_links_list_raw = likar_soup.select('.item-title a')

        for i in likar_names_list_raw:
            likar_names_list.append(i.getText().strip())
        print('numbers of likar_names = ', len(likar_names_list))

        for i in likar_links_list_raw:
            likar_links_list.append('https://instrument-orugie.ru' + i.attrs['href'])
        print('numbers of likar_links = ', len(likar_links_list))

        for i in range(len(likar_names_list)):
            likar_items_dict[likar_names_list[i]] = []
            likar_items_dict[likar_names_list[i]].append(likar_links_list[i])

    for key, value in likar_items_dict.items():
        print(key, value)

    likar_items_df = pd.DataFrame.from_dict(likar_items_dict, orient='index')
    likar_items_df.reset_index(drop=False, inplace=True)
    writer = pd.ExcelWriter('./../xlsx/likar.xls', engine='xlsxwriter')
    likar_items_df.to_excel(writer, sheet_name='main', index=False)

# DECOR
    ########  ########  ######   #######  ########
    ##     ## ##       ##    ## ##     ## ##     ##
    ##     ## ##       ##       ##     ## ##     ##
    ##     ## ######   ##       ##     ## ########
    ##     ## ##       ##       ##     ## ##   ##
    ##     ## ##       ##    ## ##     ## ##    ##
    ########  ########  ######   #######  ##     ##

    workbook = writer.book

    cell_format = workbook.add_format({
        'bold': True,
        'font_color': 'black',
        'align': 'center',
        'valign': 'center',
        'bg_color': '#ecf0f1'
    })

    for sheet in ['main']:
        worksheet = writer.sheets[sheet]
        worksheet.set_column('A:A', 60)
        worksheet.set_column('B:B', 20)
        worksheet.write('A1', 'название', cell_format)
        worksheet.write('B1', 'ссылка', cell_format)

        worksheet.freeze_panes(1, 0)

    writer.save()
    writer.close()



if __name__ == "__main__":
    instr_url_list, likar_urls_list = get_urls_from_excel()
    # parsing_instrument(instr_url_list)
    parsing_likar(likar_urls_list)
