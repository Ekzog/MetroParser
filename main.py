import asyncio
import time
import requests
import json
from bs4 import BeautifulSoup
import xlsxwriter
from aiohttp import ClientSession

def write_xls(results_):
    workbook = xlsxwriter.Workbook('metro.xlsx')
    bold = workbook.add_format({'bold': True})

    worksheet = workbook.add_worksheet()
    worksheet.write('A1', 'Магазин', bold)
    worksheet.write('B1', 'id товара', bold)
    worksheet.write('C1', 'наименование', bold)
    worksheet.write('D1', 'ссылка на товар', bold)
    worksheet.write('E1', 'регулярная цена', bold)
    worksheet.write('F1', 'промо цена', bold)
    worksheet.write('G1', 'бренд', bold)

    item = 2

    for result in results_:
        store = result["store"]
        result = result["products"]
        for product in result:
            worksheet.write(f'A{item}', store)
            worksheet.write(f'B{item}', product["id"])
            worksheet.write(f'C{item}', product["name"])
            worksheet.write(f'D{item}', 'https://online.metro-cc.ru' + product["url"])
            if product["stocks"][0]['prices']['discount'] != None:
                worksheet.write(f'E{item}', product["stocks"][0]['prices']['old_price'])
                worksheet.write(f'F{item}', product["stocks"][0]['prices']['price'])
            else:
                worksheet.write(f'E{item}', '-')
                worksheet.write(f'F{item}', product["stocks"][0]['prices']['price'])
            worksheet.write(f'G{item}', product["attributes"][0]['text'])
            item += 1

    workbook.close()

async def get_products(store_name_, store_id_, category_, sem):
    async with sem, ClientSession() as session:
        url = f'https://api.metro-cc.ru/products-api/graph'
        query = {
            "query": "query Query {"
                "category("
                "storeId: " + str(store_id_) +","
                "slug: \"" + category_ + "\","
                "inStock: true){"
                    "products(from: 0, size: 10000){\r\n"
                        "id\r\n"
                        "name\r\n"
                        "url\r\n"
                        "stocks {"
                            "prices {"
                                "discount\r\n"
                                "old_price\r\n"
                                "price\r\n"
                            "}"
                        "}"
                        "attributes {text}"
                    "}"
                "}"
            "}"
        }

        async with session.get(url=url, params=query) as response:
            products_json = await response.json()
            products_json["data"]["category"]["store"]= store_name_
            return products_json["data"]["category"]


async def stores_in_cities():
    url = "https://www.metro-cc.ru/sxa/search/results?p=1000&v=%7BBECE07BD-19B3-4E41-9C8F-E9D9EC85574F%7D&g="
    header = {
        'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3',
        'accept-encoding': 'gzip, deflate, br',
        'accept-language': 'ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7',
        'cache-control': 'no-cache',
        'dnt': '1',
        'pragma': 'no-cache',
        'sec-fetch-mode': 'navigate',
        'sec-fetch-site': 'none',
        'sec-fetch-user': '?1',
        'upgrade-insecure-requests': '1',
        'user-agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/76.0.3809.100 Safari/537.36'
    }
    session = requests.Session()
    session.headers = header
    response = session.get(url)
    html_response = json.loads(response.text)
    quantity_metro = len(html_response["Results"])
    full_html = ''

    for i in range(quantity_metro):
        full_html += html_response["Results"][i]['Html']

    return BeautifulSoup(full_html, "html.parser")


async def main(cities_, category_):
    soup = await stores_in_cities()
    stores_name = soup.find_all('span', class_='field-store-name')
    stores_id = soup.find_all('span', class_='field-store-id')

    tasks = []

    sem = asyncio.Semaphore(6)

    for i in range(len(stores_id)):
        for city in cities_:
            if city.lower() in stores_name[i].text.lower():
                print(stores_name[i].text, " ", stores_id[i].text)
                tasks.append(
                    asyncio.create_task(
                        get_products(
                            stores_name[i].text,
                            stores_id[i].text,
                            category_,
                            sem)
                    )
                )

    results = await asyncio.gather(*tasks)
    write_xls(results)


cities = [
    "Москва",
    "Красноярск",
    "Санкт-Петербург",
    "Томск",
    "Омск",
    "Краснодар",
    "Казань",
    "Воронеж"]  # список городов
category = "kofe-v-zernakh"  # категория товаров

start_time = time.time()  # старт работы

asyncio.run(main(cities, category))

end_time = time.time()  # конец работы

elapsed_time = end_time - start_time
print('Elapsed time: ', elapsed_time)  # время парсинга данных
