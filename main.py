import requests
import json
from bs4 import BeautifulSoup
import xlsxwriter


cities = ["Москва", "Санкт-Петербург"] #список городов
category = "kofe" # категория товаров

url = "https://www.metro-cc.ru/sxa/search/results?p=1000&v=%7BBECE07BD-19B3-4E41-9C8F-E9D9EC85574F%7D&g="
header = {
  'accept':'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3',
  'accept-encoding':'gzip, deflate, br',
  'accept-language':'ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7',
  'cache-control':'no-cache',
  'dnt': '1',
  'pragma': 'no-cache',
  'sec-fetch-mode': 'navigate',
  'sec-fetch-site': 'none',
  'sec-fetch-user': '?1',
  'upgrade-insecure-requests': '1',
  'user-agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/76.0.3809.100 Safari/537.36'}

session = requests.Session()
session.headers = header
response = session.get(url)

html_response = json.loads(response.text)
quantity_metro = len(html_response["Results"]) #кол-во всех магазинов сети Метро в России

full_html = ''
for i in range(quantity_metro):
  full_html += html_response["Results"][i]['Html']

soup = BeautifulSoup(full_html, "html.parser")
stores_name = soup.find_all('span', class_='field-store-name')
stores_id = soup.find_all('span', class_='field-store-id')


workbook = xlsxwriter.Workbook('metro.xlsx')
worksheet = workbook.add_worksheet()

bold = workbook.add_format({'bold': True})

worksheet.write('A1', 'Магазин', bold)
worksheet.write('B1', 'id товара', bold)
worksheet.write('C1', 'наименование', bold)
worksheet.write('D1', 'ссылка на товар', bold)
worksheet.write('E1', 'регулярная цена', bold)
worksheet.write('F1', 'промо цена', bold)
worksheet.write('G1', 'бренд', bold)

item = 2

for i in range(len(stores_id)):
  for city in cities:
    if city.lower() in stores_name[i].text.lower():
      print(stores_id[i].text," ", stores_name[i].text)
      query = {
        "query": "query Query {\r\n  category(storeId: "+stores_id[i].text+", slug: \""+category +"\", inStock: true) {\r\n    products(from: 0, size: 10000) {\r\n      id\r\n      name\r\n      url\r\n      stocks {\r\n        prices {\r\n          discount\r\n          old_price\r\n          price\r\n        }\r\n      }\r\n      attributes {\r\n        text\r\n      }\r\n    }\r\n  }\r\n}",
      }
      response = requests.get("https://api.metro-cc.ru/products-api/graph", params=query)
      products = json.loads(response.text)
      products = products["data"]["category"]["products"]

      for product in products:
        worksheet.write(f'A{item}', stores_name[i].text)
        worksheet.write(f'B{item}', product["id"])
        worksheet.write(f'C{item}', product["name"])
        worksheet.write(f'D{item}', 'https://online.metro-cc.ru' + product["url"])
        worksheet.write(f'E{item}', stores_name[i].text)
        worksheet.write(f'F{item}', stores_name[i].text)

        if product["stocks"][0]['prices']['discount'] != None:
          worksheet.write(f'E{item}', product["stocks"][0]['prices']['old_price'])
          worksheet.write(f'F{item}', product["stocks"][0]['prices']['price'])
        else:
          worksheet.write(f'E{item}', '-')
          worksheet.write(f'F{item}', product["stocks"][0]['prices']['price'])
        worksheet.write(f'G{item}', product["attributes"][0]['text'])
        item+=1
      #print(len(products)) #проверка количества

workbook.close()

