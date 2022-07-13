import requests
from lxml import html

import httplib2
import apiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials

CREDENTIALS_FILE = 'mypython-355606-9db3ff512811.json'  # Имя файла с закрытым ключом, вы должны подставить свое

# Читаем ключи из файла
credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])

httpAuth = credentials.authorize(httplib2.Http()) # Авторизуемся в системе
service = apiclient.discovery.build('sheets', 'v4', http = httpAuth) # Выбираем работу с таблицами и 4 версию API

spreadsheet = service.spreadsheets().create(body = {
    'properties': {'title': 'Первый тестовый документ', 'locale': 'ru_RU'},
    'sheets': [{'properties': {'sheetType': 'GRID',
                               'sheetId': 0,
                               'title': 'Лист номер один',
                               'gridProperties': {'rowCount': 100, 'columnCount': 15}}}]
}).execute()
spreadsheetId = spreadsheet['spreadsheetId'] # сохраняем идентификатор файла
print('https://docs.google.com/spreadsheets/d/' + spreadsheetId)

driveService = apiclient.discovery.build('drive', 'v3', http = httpAuth) # Выбираем работу с Google Drive и 3 версию API
access = driveService.permissions().create(
    fileId = spreadsheetId,
    body = {'type': 'user', 'role': 'writer', 'emailAddress': 'novak.serzj@gmail.com'},  # Открываем доступ на редактирование
    fields = 'id'
).execute()

results = service.spreadsheets().batchUpdate(spreadsheetId = spreadsheetId, body = {
  "requests": [

    # Задать ширину столбца A: 20 пикселей
    {
      "updateDimensionProperties": {
        "range": {
          "dimension": "COLUMNS",  # Задаем ширину колонки
          "startIndex": 1, # Нумерация начинается с нуля
          "endIndex": 2 # Со столбца номер startIndex по endIndex - 1 (endIndex не входит!)
        },
        "properties": {
          "pixelSize": 320 # Ширина в пикселях
        },
        "fields": "pixelSize" # Указываем, что нужно использовать параметр pixelSize
      }
    },

# Задать ширину столбца A: 20 пикселей
    {
      "updateDimensionProperties": {
        "range": {
          "dimension": "COLUMNS",  # Задаем ширину колонки
          "startIndex": 6, # Нумерация начинается с нуля
          "endIndex": 8 # Со столбца номер startIndex по endIndex - 1 (endIndex не входит!)
        },
        "properties": {
          "pixelSize": 150 # Ширина в пикселях
        },
        "fields": "pixelSize" # Указываем, что нужно использовать параметр pixelSize
      }
    },
    {
      "updateDimensionProperties": {
        "range": {
          "dimension": "COLUMNS",  # Задаем ширину колонки
          "startIndex": 10, # Нумерация начинается с нуля
          "endIndex": 11 # Со столбца номер startIndex по endIndex - 1 (endIndex не входит!)
        },
        "properties": {
          "pixelSize": 130 # Ширина в пикселях
        },
        "fields": "pixelSize" # Указываем, что нужно использовать параметр pixelSize
      }
    },
  ]
}).execute()
results = service.spreadsheets().values().batchUpdate(spreadsheetId = spreadsheetId, body = {
    "valueInputOption": "USER_ENTERED", # Данные воспринимаются, как вводимые пользователем (считается значение формул)
    "data": [
        {"range": "Лист номер один!A1:L1",
         "majorDimension": "ROWS",
         "values": [["Марка", "Модель", "Год", "Цена", "Цвет", "Пробег", "Кузов", "Двигатель", "КПП", "Привод", "Город"], ]}
    ]
}).execute()

URL = 'https://auto.ru/cars/all/?year_from=2010&year_to=2021&page='
r = requests.get(URL)
tree = html.fromstring(r.content)

num_of_page = 99
i = 1
k = 2

while i < 100:
    try:
        url = URL + str(i)
        r = requests.get(url)
        tree = html.fromstring(r.content)

        titles = tree.xpath('//a[@class="Link ListingItemTitle__link"]//text()')
        marka = [titles[i].split()[0] for i in range(len(titles))]
        prices = [price.replace(u'\xa0', ' ').replace(u' ₽', '') for price in
                  tree.xpath('//div[@class="ListingItemPrice__content"]//text()')]
        tech_param_xpath = [param.replace(u'\u2009/\u2009', ' ').replace(u'\xa0', ' ') for param in tree.xpath(
            '//div[@class="ListingItemTechSummaryDesktop ListingItem__techSummary"]//text()')]
        colors = tech_param_xpath[4::5]
        bodies = tech_param_xpath[2::5]
        drive_by = tech_param_xpath[3::5]
        engines = tech_param_xpath[::5]
        engine_by = tech_param_xpath[1::5]
        years = tree.xpath('//div[@class="ListingItem__year"]//text()')
        km_age = [km.replace(u'\xa0', ' ') for km in tree.xpath('//div[@class="ListingItem__kmAge"]//text()')]
        towns = tree.xpath('//span[@class="MetroListPlace__regionName MetroListPlace_nbsp"]//text()')
        for index in range(len(titles)):
            infovalue = [marka[index], titles[index], years[index], prices[index], colors[index], km_age[index], bodies[index], engines[index], engine_by[index], drive_by[index], towns[index]]
            results = service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheetId, body={
                "valueInputOption": "RAW",
                "data": [
                    {"range": f"A{k}:L5000",
                     "majorDimension": "ROWS",  # Сначала заполнять строки, затем столбцы
                     "values": [infovalue
                                ]}
                ]
            }).execute()
            k += 1
        i += 1
        print(i)

    except:
        continue