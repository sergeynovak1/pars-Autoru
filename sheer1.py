# Подключаем библиотеки
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


list = [0, 1, 2, 3, 4, 5]

results = service.spreadsheets().values().batchUpdate(spreadsheetId = spreadsheetId, body = {
    "valueInputOption": "USER_ENTERED",
    "data": [
        {"range": "Лист номер один!A2:F",
         "majorDimension": "ROWS", # Сначала заполнять строки, затем столбцы
         "values": [list
                   ]}
    ]
}).execute()



#размер колонок
'''results = service.spreadsheets().batchUpdate(spreadsheetId = spreadsheetId, body = {
  "requests": [

    # Задать ширину столбца A: 20 пикселей
    {
      "updateDimensionProperties": {
        "range": {
          "sheetId": sheetId,
          "dimension": "COLUMNS",  # Задаем ширину колонки
          "startIndex": 0, # Нумерация начинается с нуля
          "endIndex": 1 # Со столбца номер startIndex по endIndex - 1 (endIndex не входит!)
        },
        "properties": {
          "pixelSize": 20 # Ширина в пикселях
        },
        "fields": "pixelSize" # Указываем, что нужно использовать параметр pixelSize  
      }
    },

    # Задать ширину столбцов B и C: 150 пикселей
    {
      "updateDimensionProperties": {
        "range": {
          "sheetId": sheetId,
          "dimension": "COLUMNS",
          "startIndex": 1,
          "endIndex": 3
        },
        "properties": {
          "pixelSize": 150
        },
        "fields": "pixelSize"
      }
    },

    # Задать ширину столбца D: 200 пикселей
    {
      "updateDimensionProperties": {
        "range": {
          "sheetId": sheetId,
          "dimension": "COLUMNS",
          "startIndex": 3,
          "endIndex": 4
        },
        "properties": {
          "pixelSize": 200
        },
        "fields": "pixelSize"
      }
    }
  ]
}).execute()'''