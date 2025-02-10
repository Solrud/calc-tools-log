import os
from datetime import datetime
TODAY_DATE = datetime.now().strftime('%d.%m.%Y')

TOOLS_LOG_PATH = 'O:/BUFFER/ilin/Обработчик ToolsLog и Svod/ToolsLog 2024-2025.xlsx'
SVOD_EXCEL_PATH = 'O:/BUFFER/ilin/Обработчик ToolsLog и Svod/SvodExcel.xlsx'
# OUTPUT_PATH = os.path.expanduser("~/Desktop")
OUTPUT_PATH = ''
OUTPUT_FILE_NAME = f'Обработка ToolsLog {TODAY_DATE}'
OUTPUT_FILE_TYPE = '.xlsx'

LOGS_PATH = 'O:/BUFFER/ilin/всякое/log handler tools svod/logs.txt'

APP_VERSION = '1.2'
APP_VERSION_DATE = '10.02.2025'

HEADER_NAME_CLOUMNS_ALL = [
    'UID',
    'Раздел',
    'Тип_инструмента',
    'Обозначение',
    'Наименование',
    'Производитель',
    'Доп._параметры',
    'Остаток ИРК 20_1',
    'Остаток ИРК_21_1',
    'Остаток ИРК_22_1',
    'Остаток ИРК_23_1',
    'Остаток ИРК_20_2',
    'Остаток ИРК_23_2',
    'Выдано раз',
    'Выдано количество суммарно',
    'Дата_посл.выдачи',
    'Приход раз',
    'Приход количество суммарно',
    'Дата_посл.прихода',
    'Ячейка',
    'Стелаж'
]
HEADERS_NAME_COLUMNS_TOOLS_LOG = [
    'Раздел',
    'Тип_инструмента',
    'Выдано раз',
    'Выдано количество суммарно',
    'Дата_посл.выдачи',
    'Приход раз',
    'Приход количество суммарно',
    'Дата_посл.прихода'
]