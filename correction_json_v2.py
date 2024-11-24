import pandas as pd
import json
import datetime

# Константы
operator_name = "Кассиров Кассир Кассирович"  # Указываем имя кассира
operator_inn = "123456789012"   # Указываем ИНН кассира

# Чтение данных из Excel
main_file = 'Чеки_Касса_№1.xlsx'  # Основной файл с информацией о чеках
items_file = 'Товары_Касса_№1.xlsx'  # Файл с товарами
df_main = pd.read_excel(main_file)
df_items = pd.read_excel(items_file)


# Функция для преобразования строки с датой в нужный формат
def format_date(date):
    return date.strftime('%Y.%m.%d')


# Создание списка для хранения всех запросов
correction_requests = []

# Обработка каждой строки в основной таблице
for index, row in df_main.iterrows():
    # Список товаров для текущего ФПД
    items_list = df_items[df_items['ФПД'] == row['ФПД']]

    # Проверка соответствия количества товарных позиций
    if row['Кол-во товарных позиций'] != len(items_list):
        print(f"Внимание: количество товарных позиций не совпадает для ФПД {row['ФПД']}")
        continue

    # Формирование списка товаров
    items_data = []
    for item in items_list.itertuples():
        items_data.append({
            "type": "position",
            "name": item.Товар,
            "price": item.Цена,
            "quantity": item.Кол_во,
            "amount": item.Сумма,
            "infoDiscountAmount": 0.0,
            "department": 1,
            "measurementUnit": 0, # единица измерения - штуки
            "paymentMethod": "fullPayment",
            "paymentObject": "commodity",
            "tax": {"type": "none"} # Тут указываем налог (none - это БЕЗ НДС)
        })

    # Формирование запроса на коррекцию
    correction_request = {
        "type": "sellCorrection" if row['Тип операции'] == 'приход' else "sellReturnCorrection",
        "taxationType": "usnIncomeOutcome", # Тут указываем СНО (в данном случае УСД Д-Р)
        "electronically": True, # Если True, то чек электронный
        "ignoreNonFiscalPrintErrors": False,
        "correctionType": "self",
        "correctionBaseDate": format_date(row['Дата чека']),
        "correctionBaseNumber": str(row['ФПД']),
        "operator": {
            "name": operator_name,
            "vatin": operator_inn
        },
        "clientInfo": {
            "emailOrPhone": "mail@mail.ru" # Тут указываем почту, на которую будут приходить все пробитые чеки
        },
        "items": items_data,
        "payments": [
            {"type": "electronically", "sum": row['Безналичными']},
            {"type": "cash", "sum": row['Наличными']}
        ],
        "total": row['Сумма чека']
    }

    # Добавление запроса в список
    correction_requests.append(correction_request)

# Сохранение всех запросов в один JSON-файл
json_filename = 'correction_requests.json'
with open(json_filename, 'w', encoding='utf-8') as json_file:
    json.dump(correction_requests, json_file, ensure_ascii=False, indent=4)

print(f'Все JSON запросы сохранены в {json_filename}')
