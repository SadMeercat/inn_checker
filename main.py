from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
import time
import pandas as pd
from datetime import datetime
from pprint import pprint

service = Service('chromedriver.exe')
driver = webdriver.Chrome(service=service)

driver.get("https://npd.nalog.ru/check-status/")


def parse(inn, date):
    # Пример поиска полей и заполнения данных

    # Найдите поле для ИНН и заполняйте его
    inn_field = driver.find_element(By.ID, "ctl00_ctl00_tbINN")  # измените селектор по необходимости
    inn_field.send_keys(inn)

    # Найдите поле для даты и заполняйте его
    date_field = driver.find_element(By.ID, "ctl00_ctl00_tbDate")
    date_field.send_keys(date)

    # Использование execute_script для выполнения JavaScript из onclick
    login_button = driver.find_element(By.ID, "ctl00_ctl00_btSend")
    driver.execute_script("arguments[0].click();", login_button)

    # Время ожидания для визуализации
    time.sleep(0.5)

    # Получаем результат
    result = driver.find_element(By.ID, "ctl00_ctl00_lblInfo").text

    # Очищаем поля для следующей итерации
    inn_field = driver.find_element(By.ID, "ctl00_ctl00_tbINN")
    inn_field.clear()
    date_field = driver.find_element(By.ID, "ctl00_ctl00_tbDate")
    date_field.clear()

    return result


def format_date(date):
    if isinstance(date, pd.Timestamp):  # Проверка, что date — это pd.Timestamp
        if pd.isna(date):  # Проверка на NaT
            return 'nan'  # Или любая другая строка для NaT
        return date.strftime('%d-%m-%Y')  # Конвертация напрямую в строку нужного формата
    elif isinstance(date, str):  # Если уже строка
        # Пробуем разные форматы даты
        try:
            # Сначала пробуем формат с временем
            dt = datetime.strptime(date.strip(), '%d.%m.%Y %H:%M:%S')
        except ValueError:
            try:
                # Если не получилось, пробуем без времени
                dt = datetime.strptime(date.strip(), '%d.%m.%Y')
            except ValueError:
                print(f"Ошибка: Неподдерживаемый формат даты: {date}")
                return 'nan'
        return dt.strftime('%d-%m-%Y')
    elif isinstance(date, datetime):
        return date.strftime('%d-%m-%Y')  # Преобразование напрямую
    else:
        print(f"Ошибка: Неподдерживаемый тип данных для даты: {type(date)}")
        return 'nan'


datas = []


df = pd.read_excel('./data.xlsx')


today_date = datetime.today().strftime('%d-%m-%Y')


for index, row in df.iterrows():
    inn = row['inn']
    dates2 = []
    dates1 = row['date_start']
    dates1 = format_date(dates1)


    dates = row['date_other']
    #print(type(dates))
    #print(dates)
    if isinstance(dates, float) or pd.isnull(dates):
        dates2.append('nan')
    elif isinstance(dates, datetime):
        dates2.append(format_date(dates))
    else:
        dates = dates.split(',')
        # print(dates)
        for date in dates:
            dates2.append(format_date(date.strip()))
    print(dates2)
    dates2.append(today_date)
    if len(str(inn)) != 12:
        datas.append({'name': inn, 'date': dates1})
    else:
        datas.append({'inn': inn, 'date_start': dates1, 'date_other': dates2})
    # datas.append({'inn': inn, 'date_start': dates1})
# pprint(datas)


try:
    results = []
    for data in datas:
        if 'inn' in data:
            result1 = parse(data['inn'], data['date_start'])
            tmp_results = []
            if not 'nan' in data['date_other']:
                for date in data['date_other']:
                    tmp_results.append({"date": date ,'result' :parse(data['inn'], date)})
            else:
                tmp_results.append('nan')
            results.append({"inn": data["inn"], "date_start": data["date_start"], 'date_other': data['date_other'], 
                            "result_start": result1, 
                            "result_other": '\n'.join(f"{item['date']}: {item['result']}" for item in tmp_results) if not 'nan' in tmp_results else ""})
            # results.append({"inn": data["inn"], "date_start": data["date_start"], 
            #                 "result_start": result1, })

        else: 
            results.append({"inn": data["name"], "date": data["date"], "result": "Неверный ИНН"})


finally:
    # Закрытие драйвера
    driver.quit()


# Создаем DataFrame
df = pd.DataFrame(results)


# Сохраняем в Excel с переносами строк
with pd.ExcelWriter("results.xlsx", engine="xlsxwriter") as writer:
    df.to_excel(writer, index=False, sheet_name="Sheet1")


    # Получаем объект рабочего листа
    worksheet = writer.sheets["Sheet1"]


    # Включаем перенос текста для всех ячеек
    for col_num, col_name in enumerate(df.columns):
        worksheet.set_column(col_num, col_num, 20, writer.book.add_format({'text_wrap': True}))