from urllib.parse import urlencode
from urllib.request import urlopen, Request
from datetime import datetime
import pandas as pd
import time


# пользовательские переменные
# ticker = "FXIT"
# задаём тикер
period = 8  # задаём период. Выбор из: 'tick': 1, 'min': 2, '5min': 3, '10min': 4, '15min': 5, '30min': 6, 'hour': 7, 'daily': 8, 'week': 9, 'month': 10
start = "01.10.2020"  # с какой даты начинать тянуть котировки
end = datetime.now().date().strftime("%d.%m.%Y") # финальная дата, по которую тянуть котировки
########
periods = {'tick': 1, 'min': 2, '5min': 3, '10min': 4, '15min': 5, '30min': 6, 'hour': 7, 'daily': 8, 'week': 9,
           'month': 10}
# print("ticker=" + ticker + "; period=" + str(period) + "; start=" + start + "; end=" + end)
# каждой акции Финам присвоил цифровой код:
tickers = {'FXIT': 181750,
           'FXWO': 927606,
           'FXUS': 181754,
           'FXCN': 385054,
           'YNDX': 388383,
           'BA-RM': 2028080,
           'AMD-RM': 2028078,
           'AMZN-RM': 2028087,
           'DIS-RM': 2028091,
           'FIVE': 491944,
           'GOOG-RM': 2028079,
           'INTC-RM': 2028093,
           'AAPL-RM': 2052021,
           'MAIL': 1938060,
           'MU-RM': 2028094,
           'MSFT-RM': 2028086,
           'NFLX-RM': 2028081,
           'NVDA-RM': 2028084,
           'SBER': 3,
           'MGNT': 17086,
           'FB-RM': 2028085,
           'V-RM': 2028083,
           'BABA-RM': 2139232,
           'KO-RM': 2190358,
           'EA-RM': 2139237,
           'IBM-RM': 2190355,
           'JNJ-RM': 2190356,
           'NKE-RM': 2190360,
           'OZON': 2179435,
           'PG-RM': 2190361,
           'PFE-RM': 2028095,
           'QCOM-RM': 2139245,
           'TSLA-RM': 2139246,
           'TWTR-RM': 2028092,
           'CRM-RM': 2190353,
           'HPQ-RM': 2139242
           }
FINAM_URL = "http://export.finam.ru/"  # сервер, на который стучимся
market = 0  # можно не задавать. Это рынок, на котором торгуется бумага. Для акций работает с любой цифрой. Другие рынки не проверял.
# Делаем преобразования дат:
start_date = datetime.strptime(start, "%d.%m.%Y").date()
start_date_rev = datetime.strptime(start, '%d.%m.%Y').strftime('%Y%m%d')
end_date = datetime.strptime(end, "%d.%m.%Y").date()
end_date_rev = datetime.strptime(end, '%d.%m.%Y').strftime('%Y%m%d')
# Все параметры упаковываем в единую структуру. Здесь есть дополнительные параметры, кроме тех, которые заданы в шапке. См. комментарии внизу:
df = pd.DataFrame()
for ticker in tickers.keys():
    params = urlencode([
        ('market', market),  # на каком рынке торгуется бумага
        ('em', tickers[ticker]),  # вытягиваем цифровой символ, который соответствует бумаге.
        ('code', ticker),  # тикер нашей акции
        ('apply', 0),  # не нашёл что это значит.
        ('df', start_date.day),  # Начальная дата, номер дня (1-31)
        ('mf', start_date.month - 1),  # Начальная дата, номер месяца (0-11)
        ('yf', start_date.year),  # Начальная дата, год
        ('from', start_date),  # Начальная дата полностью
        ('dt', end_date.day),  # Конечная дата, номер дня
        ('mt', end_date.month - 1),  # Конечная дата, номер месяца
        ('yt', end_date.year),  # Конечная дата, год
        ('to', end_date),  # Конечная дата
        ('p', period),  # Таймфрейм
        ('f', ticker + "_" + start_date_rev + "_" + end_date_rev),  # Имя сформированного файла
        ('e', ".csv"),  # Расширение сформированного файла
        ('cn', ticker),  # ещё раз тикер акции
        ('dtf', 1),
        # В каком формате брать даты. Выбор из 5 возможных. См. страницу https://www.finam.ru/profile/moex-akcii/sberbank/export/
        ('tmf', 1),  # В каком формате брать время. Выбор из 4 возможных.
        ('MSOR', 0),  # Время свечи (0 - open; 1 - close)
        ('mstime', "on"),  # Московское время
        ('mstimever', 1),  # Коррекция часового пояса
        ('sep', 1),
        # Разделитель полей    (1 - запятая, 2 - точка, 3 - точка с запятой, 4 - табуляция, 5 - пробел)
        ('sep2', 1),  # Разделитель разрядов
        ('datf', 1),  # Формат записи в файл. Выбор из 6 возможных.
        ('at', 1)])  # Нужны ли заголовки столбцов
    url = FINAM_URL + ticker + "_" + start_date_rev + "_" + end_date_rev + ".csv?" + params  # урл составлен!
    # print("Стучимся на Финам по ссылке: " + url)
    one_stock = pd.read_csv(urlopen(Request(url, headers={'User-Agent': 'Mozilla/5.0'})), sep=',')
    df = pd.concat([df, one_stock])
    time.sleep(2)
df.to_excel('report_stock_price.xlsx', engine='xlsxwriter')
