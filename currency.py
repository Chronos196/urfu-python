import pandas as pd
import requests
from xml.etree import ElementTree as ET

currencies = ['USD', 'RUR', 'EUR', 'KZT', 'UAH', 'BYR']
pr = pd.period_range(start='2005-10', end='2022-07', freq='M')
prTupes = [(period.month,period.year) for period in pr]

df = pd.DataFrame()

for (m, y) in prTupes:
    responce = requests.get(f'http://www.cbr.ru/scripts/XML_daily.asp?date_req=01/{m}/{y}')
    tree = ET.fromstring(responce.content).findall('Valute')
    info = {}
    for cur in currencies:
        for valute in tree:
            if valute.find('CharCode').text == cur:
                info[cur] = float(valute.find('Value').text.replace(',', '.')) / int(valute.find('Nominal').text.replace(',', '.'))
                break
    
    df = df.append(pd.DataFrame(info, [f"{y}-{m}"]))

df.index.rename('date', inplace=True)

df.to_csv('currencies_df.csv')
