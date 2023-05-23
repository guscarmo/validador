from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys
from time import sleep
import pyperclip
import pandas as pd
from dotenv import load_dotenv
import os

publisher = 'GooExample'

pd.set_option('display.max_columns', None)  # Exibir todas as colunas
pd.set_option('display.expand_frame_repr', False)  # Não

chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--user-data-dir=C:/Users/Gustavo/AppData/Local/Google/Chrome/User Data")
chrome_options.add_argument("--profile-directory=Default")

driver = webdriver.Chrome('C:/Users/Gustavo/Desktop/automate/chromedriver.exe', options=chrome_options)

linkSheet = 'https://docs.google.com/spreadsheets/d/1m8AEqwd0_E9JxAjoR1xG9KEsHB7m5CJ5HKIiEMTiUHg/edit#gid=0'
lastRowCell = 'H1'
rowDays = 7

load_dotenv()
Metric1 = os.getenv('Metric1')
Metric2 = os.getenv('Metric2')
Metric3 = os.getenv('Metric3')

print(Metric1)
print(Metric2)

linkLastRow = f'{linkSheet}&range={lastRowCell}'

driver.get(linkLastRow)

sleep(2)

a = ActionChains(driver)
a.key_down(Keys.CONTROL).send_keys('C').key_up(Keys.CONTROL).perform()


numberLastRow = pyperclip.paste()
numberFirstRowIndex = int(numberLastRow) - rowDays
linkInterval = f'{linkSheet}&range=A{numberFirstRowIndex}:D{numberLastRow}'

driver.get(linkInterval)
a.key_down(Keys.CONTROL).send_keys('C').key_up(Keys.CONTROL).perform()

df_sheet = pd.read_clipboard()
column_names = ['DATE', Metric1, Metric2, Metric3]
df_sheet.columns = column_names

df_sheet['DATE'] = pd.to_datetime(df_sheet['DATE'], format='%Y-%m-%d %H:%M:%S').dt.strftime('%Y-%m-%d')

df_sheet = df_sheet.replace(',', '.', regex=True)

df_sheet[Metric1] = df_sheet[Metric1].astype(float)
df_sheet[Metric2] = df_sheet[Metric2].astype(float)
df_sheet[Metric3] = df_sheet[Metric3].astype(float)


df_bq = pd.read_excel('bq.xlsx', sheet_name='Planilha1', header=1, engine='openpyxl')
column_names = ['DATE', Metric1, Metric2, Metric3]
df_bq.columns = column_names

df_bq['DATE'] = pd.to_datetime(df_bq['DATE'], format='%Y-%m-%d %H:%M:%S').dt.strftime('%Y-%m-%d')

df_bq = df_bq.replace(',', '.', regex=True)

df_bq[Metric1] = df_bq[Metric1].astype(float)
df_bq[Metric2] = df_bq[Metric2].astype(float)
df_bq[Metric3] = df_bq[Metric3].astype(float)

df_merge = pd.merge(df_sheet, df_bq, on='DATE', suffixes=('_sheet', '_bq'))

#print(df_bq)

#print(df_merge)


def difference(row):
    diffs = []
    for column in [Metric1, Metric2, Metric3]:
        col1 = f'{column}_sheet'
        col2 = f'{column}_bq'
        try:
            diff_pct = ((row[col2] - row[col1]) / row[col1]) * 100
            if abs(diff_pct) > 5:
                diffs.append({
                    'DATE': row['DATE'],
                    'Dimensão': column,
                    'Diferença (%)': diff_pct
                })
        except KeyError:
            continue
    return diffs

resulted = df_merge.apply(difference, axis=1)

df_resulted = pd.DataFrame([item for sublist in resulted for item in sublist])
df_resulted.name = f'Diferenças para o veículo {publisher}'

dates = df_resulted['DATE'].unique()
dates_line = ' '.join(dates)


print(f"{df_resulted.name}\n\
{df_resulted}\n\
Datas distintas: {dates_line}")
