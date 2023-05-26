from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys
from time import sleep
import pyperclip
import pandas as pd
from dotenv import load_dotenv
import os
import json
import subprocess


def load_environment_variables(env_file):
    #env = {}
    load_dotenv(env_file, override=True)
    env = dict(os.environ)
    return env


def access_last_row_link(driver, link_last_row):
    driver.get(link_last_row)
    sleep(2)
    a = ActionChains(driver)
    a.key_down(Keys.CONTROL).send_keys('C').key_up(Keys.CONTROL).perform()


def access_interval_link(driver, link_interval):
    driver.get(link_interval)
    a = ActionChains(driver)
    a.key_down(Keys.CONTROL).send_keys('C').key_up(Keys.CONTROL).perform()


def get_dataframe_from_clipboard():
    df = pd.read_clipboard()
    return df


def format_dataframe(df, metric_columns):
    df.columns = ['DATE'] + metric_columns
    df['DATE'] = pd.to_datetime(df['DATE'], format='%Y-%m-%d %H:%M:%S').dt.strftime('%Y-%m-%d')
    df = df.replace(',', '.', regex=True)
    for column in metric_columns:
        df[column] = df[column].astype(float)
    return df


def merge_dataframes(df_sheet, df_bq, metric_columns):
    df_merge = pd.merge(df_sheet, df_bq, on='DATE', suffixes=('_sheet', '_bq'))
    return df_merge


def calculate_differences(row, metric_columns, percentagediff):
    diffs = []
    for column in metric_columns:
        col1 = f'{column}_sheet'
        col2 = f'{column}_bq'
        try:
            diff_pct = ((row[col2] - row[col1]) / row[col1]) * 100
            if abs(diff_pct) > percentagediff:
                diffs.append({
                    'DATE': row['DATE'],
                    'Dimensão': column,
                    'Diferença (%)': round(diff_pct, 2)
                })
        except KeyError:
            continue
    return diffs


def generate_output_data(df_resulted, publisher):
    if df_resulted.empty:
        output_data = {
            'name': f'{publisher} - Ok',
            'result': [],
            'dates': ''
        }
    else:
        dates = df_resulted['DATE'].unique()
        dates_line = ' '.join(dates)
        output_data = {
            'name': f'Diferenças para o veículo {publisher}',
            'result': df_resulted.to_dict(orient='records'),
            'dates': dates_line
        }
    return output_data
#
# print(f"{df_resulted.name}\n\
# {df_resulted}\n\
# Datas distintas: {dates_line}")

#df_resulted_dict = df_resulted.to_dict(orient='records')

def save_output_data(output_data, output_file):
    with open('resultados.json', 'w') as json_file:
        json.dump(output_data, json_file, indent=4)


def main():
    env_files = ["a.env", "b.env"]

    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--user-data-dir=C:/Users/Gustavo/AppData/Local/Google/Chrome/User Data")
    chrome_options.add_argument("--profile-directory=Default")

    for env_file in env_files:
        env = load_environment_variables(env_file)

        pd.set_option('display.max_columns', None)  # Exibir todas as colunas
        pd.set_option('display.expand_frame_repr', False)  # Não

        publisher = env.get('PUBLISHER')
        linkSheet = env.get('LINKSHEET')
        lastRowCell = env.get('LASTROWCELL')
        rowDays = int(env.get('ROWDAYS'))
        metric_columns = [env.get('METRIC1'), env.get('METRIC2'), env.get('METRIC3')]
        percentagediff = int(env.get('PERCENTAGEDIFF'))


        driver = webdriver.Chrome('C:/Users/Gustavo/Desktop/automate/chromedriver.exe', options=chrome_options)

        linkLastRow = f'{linkSheet}&range={lastRowCell}'
        access_last_row_link(driver, linkLastRow)

        numberLastRow = pyperclip.paste()

        numberFirstRowIndex = int(numberLastRow) - rowDays
        linkInterval = f'{linkSheet}&range=A{numberFirstRowIndex}:D{numberLastRow}'
        access_interval_link(driver, linkInterval)

        df_sheet = get_dataframe_from_clipboard()
        df_sheet = format_dataframe(df_sheet, metric_columns)


        df_bq = pd.read_excel('bq.xlsx', sheet_name='Planilha1', header=1, engine='openpyxl')
        df_bq = format_dataframe(df_bq, metric_columns)


        df_merge = merge_dataframes(df_sheet, df_bq, metric_columns)
        print(df_merge)

        resulted = df_merge.apply(calculate_differences, axis=1, args=(metric_columns, percentagediff,))
        df_resulted = pd.DataFrame([item for sublist in resulted for item in sublist])


        output_data = generate_output_data(df_resulted, publisher)
        save_output_data(output_data, 'resultados.json')

        subprocess.run(['python', 'botDiscord.py'])

        driver.quit()

if __name__ == "__main__":
    main()


