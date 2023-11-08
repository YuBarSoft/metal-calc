from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from subprocess import Popen
import time
import datetime
import winsound
import pandas as pd

options = webdriver.ChromeOptions()
options.page_load_strategy = 'none'
driver = webdriver.Chrome(options=options)

my_time = datetime.datetime.now().strftime("%d.%m.%Y_%H-%M-%S")


if __name__ == '__main__':
    HOST = 'https://metal-calculator.ru/page/app'
    driver.get(HOST)  # переходим на начальную страницу парсинга

    metals = WebDriverWait(driver, 3).until(
        EC.presence_of_element_located((By.CLASS_NAME, 'unit-40'))
    ).text
    driver.execute_script("window.stop();")

    metals = metals.split()
    metals = metals[1:]
    print(f'Вид металла. Введите ниже соответствующую цифру:')

    count = 0
    for metal in metals:
        count += 1
        print(f'{count} - {metal}')
    print(f'-' * 30)

    m = int(input())
    print('Кликаем по ссылке')
    driver.find_element(By.LINK_TEXT, metals[m - 1]).click()

    types = driver.find_element(By.CLASS_NAME, 'unit-60').text
    types = types.split('\n')
    types = types[1:]
    print(f'Вид проката. Введите ниже соответствующую цифру:')
    count = 0

    for type in types:
        count += 1
        print(f'{count} - {type}')
    print(f'-' * 30)

    s = int(input())

    print('Кликаем по виду проката')
    driver.find_element(By.LINK_TEXT, types[s - 1]).click()

    marks = WebDriverWait(driver, 3).until(
        EC.presence_of_element_located((By.NAME, 'p'))
    ).text
    count = 0
    marks = marks.strip()
    marks = marks.replace(' ', '').replace('\n\n', '\n').split('\n')
    print(marks)
    print('Выбираем марку')
    for mark in marks:
        count += 1
        # print(f'{count} - {mark}')

    # CSV-ФАЙЛЫ
    tz_file = 'tz.csv'  # исходный файл
    output_file = f'result_{my_time}.csv'  # выходной файл

    print('Распарсиваем файл CSV')

    df = pd.read_csv(tz_file, delimiter=';', encoding='windows-1251', dtype='str')  # читаем файл csv

    data = pd.DataFrame(columns=['Наименование', 'Код артикула', 'Металл', 'Сортамент', 'Марка', 'Толщина', 'Ширина', 'Длина', 'Вес'])

    for index, row in df.iterrows():
        name = row.iloc[0]   # Получение данных из столбца 'Наименование'
        artikul = row.iloc[1]  # Получение данных из столбца 'Код артикула'
        metal = row.iloc[2]  # Получение данных из столбца 'Металл'
        type = row.iloc[3]  # Получение данных из столбца 'Сортамент'
        mark = row.iloc[4]  # Получение данных из столбца 'Марка'
        t = row.iloc[5]  # Получение данных из столбца 'Толщина'
        a = row.iloc[6]  # Получение данных из столбца 'Ширина'
        b = row.iloc[7]  # Получение данных из столбца 'Длина'
        # ves = row.iloc[8]  # Получение данных из столбца 'Вес'

        print(f'{name} - {artikul} - {metal} - {type} - {mark} - {t} - {a} - {b}')

        if mark in marks:
            select = Select(driver.find_element(By.NAME, 'p'))
            select.select_by_visible_text(mark)
        else:
            # print('Ошибка при выборе марки металла. Присваиваем дефолтную марку.')
            select = Select(driver.find_element(By.NAME, 'p'))
            print(mark)
            mark = 'Прочее'
            select.select_by_visible_text(mark)

        #print('Толщина')
        driver.find_element(By.NAME, 't').click()
        driver.find_element(By.NAME, 't').clear()
        driver.find_element(By.NAME, 't').send_keys(t)

        #print('Ширина')
        driver.find_element(By.NAME, 'a').click()
        driver.find_element(By.NAME, 'a').clear()
        driver.find_element(By.NAME, 'a').send_keys(a)

        #print('Длина')
        driver.find_element(By.NAME, 'b').click()
        driver.find_element(By.NAME, 'b').clear()
        driver.find_element(By.NAME, 'b').send_keys(b)

        #print('Количество')
        n = 1
        driver.find_element(By.NAME, 'n').click()
        driver.find_element(By.NAME, 'n').clear()
        driver.find_element(By.NAME, 'n').send_keys(n)

        #print('Кликаем по кнопке "Рассчитать"')
        driver.find_element(By.CLASS_NAME, 'btn-green').click()

        #print('Парсим вес')
        ves = driver.find_elements(By.CLASS_NAME, 'result-item-value')
        ves = ves[1].text
        print(ves)

        row_data = {
            'Наименование': name,
            'Код артикула': artikul,
            'Металл': metal,
            'Сортамент': type,
            'Марка': mark,
            'Толщина': t,
            'Ширина': a,
            'Длина': b,
            'Вес': ves
        }
        data = data._append(row_data, ignore_index=True)  # Добавление строки в DataFrame

    data.to_csv(output_file, sep=';', index=False, encoding='windows-1251')  # Запись DataFrame в CSV-файл

    end_time = datetime.datetime.now().strftime("%d.%m.%Y_%H-%M-%S")
    print(my_time + ' - Программа начала работу')
    print(end_time + ' - Программа завершена')

    winsound.PlaySound("SystemExit", winsound.SND_ALIAS)
    print(f'Месторасположение файла: {output_file}')
    print('Открываем полученный файл в Excel')
    Popen(output_file, shell=True)
    print('МОЖЕТЕ ЗАКРЫТЬ ПРОГРАММУ')

    time.sleep(300)
