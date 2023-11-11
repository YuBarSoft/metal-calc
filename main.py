import time
import winsound
import pandas as pd
from subprocess import Popen
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

options = webdriver.ChromeOptions()
options.headless = False
options.page_load_strategy = "eager"
driver = webdriver.Chrome(options=options)

start_time = time.time()
URL = 'https://metal-calculator.ru/page/app'
driver.get(URL)

df = pd.read_csv('tz.csv', delimiter=';', encoding='windows-1251')
metal = df.loc[0, 'Наименование']  # берем металл из наименования
if 'Алюминиевый' in metal:
    metal = 'Алюминий'
elif 'Латунный' in metal:
    metal = 'Латунь'
else:
    print('Это точно алюминиевый или латунный лист?')
    exit()

driver.find_element(By.LINK_TEXT, metal).click()
driver.find_element(By.LINK_TEXT, 'Лист/плита').click()

print('Распарсиваем файл CSV')
columns_list = ['Наименование', 'Код артикула', 'Марка', 'Толщина', 'Ширина', 'Длина', 'Вес']
data = pd.DataFrame(columns=columns_list)
for_print = ''

select = Select(driver.find_element(By.NAME, 'p'))
marks = WebDriverWait(driver, 3).until(
    EC.presence_of_element_located((By.NAME, 'p'))
).text

for index, row in df.iterrows():
    name = row.iloc[0]  # Получение данных из столбца 'Наименование'
    artikul = row.iloc[1]  # Получение данных из столбца 'Код артикула'
    mark = row.iloc[2]  # Получение данных из столбца 'Марка'
    mark_for_csv = mark
    t = row.iloc[3]  # Получение данных из столбца 'Толщина'
    a = row.iloc[4]  # Получение данных из столбца 'Ширина'
    b = row.iloc[5]  # Получение данных из столбца 'Длина'
    for_print = f'{name};{artikul};{mark};{t};{a};{b}'

    if mark in marks:
        select.select_by_visible_text(mark)
    else:
        mark = 'Прочее'
        select.select_by_visible_text(mark)  # марки нет в выпадающем списке

    input_t = driver.find_element(By.NAME, 't')  # Толщина
    input_t.clear()
    input_t.send_keys(t)

    input_a = driver.find_element(By.NAME, 'a')  # Ширина
    input_a.clear()
    input_a.send_keys(a)

    input_b = driver.find_element(By.NAME, 'b')  # Длина
    input_b.clear()
    input_b.send_keys(b)

    n = 1  # Количество
    input_n = driver.find_element(By.NAME, 'n')
    input_n.clear()
    input_n.send_keys(n)

    driver.find_element(By.CLASS_NAME, 'btn-green').click()  # Рассчитать

    ves = driver.find_elements(By.CLASS_NAME, 'result-item-value')  # Парсим вес
    ves = ves[1].text
    ves = ves.replace('кг.', 'кг')
    print(f'{for_print};{ves}')

    row_data = {
        columns_list[0]: name,
        columns_list[1]: artikul,
        columns_list[2]: mark_for_csv,
        columns_list[3]: t,
        columns_list[4]: a,
        columns_list[5]: b,
        columns_list[6]: ves
    }
    data = data._append(row_data, ignore_index=True)  # Добавление строки в DataFrame

my_time = time.strftime('%d.%m.%Y_%H-%M', time.localtime())
output_file = f'result_{my_time}.csv'  # выходной файл
data.to_csv(output_file, sep=';', index=False, encoding='windows-1251')  # Запись DataFrame в CSV-файл
print(f'Время выполнения: {time.time() - start_time:.2f}с.')
winsound.PlaySound("SystemExit", winsound.SND_ALIAS)
print(f'Имя файла: {output_file}')
print('\nПрограмма закроется через 20 секунд.')
print('Открываем полученный файл в Excel')
Popen(output_file, shell=True)
time.sleep(20)
