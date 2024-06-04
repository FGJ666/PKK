import pandas as pd
import numpy as np

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC

adres = pd.read_excel('kad.xlsx', header=None)
llist = adres()

def startChrome():
    chrome_options = Options()
    driver = webdriver.Chrome(options=chrome_options)
    driver.set_page_load_timeout(30)
    WebDriverWait(driver, 2)
    return driver

def wait_element(path):
    try:
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, path)))
    except:
        driver.refresh()
    return driver.find_element_by_xpath(path)

def wait_element_text(path):
    try:
        element = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, path)))
    except:
        aaa = np.nan
    else:
        aaa = driver.find_element_by_xpath(path).text
    return aaa

# создаем пустые таблицы
df = df_temp = pd.DataFrame()

# запускаем браузер
driver = startChrome()
# переходим на веб-страницу ранее упомняутого сервиса
driver.get('https://rosreestr.net/proverit-kvartiru')
# циклически извлекаем кадастровые номера объектов недвижимости из списка
for i in list(range(0, len(llist))):
    # присваиваем пустые знаечния переменным
    a0 = a1 = a2 = np.nan
    # вводим кадастровый номер в поле поиска и нажимаем Enter
    wait_element(
        '//*[@id="search_main"]').send_keys(str(llist[i]), Keys.RETURN)
    # используем функцию ожидания до тех пор пока не исчезнет таймер поиска
    while driver.find_element_by_xpath('//*[@id="table_search_timer"]').is_displayed():
        time.sleep(1)
    else:
        # если совпадений не найдено, вывести сообщение
        if driver.find_element_by_xpath('//div[@class="table_search_not-title"]').text \
                == 'Совпадений не найдено... Что делать?':
            print('table_search_not')
        # иначе
        else:
            print('table_search')
            try:
                # нажать на кнопку выбрать у проверяемого объекта
                wait_element('//a[contains(@href,"/kadastr/")] \
                             /div[@class="table__btn"]').click()
                # задать переменной тип объекта
                a0 = wait_element_text('//div[@class="test__data"] \
                /div[contains(text(),"Тип")]/strong')
                # задать переменной кадастрвый номер объекта
                a1 = wait_element_text('//div[@class="test__data"] \
                /div[contains(text(),"Кадастровый номер")]/strong')
                # задать переменной адрес объекта
                a2 = wait_element_text('// div[@class="test__data"] \
                                       / div[contains(text(), "Адрес полный")]/strong')
                print(i, a0, a1, a2)
            except:
                print('no buton')
            # создать временную таблицу из набора переменных
            df_temp = pd.DataFrame([llist[i], a0, a1, a2]).transpose()
            # добавить временную таблицу в основную таблицу
            df = df.append(df_temp)
# по завершению цикла закрыть браузер
driver.close()
# сохранит результат в электронную таблицу
df.to_excel('result.xlsx', index=None)
