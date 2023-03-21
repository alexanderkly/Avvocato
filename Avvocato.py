from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from time import sleep
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from conftest import driver, is_visible_by_xpath, is_visible_by_ID
import openpyxl
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
import undetected_chromedriver as uc
import json

url = 'https://sfera.sferabit.com/servizi/alboonlineBoot/index.php?id=1080'

minicount = 0
count = 281
page = 1
driver.get(url=url)
is_visible_by_xpath("//*[@class='select classOnchange form-control form-control-sm']")
second = driver.find_element(By.XPATH, "//*[@class='select classOnchange form-control form-control-sm']")
second.click()
is_visible_by_ID('filtroIdTipiAnagraficheCategorie')
select_element = Select(driver.find_element(By.ID, 'filtroIdTipiAnagraficheCategorie'))
selectel = select_element.select_by_value("1001")
is_visible_by_xpath('//*[@class="btn btn-primary"]')

enter = driver.find_element(By.XPATH, '//*[@class="btn btn-primary"]')
enter.click()

workbook = openpyxl.Workbook()
worksheet = workbook.active

# scroll to desired page
for i in range(1, 15):
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[text()="Avanti >"]')))
    arrow_r = driver.find_element(By.XPATH, '//*[text()="Avanti >"]')
    arrow_r.click()
    if i == 15:
        break

try:
    while True:
        is_visible_by_ID('dettagli')
        elements = driver.find_elements(By.ID, 'dettagli')
        for element in elements:
            count += 1
            minicount += 1
            sleep(1)
            # antibot waiting
            if minicount == 30:
                sleep(600)
                minicount = 0
            sleep(1)

            element.click()
            is_visible_by_xpath("(//*[contains(text(),'@')])[2]")
            try:
                mail = driver.find_element(By.XPATH, "(//*[contains(text(),'@')])[2]").text
            except Exception:
                mail = ' '
                print('eror')
                pass
            try:
                name = driver.find_element(By.XPATH, "(//*[contains(text(),'Avv')])[1]").text
            except Exception:
                name = ' '
                print('eror')
                pass
            is_visible_by_xpath('//*[@class="btn btn-secondary"]')
            close = driver.find_element(By.XPATH, '//*[@class="btn btn-secondary"]')
            close.click()

            row = [count, name, mail]
            worksheet.append(row)
            print(count, name, mail)
            page += 1
            with open("italy280.json", "a", encoding='utf-8') as f:
                json.dump(row, f, ensure_ascii=False)
                f.write(',\n')

        is_visible_by_xpath('//*[text()="Avanti >"]')
        arrow_r = driver.find_element(By.XPATH, '//*[text()="Avanti >"]')
        arrow_r.click()
        sleep(2)

except Exception as ex:
    print(ex)
    workbook.save('ItalyF280.xlsx')
finally:
    driver.close()
    driver.quit()
    workbook.save('ItalyF280.xlsx')
