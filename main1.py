import os.path
import io
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from time import sleep
from openpyxl import load_workbook
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
driver.get("https://www.webcat-solutions.com/ro/portal/loginNew.aspx?14=21")
sleep(5)
username_box = driver.find_element(By.XPATH, "//input[@name='txtUser']")
password_box = driver.find_element(By.XPATH, "//input[@name='txtPassword']")
submit_button = driver.find_element(By.XPATH, "//input[@name='btnLogin']")
username_box.click()
username_box.send_keys("toptech")
password_box.click()
password_box.send_keys("bodale2021")
submit_button.click()
sleep(5)
workbook = load_workbook("C:\\Users\\HP\\Desktop\\DATABASE_AUTONET\\lista_artiole.xlsx")
worksheet = workbook["lista_artiole_wes_ORIGINAL"]
column_code = worksheet["A"]
code_list = [column_code[x].value for x in range(len(column_code))]

for code in code_list[50000:]:
    code = code.replace("/", "")
    if os.path.isfile(f"C:\\Users\\HP\\Desktop\\DATABASE_AUTONET\\code_full_page\\{code}.html"):
        print(f"Produs extractat: {code}")
    else:
        iframe = driver.find_element(By.XPATH, "//iframe[@name='iframe_extern_promo']").get_attribute("src")
        driver.execute_script('''window.open("about:blank");''')
        driver.switch_to.window(driver.window_handles[1])
        driver.get(iframe)
        sleep(5)
        session_expired = driver.find_element(By.XPATH, "//div[@class='alert alert-danger m-1']")
        print("Sesiune expirata!!")
        driver.close()
        sleep(5)
        driver.get("https://www.webcat-solutions.com/ro/portal/loginNew.aspx?14=21")
        sleep(5)
        username_box = driver.find_element(By.XPATH, "//input[@name='txtUser']")
        password_box = driver.find_element(By.XPATH, "//input[@name='txtPassword']")
        submit_button = driver.find_element(By.XPATH, "//input[@name='btnLogin']")
        username_box.click()
        username_box.send_keys("toptech")
        password_box.click()
        password_box.send_keys("bodale2021")
        submit_button.click()
        sleep(5)
        # except NoSuchElementException:
        #     print("Sesiune in curs")
        #     print(f"In proces: {code}")
        #     sleep(5)
        #     code_insert = driver.find_element(By.XPATH, "//input[@id='tp_articlesearch_txt_art_direkt']")
        #     code_submit_button = driver.find_element(By.XPATH, "//div[@id='tp_articlesearch_artikelsuche_button_trader']/input")
        #     code_insert.send_keys(Keys.CONTROL+"A")
        #     code_insert.send_keys(Keys.BACKSPACE)
        #     code_insert.send_keys(code)
        #     code_submit_button.click()
        #     sleep(10)
        #     try:
        #         full_page = driver.find_element(By.XPATH, "//div[@id='content_table_pnl']").get_attribute("innerHTML")
        #         html = open(f"C:\\Users\\HP\\Desktop\\DATABASE_AUTONET\\code_full_page\\{code}.html", "w", encoding="utf-8")
        #         html.write(full_page)
        #         html.close()
        #         sleep(1)
        #     except NoSuchElementException:
        #         pass
