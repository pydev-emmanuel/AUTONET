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

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
driver.get("https://www.webcat-solutions.com/ro/portal/loginNew.aspx?14=21")
sleep(5)
username_box = driver.find_element(By.XPATH, "//input[@name='txtUser']")
password_box = driver.find_element(By.XPATH, "//input[@name='txtPassword']")
submit_button = driver.find_element(By.XPATH, "//input[@name='btnLogin']")
username_box.click()
username_box.send_keys("toptech1972")
password_box.click()
password_box.send_keys("2KH0BJF4")
submit_button.click()
sleep(5)
workbook = load_workbook("C:\\Users\\HP\\Desktop\\DATABASE_AUTONET\\lista_artiole.xlsx")
worksheet = workbook["lista_artiole_wes_ORIGINAL"]
column_code = worksheet["A"]
code_list = [column_code[x].value for x in range(len(column_code))]

for code in code_list:
    if os.path.isfile(f"C:\\Users\\HP\\Desktop\\DATABASE_AUTONET\\code_full_page\\{code}.html"):
        print(f"Produs extractat: {code}")
    else:
        print(f"In proces: {code}")
        sleep(5)
        code_insert = driver.find_element(By.XPATH, "//input[@id='tp_articlesearch_txt_art_direkt']")
        code_submit_button = driver.find_element(By.XPATH, "//div[@id='tp_articlesearch_artikelsuche_button_trader']/input")
        code_insert.send_keys(Keys.CONTROL+"A")
        code_insert.send_keys(Keys.BACKSPACE)
        code_insert.send_keys(code)
        code_submit_button.click()
        sleep(10)
        full_page = driver.find_element(By.XPATH, "//div[@id='content_table_pnl']").get_attribute("innerHTML")
        html = open(f"C:\\Users\\HP\\Desktop\\DATABASE_AUTONET\\code_full_page\\{code}.html", "w", encoding="utf-8")
        html.write(full_page)
        html.close()
        sleep(1)
