import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

file_name = input("Write path to xlsx - ")

excel_file = openpyxl.load_workbook(file_name)
excel_sheet = excel_file.active

data_column = 'A'
result_column = 'B'

driver = webdriver.Chrome()
driver.get("https://www6.vid.gov.lv/PVN")

for row in range(1, excel_sheet.max_row + 1):
    driver.refresh()

    field = driver.find_element(By.NAME, "Code")
    button = driver.find_element(By.NAME, "search")

    field.send_keys(excel_sheet[f"{data_column}{row}"].value)
    button.click()

    try:
        myElem = WebDriverWait(driver, 4).until(EC.presence_of_element_located((By.TAG_NAME, "td")))

        res = driver.find_elements(By.TAG_NAME, "td")
        elements = [elem.text for elem in res]

        excel_sheet[f'{result_column}{row}'].value = elements[2]
    except Exception:
        excel_sheet[f'{result_column}{row}'].value = "unknown"

excel_file.save(file_name)
driver.quit()
