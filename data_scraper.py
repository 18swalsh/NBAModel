from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import xlsxwriter
import time
from openpyxl import load_workbook
import pandas as pd

start_time = time.clock()

option = webdriver.ChromeOptions()
option.add_argument("--incognito")

browser = webdriver.Chrome(executable_path="C:/Users/cwalsh/Documents/Steve/chromedriver/chromedriver")
browser.get("https://stats.nba.com/teams/advanced/?sort=W&dir=-1")

timeout = 10

try:
    WebDriverWait(browser, timeout).until(EC.visibility_of_element_located((By.XPATH,"//table")))
except TimeoutException:
    print("Timed out waiting for the page to load")
    browser.quit()

table_headers = browser.find_elements_by_xpath("//table/thead/tr/th")
headers = [x.text for x in table_headers]
# print(headers)
data_table = browser.find_elements_by_xpath("//table/tbody/tr/td")
values = [x.text for x in data_table]
# print(values)
browser.quit()

print("Values retrieved")

# create excel workbook
workbook = xlsxwriter.Workbook("C:/Users/cwalsh/Documents/Steve/NBA_Output_"+str(round(time.time(),2))+".xlsx")

# Create output sheet
workbook.add_worksheet("Teams General Traditional")
workbook.add_worksheet("Teams General Advanced")
workbook.add_worksheet("Teams General Four Factors")
workbook.add_worksheet("Teams General Misc")
workbook.add_worksheet("Teams General Scoring")
workbook.add_worksheet("Teams General Opponent")
workbook.add_worksheet("Teams General Defense")


# activate sheet
worksheet = workbook.get_worksheet_by_name("Teams General Advanced")


# 20 headers
for y in range(0,20):
    worksheet.write(0,y,headers[y])

# values
index = 0
for z in range(1,31):
    for x in range(0,20):
        worksheet.write(z,x,values[index])
        index += 1

workbook.close()

print("Model created in " + str(round(time.clock() - start_time,2)) + " seconds")
