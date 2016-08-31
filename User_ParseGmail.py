from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.common.action_chains import ActionChains
import time
import openpyxl

# updating invoices enmass to flag the 'paid' flag to checked (yes) for
# a list of invoices in a spreadsheet that are not already flagged

StartTime = time.time()

wb = openpyxl.load_workbook('User_ParseGmail_Output.xlsx')
shx = wb.get_sheet_by_name('Login')
Login = shx.cell(row=2, column=1).value
Password = shx.cell(row=2, column=2).value
clientURL = shx.cell(row=2, column = 3).value

driver = webdriver.Chrome()
driver.get(clientURL)
driver.implicitly_wait(10)
