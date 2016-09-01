from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.common.action_chains import ActionChains
import time
import openpyxl

# loops through an output file generated from GoogleAPI_quickstart.py which contains several
# links to reset an email address and sets all the passwords to a default (ie. Password)

wb = openpyxl.load_workbook('GoogleAPI_quickstart_Output.xlsx')
sh = wb.get_sheet_by_name('Sheet1')
MaxRow = sh.max_row

