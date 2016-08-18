from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions
import time
import openpyxl

# this script loops through the IDs in Coupa Expense Delegates excel file and does something

browser = webdriver.Firefox()
browser.get('https://deloitte-ca2.coupacloud.com/session')
browser.implicitly_wait(1)

wbx = openpyxl.load_workbook('Login.xlsx')
shx = wbx.get_sheet_by_name('Sheet1')
Login = shx.cell(row=1, column=1).value
Password = shx.cell(row=1, column=2).value

browser.find_element_by_id('user_login').send_keys(Login)
browser.find_element_by_id('user_password').send_keys(Password)
browser.find_element_by_class_name('button').click()

wb = openpyxl.load_workbook('Coupa_Expense_Delegates.xlsx')
sh = wb.get_sheet_by_name('Sheet1')
MaxRow = sh.max_row

for x in range(1,MaxRow+1):
    AccountID = sh.cell(row=x, column=1).value
    print(AccountID)
    browser.get('https://deloitte-ca2.coupacloud.com/user/edit_user/%s' % AccountID)
    SampleText = browser.find_element_by_id('pageHeader').text
    print(SampleText)

    FullName = sh.cell(row=x, column=2).value
    browser.find_element_by_id('user_can_delegate_expenses_to_ids_name').send_keys(FullName)

    for y in range(1,10):
        try:
            print(browser.find_element_by_id('ui-id-%s' % y).text)
            if sh.cell(row=x,column=3).value in browser.find_element_by_id('ui-id-%s' % y).text:
                browser.find_element_by_id('ui-id-%s' % y).click()
                sh.cell(row=x, column=4).value = 'Delegate Assigned'
                browser.find_element_by_xpath("//button[@class='button blue'][@type='submit']").click()
        except:
            pass

browser.close()
wb.save('Coupa_Expense_Delegate_Output.xlsx')

