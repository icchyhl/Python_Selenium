from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions
import time
import openpyxl
from selenium.webdriver.common.action_chains import ActionChains

# this example loops through an excel file and logs in with the information from
# Columns A and B and then writes a value into Column C saving it in a new excel file
# it demonstrates logging in and out of Coupa
# it also sets notifications for the user either ON or OFF depending on the true/false in
# the "select the notifications online & email" section of the code below
# 123

browser=webdriver.Firefox()
browser.implicitly_wait(1)

wb = openpyxl.load_workbook('Coupa_Login_Accounts.xlsx')
sh = wb.get_sheet_by_name('Sheet1')
MaxRow = sh.max_row

for x in range(1,MaxRow+1):
    print(sh.cell(row=x, column=1).value)
    Login = sh.cell(row=x, column=1).value
    Password = sh.cell(row=x, column=2).value
    browser.get('https://deloitte-ca2.coupacloud.com/session')
    browser.find_element_by_id('user_login').send_keys(Login)
    browser.find_element_by_id('user_password').send_keys(Password)
    browser.find_element_by_class_name('button').click()
    browser.get('https://deloitte-ca2.coupacloud.com/inbox/preferences')

    Requisition_Notification_Online = browser.find_element_by_xpath("//div[@class='section'][1]"
                                                                    "/div[@class='inline_form_element']"
                                                                    "/input[1]").is_selected()
    Requisition_Notification_Email = browser.find_element_by_xpath("//div[@class='section'][1]"
                                                                   "/div[@class='inline_form_element']"
                                                                   "/input[2]").is_selected()
    print(Requisition_Notification_Online)
    print(Requisition_Notification_Email)

    # ===== select the notifications online & email =====
    if Requisition_Notification_Online is False:
        browser.find_element_by_xpath("//div[@class='section'][1]/div[@class='inline_form_element']/input[1]").click()
    if Requisition_Notification_Email is False:
        browser.find_element_by_xpath("//div[@class='section'][1]/div[@class='inline_form_element']/input[2]").click()
    # ====================================================

    Requisition_Notification_Online = browser.find_element_by_xpath("//div[@class='section'][1]"
                                                                    "/div[@class='inline_form_element']"
                                                                    "/input[1]").is_selected()
    Requisition_Notification_Email = browser.find_element_by_xpath("//div[@class='section'][1]"
                                                                   "/div[@class='inline_form_element']"
                                                                   "/input[2]").is_selected()

    print(Requisition_Notification_Online)
    print(Requisition_Notification_Email)

    browser.find_element_by_xpath("//button[@class='button blue'][@type='submit']").click()
    sh.cell(row=x, column=3).value = "Configured"

    # ==== LOG OUT =========
    ActionChains(browser).move_to_element(browser.find_element_by_id('my_account')).perform()
    browser.find_element_by_xpath('//a[@href="' + "/sessions/destroy" + '"]').click()
    # ======================

browser.close()
wb.save('Excel_Output.xlsx')
