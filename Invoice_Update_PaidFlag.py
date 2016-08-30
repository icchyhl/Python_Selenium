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

wb = openpyxl.load_workbook('Invoice_Update_PaidFlag_Input.xlsx')
shx = wb.get_sheet_by_name('Login')
Login = shx.cell(row=2, column=1).value
Password = shx.cell(row=2, column=2).value
clientURL = shx.cell(row=2, column = 3).value

sh = wb.get_sheet_by_name('Sheet1')
MaxRow = sh.max_row

driver = webdriver.Chrome()
driver.get(clientURL)

# >>>>>>>>>>>>>> LOGIN >>>>>>>>>>>>>>>
loginElement = WebDriverWait(driver, 10).until(lambda driver: driver.find_element_by_id('user_login'))
userPasswordElement = WebDriverWait(driver, 10).until(lambda driver: driver.find_element_by_id('user_password'))
buttonElement = WebDriverWait(driver, 10).until(lambda driver: driver.find_element_by_class_name('button'))

loginElement.send_keys(Login)
userPasswordElement.send_keys(Password)
buttonElement.click()
# <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

# >>>>>>>>>>>>>>>>>>>>> Click into invoice and update 'Paid' flag to YES >>>>>>>>>>>>>>>>>>>>
for x in range(2,MaxRow+1):
    invoiceSearchInput = sh.cell(row=x, column=1).value
    invoiceURL = shx.cell(row=2, column=4).value
    print('Invoice: ' + str(invoiceSearchInput))
    driver.get(invoiceURL) # navigate to invoice page

    invoiceSearchElement = WebDriverWait(driver, 10).until(lambda driver: driver.find_element_by_id('sf_invoice_header'))
    invoiceSearchElement.clear()
    invoiceSearchElement.send_keys(str(invoiceSearchInput) + '\n')
    time.sleep(2) # wait for invoice elements to refresh properly
    firstInvoiceElement = WebDriverWait(driver, 10).until(lambda driver: driver.find_element_by_xpath(
        "//tr[contains(@id,'invoice_header_row')]//span[@class='dt_open_link']"))
    firstInvoiceElements = WebDriverWait(driver, 10).until(lambda driver: driver.find_elements_by_xpath(
        "//tr[contains(@id,'invoice_header_row')]//span[@class='dt_open_link']"))
    print('Number of Invoices from search: ' + str(len(firstInvoiceElements)))

    if len(firstInvoiceElements) == 1: # if only 1 invoice is found form search then click into invoice
        firstInvoiceElement = WebDriverWait(driver, 10).until(lambda driver: driver.find_element_by_xpath(
            "//tr[contains(@id,'invoice_header_row')]//span[@class='dt_open_link']"))
        firstInvoiceElement.click()

        # >>>>>>>>> selecting 'paid' >>>>>>>>>>>>
        invoicePaidElement = WebDriverWait(driver, 10).until(lambda driver: driver.find_element_by_id('invoice_paid'))
        print(invoicePaidElement.is_selected())
        if invoicePaidElement.is_selected() is False:  # Update this to TRUE if wanting to change to Paid flag as no
            invoicePaidElement.click()  # checkbox for 'paid'
            savePaymentDetailsElement = WebDriverWait(driver, 10).until(lambda driver: driver.find_element_by_xpath("//a[@class=' rollover button ']//img[@class='icon icon_button sprite-disk']"))
            savePaymentDetailsElement.click()
            WebDriverWait(driver, 10).until(lambda driver: driver.find_element_by_id('pageHeader').text == 'Invoices')
            # wait for invoice page to load = previous invoice was saved
        # <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

            sh.cell(row=x, column=2).value = 'Paid checkbox configured'
            wb.save('Invoice_Update_PaidFlag_Output.xlsx')
            print(time.time() - StartTime)
        else: sh.cell(row=x, column=2).value = 'Paid checkbox was already configured'
    else: sh.cell(row=x, column=2).value = 'More than 1 invoice returned from search'
# <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

wb.save('Invoice_Update_PaidFlag_Output.xlsx')
driver.quit()
print(time.time() - StartTime)













