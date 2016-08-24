from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.common.action_chains import ActionChains
import openpyxl
import unittest
import traceback

# Test Automation scripts for Coupa R15
wb = openpyxl.load_workbook('Main_Input.xlsx')
sh_Setup = wb.get_sheet_by_name('Setup')
ClientURL = sh_Setup.cell(row=1,column=2).value

class TestInvoice(unittest.TestCase):
    """
    Test Scenario for (INV3): "Receive Invoice through attachment (PDF) - Paper Scanned Invoice (PO backed Invoice)"
    """
    def setUp(self):
        self.driver = webdriver.Firefox()
        self.driver.get(ClientURL)
        self.wb = wb
        self.sh = wb.get_sheet_by_name('Paper Invoice')
        self.sh_Setup = sh_Setup

    def testLogin(self):
        """
        step "INV3.1" in the Scenario "Login to Coupa, if you are already logged in,
        Please go to home page of Coupa by clicking home icon right beneath the Company logo"
        """
        try:
            driver = self.driver
            coupaLogin           = sh_Setup.cell(row=4, column=1).value
            coupaPassword        = sh_Setup.cell(row=4, column=2).value
            loginFieldID         = 'user_login'
            passwordFieldID      = 'user_password'
            loginButtonClassName = 'button'
            coupaLogoID          = 'global-logo'

            loginFieldElement    = WebDriverWait(driver, 10).until(lambda driver: driver.find_element_by_id(loginFieldID))
            passwordFieldElement = WebDriverWait(driver, 10).until(lambda driver: driver.find_element_by_id(passwordFieldID))
            loginButtonElement   = WebDriverWait(driver, 10).until(lambda driver: driver.find_element_by_class_name(loginButtonClassName))

            loginFieldElement.clear()
            loginFieldElement.send_keys(coupaLogin)
            passwordFieldElement.clear()
            passwordFieldElement.send_keys(coupaPassword)
            loginButtonElement.click()
            WebDriverWait(driver,10).until(lambda driver: driver.find_element_by_id(coupaLogoID))
        except:
            traceback.print_exc()
            self.sh.cell(row=5,column=8).value = 'Fail'
        else:
            self.sh.cell(row=5, column=8).value = 'Pass'

    def tearDown(self):
        self.driver.quit()
        self.wb.save('Main_Output.xlsx')

if __name__ == "__main__":
    unittest.main()


