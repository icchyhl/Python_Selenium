from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.common.action_chains import ActionChains
import openpyxl
import unittest
from unittest import runner
import traceback

# Test Automation scripts for Coupa R15

class TestResult(runner.TextTestResult):
    """
    Used to show the different test results
    """
    def addError(self, test, err):
        test.markCell('Error')
        super(TestResult, self).addError(test, err)

    def addFailure(self, test, err):
        test.markCell('Failure')
        super(TestResult, self).addFailure(test, err)

    def addSuccess(self, test):
        test.markCell('Success')
        super(TestResult, self).addSuccess(test)

class TestInvoice(unittest.TestCase):
    """
    Test Scenario for (INV3): "Receive Invoice through attachment (PDF) - Paper Scanned Invoice (PO backed Invoice)"
    """
    @classmethod
    def setUpClass(cls):
        # this will set up the initial values for this class, which the
        # test runner will do.
        cls.wb = openpyxl.load_workbook('Main_Input.xlsx')
        cls.sh_Setup = cls.wb.get_sheet_by_name('Setup')
        cls.ClientURL = cls.sh_Setup.cell(row=1, column=2).value
        cls.driver = webdriver.Firefox()
        cls.driver.get(cls.ClientURL)

    def tearDown(self):
        # This will save the progress if test scenario is aborted before full execution
        self.wb.save('Main_Output.xlsx')

    @classmethod
    def tearDownClass(cls):
        cls.driver.quit()
        cls.wb.save('Main_Output.xlsx')

    def sheetLocation(self, sheet, row, col):
        # initialize the initial value to success first
        self._sh, self._row, self._col = sheet, row, col
        self.sh = self.wb.get_sheet_by_name(self._sh)

    def markCell(self, value):
        self.sh.cell(row=self._row, column=self._col).value = value

    # ============= Beginning of the test scenarios start from below ===========

    def testLogin(self):
        """
        step "INV3.1" in the Scenario "Login to Coupa, if you are already logged in,
        Please go to home page of Coupa by clicking home icon right beneath the Company logo"
        """
        self.sheetLocation('Paper Invoice',5,8)

        driver = self.driver
        coupaLogin = self.sh_Setup.cell(row=4, column=1).value
        coupaPassword = self.sh_Setup.cell(row=4, column=2).value
        loginFieldID = 'user_login'
        passwordFieldID = 'user_password'
        loginButtonClassName = 'button'
        coupaLogoID = 'global-logo'

        loginFieldElement = WebDriverWait(driver, 10).until(lambda driver: driver.find_element_by_id(loginFieldID))
        passwordFieldElement = WebDriverWait(driver, 10).until(lambda driver: driver.find_element_by_id(passwordFieldID))
        loginButtonElement = WebDriverWait(driver, 10).until(lambda driver: driver.find_element_by_class_name(loginButtonClassName))

        loginFieldElement.clear()
        loginFieldElement.send_keys(coupaLogin)
        passwordFieldElement.clear()
        passwordFieldElement.send_keys(coupaPassword)
        loginButtonElement.click()
        WebDriverWait(driver,10).until(lambda driver: driver.find_element_by_id(coupaLogoID))

        # add an assert for the logo to appear

if __name__ == "__main__":
    unittest.main(testRunner=runner.TextTestRunner(resultclass=TestResult))
