import time
import unittest
from html_test_report import HtmlTestRunner
from selenium import webdriver
from selenium.webdriver.common.by import By
import EXCELUTIL
import openpyxl
from selenium.webdriver.common.action_chains import ActionChains


# class DemoImplicitWait():
class Testsample(unittest.TestCase):
    @classmethod
    def setUpClass(self):
        self.driver = webdriver.Chrome(executable_path="C:\seleniumwebdriver\chromedriver.exe")
        #time.sleep(3)

    def test_amazon(self):
        action = ActionChains(self.driver)
        self.driver.implicitly_wait(40)
        self.driver.get("https://www.amazon.in")
        titleOfWebPage = self.driver.title
        self.assertNotEqual("Amazon", titleOfWebPage, "PASSED")
        # seconds

        self.driver.save_screenshot(r"C:\Users\Admin\OneDrive\Desktop\screenschot\homepage1.png")
        self.driver.maximize_window()

        # capture the title of the page
        amazon = self.driver.title
        path = (r"C:\Users\Admin\OneDrive\Desktop\report.xlsx")

        # access the RowCount method
        rows = EXCELUTIL.getRowCount(path, "Sheet1")
        # perform or read the value from excel file and pass to application

        for r in range(2, rows + 1):
            Actionkeywords = EXCELUTIL.ReadData(path, "Sheet1", r, 1)
            status = EXCELUTIL.ReadData(path, "Sheet1", r, 2)

        # time.sleep(4)
        a =self.driver.find_element(by=By.XPATH, value=" //input[@id='twotabsearchtextbox']")
        action.scroll_to_element(a)
        a.send_keys("oneplus Mobile under 30000")

        # screenshot
        self.driver.save_screenshot(r"C:\Users\Admin\OneDrive\Desktop\screenschot\search MOBILE.png")
        time.sleep(4)
        self.driver.find_element(by=By.XPATH, value=" //input[@id='nav-search-submit-button']").click()

        self.driver.find_element(by=By.XPATH, value=" //span[contains(text(),'Featured')]").click()
        # time.sleep(4)
        self.driver.find_element(by=By.XPATH, value=" //a[@id='s-result-sort-select_4']").click()

        time.sleep(4)
        # screenshot
        self.driver.save_screenshot(r"C:\Users\Admin\OneDrive\Desktop\screenschot\ARRIVAL.png")
        # time.sleep(4)

        # self.driver.execute_script("arguments[0].click();",search)
        self.driver.execute_script("alert('search successfully');")
        time.sleep(4)

        self.driver.switch_to.alert.dismiss()

        workbook = openpyxl.load_workbook(r"C:\Users\Admin\OneDrive\Desktop\report.xlsx")

        # load the sheet
        sheet = workbook.active

        for r in range(1, 6):
            for c in range(1, 2):
                sheet.cell(row=r, column=c).value = "pass "

        workbook.save(r"C:\Users\Admin\OneDrive\Desktop\report.xlsx")

        print("end of file writting")

        # time.sleep(4)
    @classmethod
    def tearDown(self):
        self.driver.close()


if __name__ == "__main__":
    unittest.main(testRunner=HtmlTestRunner.HTMLTestRunner(output=r"..//C:\Users\Admin\PycharmProjects\pythonProject7\REPORT"))