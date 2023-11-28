from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from webdriver_manager.firefox import GeckoDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from locator27 import Orange_Locator
from data27 import Orange_data
import datetime


class Orange_Hrm:

    def __init__(self):
        self.driver = webdriver.Firefox(service=Service(GeckoDriverManager().install()))
        self.wait = WebDriverWait(self.driver, 40)

    def open_url(self):
        try:
            self.driver.maximize_window()
            self.driver.get(Orange_Locator().url)
        except Exception as e:
            print(e)

    def pro_data(self, max_row):
        try:
            for row in range(2, max_row+1):
                username = org.access_data(row, 6)
                password = org.access_data(row, 7)
                
                user_info = self.wait.until(EC.element_to_be_clickable((By.NAME, Orange_Locator().user_name)))
                user_info.send_keys(username)

                pass_info = self.wait.until(EC.element_to_be_clickable((By.NAME,Orange_Locator().pass_name)))
                pass_info.send_keys(password)

                self.wait.until(EC.element_to_be_clickable((By.XPATH, Orange_Locator().log_xpath))).click()

                or_url = Orange_Locator().url
                dashboard_url = self.driver.current_url 
                dash = Orange_Locator().dash_url
                current_datetime = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                if dash in dashboard_url:
                    print("Login Success")
                    org.write_data(row, 8, "LOGIN TEST PASSED!")
                    org.write_data(row, 4, current_datetime)
                    log_out = self.wait.until(EC.element_to_be_clickable((By.XPATH, Orange_Locator().user_drop_xpath))).click()
                    log_out1 = self.wait.until(EC.element_to_be_clickable((By.XPATH, Orange_Locator().logout_xpath))).click()
                
                elif or_url in dashboard_url:
                    print("Login Failed")
                    org.write_data(row, 8, "LOGIN TEST FAILED!")
                    org.write_data(row, 4, current_datetime)
                    self.driver.refresh()

        except Exception as e:
            print(f"Error in pro_data: {e}")

        finally:
            self.driver.quit()

# Example usage:
excel_file = r"E:\Automation Testing\practice\Task\Task27\Orange_data.xlsx"
sheet_name = "Sheet1"
org = Orange_data(excel_file, sheet_name)
max_row = org.row_count()

qr = Orange_Hrm()
qr.open_url()
qr.pro_data(max_row)
