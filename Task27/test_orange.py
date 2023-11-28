from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from webdriver_manager.firefox import GeckoDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from locator27 import Orange_Locator
from data27 import Orange_data
import datetime

class Orange_Hrm:

    def __init__(self, org_instance):
        self.driver = webdriver.Firefox(service=Service(GeckoDriverManager().install()))
        self.wait = WebDriverWait(self.driver, 40)
        self.org_instance = org_instance
        self.login_successful = False

    def open_url(self):
        try:
            self.driver.maximize_window()
            self.driver.get(Orange_Locator().url)
        except Exception as e:
            print(e)

    def pro_data(self, max_row):
        try:
            for row in range(2, max_row+1):
                username = self.org_instance.access_data(row, 6)
                password = self.org_instance.access_data(row, 7)

                user_info = self.wait.until(EC.element_to_be_clickable((By.NAME, Orange_Locator().user_name)))
                user_info.send_keys(username)

                pass_info = self.wait.until(EC.element_to_be_clickable((By.NAME, Orange_Locator().pass_name)))
                pass_info.send_keys(password)

                self.wait.until(EC.element_to_be_clickable((By.XPATH, Orange_Locator().log_xpath))).click()

                or_url = Orange_Locator().url
                dashboard_url = self.driver.current_url
                dash = Orange_Locator().dash_url
                current_datetime = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                if dash in dashboard_url:
                    print("Login Success")
                    self.login_successful = True
                    self.org_instance.write_data(row, 8, "LOGIN TEST PASSED!")
                    self.org_instance.write_data(row, 4, current_datetime)
                    log_out = self.wait.until(
                        EC.element_to_be_clickable((By.XPATH, Orange_Locator().user_drop_xpath))).click()
                    log_out1 = self.wait.until(
                        EC.element_to_be_clickable((By.XPATH, Orange_Locator().logout_xpath))).click()

                elif or_url in dashboard_url:
                    print("Login Failed")
                    self.login_successful = False
                    self.org_instance.write_data(row, 8, "LOGIN TEST FAILED!")
                    self.org_instance.write_data(row, 4, current_datetime)
                    self.driver.refresh()

        except Exception as e:
            print(f"Error in pro_data: {e}")

        finally:
            self.quit_driver()

    def is_login_successful(self):
        # Check if the logout button is present, indicating a successful login
        try:
            logout_button = self.wait.until(EC.element_to_be_clickable((By.XPATH, Orange_Locator().logout_xpath)))
            return logout_button is not None
        except Exception:
            # If timeout exception occurs, the logout button is not present, indicating a failed login
            return False

    def quit_driver(self):
        self.driver.quit()


# Example usage:
excel_file = r"E:\Automation Testing\practice\Task\Task27\Orange_data.xlsx"
sheet_name = "Sheet1"
org = Orange_data(excel_file, sheet_name)
max_row = org.row_count()

# Function to yield Orange_Hrm instance for each set of login credentials
def orange_hrm_instance_generator():
    for username, password in [
        ("Shelby", "cappsey234"),
        ("Admin", "admin123"),
        ("Carloton", "esseley678"),
        ("Jamey", "martin908")
    ]:
        yield Orange_Hrm(org), username, password


# Pytest function to test login for each set of credentials
import pytest

@pytest.mark.parametrize("orange_hrm_instance, username, password", orange_hrm_instance_generator())
def test_login(orange_hrm_instance, username, password):
    orange_hrm_instance.open_url()
    orange_hrm_instance.pro_data(1)  # Assuming 1 as the maximum row for this example
    assert orange_hrm_instance.is_login_successful()

    # Get the expected information in the URL after successful login based on the username
    expected_info_in_url = get_expected_info_in_url(username)

    # Check if the expected information is present in the current URL
    assert expected_info_in_url in orange_hrm_instance.driver.current_url.lower(), f"Unexpected URL after login for {username}"

def get_expected_info_in_url(username):
    return f"dashboard/index"
