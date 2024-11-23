import sys
import time
import schedule
import configparser
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.microsoft import EdgeChromiumDriverManager


class ApexDB:
    def __init__(self, username, password, file):
        self.username = username
        self.password = password
        self.file = file

    def send_to_apex(self):
        # Read required information
        config = configparser.ConfigParser()
        config.read('config.ini')

        # Get destination URL
        base_url = config.get('base_links', 'Apex12')

        # Configure Webdriver
        options = webdriver.EdgeOptions()
        options.use_chromium = True
        options.add_argument("start-maximized")
        options.add_argument("inprivate")
        options.add_experimental_option("prefs", {
            "download.default_directory": r"C:\ERWIN\tes"
        })

        # Initialize Selenium Webdriver
        driver = webdriver.Edge(service=Service(
            EdgeChromiumDriverManager().install()
        ), options=options)
        wait = WebDriverWait(driver, 3600)

        try:
            # Access APEX DB
            driver.get(base_url)

            # Wait for Login page to load
            username_form = wait.until(
                EC.presence_of_element_located((By.ID, 'P101_USERNAME')))
            password_form = wait.until(
                EC.presence_of_element_located((By.ID, 'P101_PASSWORD')))
            login_btn = wait.until(
                EC.presence_of_element_located((By.ID, 'P101_LOGIN')))

            # Login
            username_form.send_keys(self.username)
            password_form.send_keys(self.password)
            login_btn.click()

            # Go to Report Status AWB page
            report_link = wait.until(EC.presence_of_element_located(
                (By.XPATH, '//*[@id="tabs"]/li[2]/a')))
            report_link.click()

            # Check report Type
            type_select = wait.until(EC.presence_of_element_located(
                (By.XPATH, '//*[@id="P45_PARAM_TYPE"]')))
            type_select.click()

            upload_option = wait.until(EC.presence_of_element_located(
                (By.XPATH, '//*[@id="P45_PARAM_TYPE"]/option[1]')))
            upload_option.click()

            # Upload file
            upload_file = wait.until(EC.presence_of_element_located(
                (By.ID, 'P45_BLOB_CONTENT')))
            upload_file.send_keys(self.file)

            # Process
            upload_btn = wait.until(EC.presence_of_element_located(
                (By.XPATH, '//*[@id="B49858870221701936"]')))
            upload_btn.click()

            # Wait until processed
            time.sleep(30)
            wait.until(EC.presence_of_element_located(
                (By.XPATH, '//*[@id="success-message"]')))

            # Close Webdriver
            driver.quit()
        except:
            # Close Webdriver on catch
            driver.quit()

    def runFn(self):
        # tes = ApexDB(username="ami.opr-cabang@jne.co.id",
        #              password="jne1234", file=r"C:\Users\Lenovo\Documents\debug_test5.csv")

        # tes.send_to_apex()
        # return schedule.CancelJob
        print("test")
        return schedule.CancelJob

    def scheduleRun(self):
        # schedule.every().day.at('13:25').do(runFn)
        schedule.every(3).seconds.do(self.runFn)

        while True:
            schedule.run_pending()
            time.sleep(1)
