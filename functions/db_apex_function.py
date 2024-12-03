import os
import time
import math
import schedule
import configparser
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.microsoft import EdgeChromiumDriverManager


apex_files = []


class ApexDB:

    def __init__(self, username, password, files, name, time, apex_type, working_dir, awb_column):
        self.username = username
        self.password = password
        self.files = files
        self.name = name
        self.time = time
        self.apex_type = apex_type
        self.working_dir = working_dir
        self.awb_column = awb_column

    def get_awb(self):
        # Clear data
        apex_files.clear()

        # Read data
        df = pd.read_excel(self.files)
        df_length = len(df.index)
        filename = os.path.split(self.name)
        filepath = filename[0]
        desired_name = filename[1].split(".")[0]
        max_lines = 20000

        # Export to csv
        if not os.path.exists(self.working_dir):
            os.makedirs(self.working_dir)

        if df_length > 20000:
            loop = math.ceil(df_length / max_lines)
            for i in range(loop):
                export_to = rf"{filepath}/{desired_name}({i}).csv"
                data = df[(max_lines * i):(max_lines * (i + 1))]
                data.to_csv(export_to, columns=[self.awb_column],
                            index=False, header=False)
                apex_files.append(export_to)
        else:
            df.to_csv(rf"{self.name}", columns=[self.awb_column],
                      index=False, header=False)
            apex_files.append(self.name)

    def send_to_apex(self):
        # Read required information
        config = configparser.ConfigParser()
        config.read('config.ini')

        # Get required File
        self.get_awb()

        # Get ApexDB URL
        base_url = config.get('base_links', self.apex_type)

        # Configure Webdriver
        options = webdriver.EdgeOptions()
        options.use_chromium = True
        options.add_argument("start-maximized")
        options.add_argument("inprivate")

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

            for file in apex_files:
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
                upload_file.send_keys(file)

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
            return schedule.CancelJob
        except:
            driver.quit()

    def scheduleRun(self):
        # Create scheduled job to run at appointed time
        schedule.every().day.at(self.time).do(self.send_to_apex)

        # Run scheduled job if the condition meet
        while True:
            schedule.run_pending()
            if not schedule.jobs:
                break
            time.sleep(1)
