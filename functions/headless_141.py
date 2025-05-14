import os
import time
import configparser
from selenium import webdriver
from datetime import timedelta
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from functions.db_apex_function import ApexDB
from functions.general_function import combine_files, check_file


class DB141:
    def __init__(self, mode, username_141, password_141, username_apex, password_apex, date_from, date_thru, loop, is_combine, is_apex, working_dir, apex_type, apex_file_name, penarikan, driver):
        self.mode = mode
        self.username_db = username_141
        self.password_db = password_141
        self.username_apex = username_apex
        self.password_apex = password_apex
        self.date_from = date_from
        self.date_thru = date_thru
        self.loop = loop
        self.is_combine = is_combine
        self.is_apex = is_apex
        self.working_dir = working_dir
        self.apex_type = apex_type
        self.apex_name = apex_file_name
        self.penarikan = penarikan
        self.driver = driver

    def outbound_data141(self):
        # Read app configuration
        config = configparser.ConfigParser()
        config.read('config.ini')
        base_link = config.get('base_links', '141')

        # Initialize webdriver
        download_dir = rf"{self.working_dir}\{self.penarikan}"
        options = webdriver.ChromeOptions()
        options.use_chromium = True
        options.add_argument("--headless=new")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_experimental_option("prefs", {
            "download.default_directory": download_dir})
        driver = webdriver.Chrome(
            service=Service(self.driver), options=options)
        wait = WebDriverWait(driver, 3600)

        # Read first date, Ascending order
        used_date = self.date_from

        # List of downloaded file
        saved_files = []

        # Access Dashboard
        driver.get(base_link)

        try:
            # Login
            username_form = wait.until(
                EC.presence_of_element_located((By.ID, 'P101_USERNAME')))

            password_form = wait.until(
                EC.presence_of_element_located((By.ID, 'P101_PASSWORD')))

            login_button = wait.until(
                EC.presence_of_element_located((By.ID, 'B12943410900979389915')))

            username_form.send_keys(self.username_db)
            password_form.send_keys(self.password_db)
            login_button.click()

            # Get current page information
            current_url = driver.current_url
            url_parts = current_url.split(":")

            current_page = url_parts[2]
            session = url_parts[3]

            # Outbound Data Page URL
            outbound_page = f"{base_link}:54:{session}::NO:RIR,54:P0_REP_FLAG:1"

            # Download data
            for i in range(0, self.loop):
                is_verified = False

                # Verifiy datas date before going into next iteration
                while is_verified == False:
                    # Go to Outbound Data Page
                    driver.get(outbound_page)

                    # Clear Date Input Form
                    wait.until(
                        EC.presence_of_element_located((By.NAME, 'P0_DATE1')))
                    date_from_input = driver.find_element(By.NAME, 'P0_DATE1')
                    date_from_input.clear()
                    wait.until(
                        EC.presence_of_element_located((By.NAME, 'P0_DATE2')))
                    date_thru_input = driver.find_element(By.NAME, 'P0_DATE2')
                    date_thru_input.clear()

                    # Input Date
                    date_from_input.send_keys(used_date.strftime('%d-%b-%Y'))
                    date_thru_input.send_keys(used_date.strftime('%d-%b-%Y'))

                    # Proceed
                    go_btn = wait.until(EC.presence_of_element_located(
                        (By.ID, 'B12943525354637205515')))
                    go_btn.click()

                    # Wait until loading done
                    time.sleep(10)
                    wait.until(EC.invisibility_of_element_located(
                        (By.XPATH, '//*[@id="loadingIcon"]')
                    ))

                    # Wait data to be generated
                    time.sleep(10)
                    wait.until(EC.invisibility_of_element_located(
                        (By.XPATH,
                         '//*[@id="report_R29576385550981813706"]/div/div[1]/table/tbody/tr/td[2]/button/img')
                    ))

                    # Refresh page after data generated
                    time.sleep(5)
                    driver.refresh()

                    # Go to result page
                    result_link = wait.until(EC.presence_of_element_located(
                        (By.XPATH, '//*[@id="14532092005068365313_orig"]/tbody/tr[2]/td[1]/a')))
                    result_link.click()

                    wait.until(EC.presence_of_element_located(
                        (By.XPATH, '//*[@id="C5777933338987697868"]/a')))

                    # Start download proccess
                    download_page = f"{base_link}:95:{session}:CSV::::"
                    driver.get(download_page)

                    # Wait for file to be downloaded
                    while not os.path.isfile(rf'{download_dir}\detail.csv'):
                        time.sleep(15)

                    # Rename downloaded file
                    default_filepath = rf'{download_dir}\detail.csv'
                    if os.path.isfile(default_filepath):
                        # Rename file
                        renamed_file = rf'{download_dir}\{used_date.strftime("%d-%b-%Y")}.csv'
                        os.rename(default_filepath, renamed_file)
                        # Verifiy date inside file

                        is_verified = check_file(
                            file=renamed_file, current_date=used_date.date(), column_name='CNOTE DATE')

                        # Append file to list of downloaded file
                        if is_verified:
                            saved_files.append(renamed_file)

                # Change date and proceed to next iterarion
                used_date -= timedelta(days=1)

            # Close webdriver
            driver.quit()

            # Combine file if required
            combined_file = ""
            if self.is_combine == True:
                combined_file = combine_files(files=saved_files,
                                              start_date=self.date_from, end_date=self.date_thru, is_standalone=False)

            # Upload to apex
            if self.is_apex == True:
                apex_fn = ApexDB(username=self.username_apex,
                                 password=self.password_apex,
                                 files=combined_file,
                                 name=self.apex_name,
                                 time="",
                                 apex_type=self.apex_type,
                                 working_dir=self.working_dir,
                                 awb_column="CNOTE NO",
                                 driver=self.driver)
                apex_fn.send_to_apex()
        except Exception as e:
            driver.quit()

    def inbound_data141(self):
        # Read app configuration
        config = configparser.ConfigParser()
        config.read('config.ini')
        base_link = config.get('base_links', '141')

        # Initialize webdriver
        download_dir = rf"{self.working_dir}\{self.penarikan}"
        options = webdriver.ChromeOptions()
        options.use_chromium = True
        options.add_argument("--headless=new")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_experimental_option("prefs", {
            "download.default_directory": download_dir})
        driver = webdriver.Edge(service=Service(self.driver), options=options)
        wait = WebDriverWait(driver, 3600)

        # List of downloaded file
        saved_files = []

        # Read first date, Ascending order
        used_date = self.date_from

        # Access DB
        process = ['Inbound End-to-End',
                   'Summary Inbound',
                   'Intracity']
        driver.get(base_link)

        try:
            # Login
            username_form = wait.until(
                EC.presence_of_element_located((By.ID, 'P101_USERNAME')))
            password_form = wait.until(
                EC.presence_of_element_located((By.ID, 'P101_PASSWORD')))
            login_button = wait.until(
                EC.presence_of_element_located((By.ID, 'B12943410900979389915')))

            username_form.send_keys(self.username_db)
            password_form.send_keys(self.password_db)

            login_button.click()

            # Read current page information
            current_url = driver.current_url
            url_parts = current_url.split(":")

            current_page = url_parts[2]
            session = url_parts[3]

            page_code = [
                # [Page Code, End Page Code, Result Code]
                ['1', '1', '120'],       # 0 - End-To-End
                ['24', '24', '64'],     # 1 - Summary
                ['72', '1', '75'],      # 2 - Intracity
            ]

            # Inbound Data Page URL
            inbound_page = f"{base_link}:{page_code[self.mode][0]}:{session}::NO:RIR,{page_code[self.mode][0]}:P0_REP_FLAG:{page_code[self.mode][1]}"

            for i in range(0, self.loop):
                is_verified = False

                # Verifiy datas date before going into next iteration
                while is_verified == False:
                    # Go to Inbound Data Page
                    driver.get(inbound_page)

                    # Change data type
                    if self.mode == 2:
                        wait.until(EC.presence_of_element_located(
                            (By.NAME, 'P0_REP_DATE')))
                        option_input = driver.find_element(
                            By.NAME, 'P0_REP_DATE')
                        option_input.send_keys('AWB')

                    # Clear date input form
                    wait.until(
                        EC.presence_of_element_located((By.NAME, 'P0_DATE1')))
                    date_from_input = driver.find_element(By.NAME, 'P0_DATE1')
                    date_from_input.clear()
                    wait.until(
                        EC.presence_of_element_located((By.NAME, 'P0_DATE2')))
                    date_thru_input = driver.find_element(By.NAME, 'P0_DATE2')
                    date_thru_input.clear()

                    # Input date
                    date_from_input.send_keys(used_date.strftime('%d-%b-%Y'))
                    date_thru_input.send_keys(used_date.strftime('%d-%b-%Y'))

                    # Proceed
                    go_btn = wait.until(EC.presence_of_element_located(
                        (By.ID, 'B12943525354637205515')))
                    go_btn.click()

                    # Wait until loading finished
                    time.sleep(10)
                    wait.until(EC.invisibility_of_element_located(
                        (By.XPATH, '//*[@id="loadingIcon"]')
                    ))
                    time.sleep(10)

                    # Intracity data doesn't need to be generated
                    if self.mode < 2:
                        # Wait data to be generated
                        loading_img = [
                            '//*[@id="report_R37362249766539564772"]/div/div[1]/table/tbody/tr/td[2]/button/img',
                            '//*[@id="report_R29705817846763909733"]/div/div[1]/table/tbody/tr/td[2]/button/img'
                        ]
                        wait.until(EC.invisibility_of_element_located(
                            (By.XPATH, loading_img[self.mode])
                        ))

                        # Refresh page after data generated
                        time.sleep(5)
                        driver.refresh()

                    # Link to Result Page
                    result_xpath = [
                        '//*[@id="12887203838359850643_orig"]/tbody/tr[2]/td[3]/a',
                        '//*[@id="12913870024091317880_orig"]/tbody/tr[2]/td[3]/a',
                        '//*[@id="12859194002749946846_orig"]/tbody/tr[2]/td[1]/a',
                    ]

                    result_page = wait.until(EC.presence_of_element_located(
                        (By.XPATH, result_xpath[self.mode])))

                    # Go to Result Page
                    result_page.click()
                    time.sleep(5)

                    # Start download proccess
                    download_page = f"{base_link}:{page_code[self.mode][2]}:{session}:CSV::::"
                    driver.get(download_page)

                    # Wait for file to be downloaded
                    while not os.path.isfile(rf'{download_dir}\detail.csv'):
                        time.sleep(15)

                    # Rename downloaded file
                    default_filepath = rf'{download_dir}\detail.csv'
                    if os.path.isfile(default_filepath):
                        # Rename file
                        renamed_file = rf'{download_dir}\{used_date.strftime("%d-%b-%Y")}.csv'
                        os.rename(default_filepath, renamed_file)

                        # Verifiy date inside file
                        column_header = [
                            'MANIFEST INBOUND DATE', 'MANIFEST INBOUND DATE', 'CNOTE DATE']
                        is_verified = check_file(
                            file=renamed_file, current_date=used_date.date(), column_name=column_header[self.mode])

                        # Append file to list of downloaded file
                        if is_verified:
                            saved_files.append(renamed_file)

                # Change date and go to next iteration
                used_date += timedelta(days=1)

            # Close webdriver
            driver.quit()

            # Combine data if required
            combined_file = ""
            if self.is_combine == True:
                combined_file = combine_files(files=saved_files,
                                              start_date=self.date_from, end_date=self.date_thru, is_standalone=False)

            # Upload to apex
            if self.is_apex == True:
                apex_fn = ApexDB(username=self.username_apex,
                                 password=self.password_apex,
                                 files=combined_file,
                                 name=self.apex_name,
                                 time="",
                                 apex_type=self.apex_type,
                                 working_dir=self.working_dir,
                                 awb_column="AWB NO",
                                 driver=self.driver)
                apex_fn.send_to_apex()
        except Exception as e:
            driver.quit()
