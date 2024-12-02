import os
import time
import configparser
import tkinter as tk
from selenium import webdriver
from tkinter.messagebox import showinfo
from datetime import timedelta, datetime
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from functions.general_function import combine_files
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from functions.db_apex_function import ApexDB


class DB141:
    def __init__(self, mode, username_141, password_141, username_apex, password_apex, date_from, date_thru, loop, is_combine, is_apex, working_dir, apex_type, apex_file_name, penarikan, log):
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
        self.log = log

    def outbound_data141(self):
        # Read app configuration
        config = configparser.ConfigParser()
        config.read('config.ini')
        base_link = config.get('base_links', '141')

        # Initialize webdriver
        download_dir = rf"{self.working_dir}\{self.penarikan}"
        options = webdriver.EdgeOptions()
        options.use_chromium = True
        options.add_argument("start-maximized")
        options.add_argument("inprivate")
        options.add_experimental_option("prefs", {
            "download.default_directory": download_dir})
        driver = webdriver.Edge(service=Service(
            EdgeChromiumDriverManager().install()), options=options)
        wait = WebDriverWait(driver, 3600)

        # Read first date, Ascending order
        used_date = self.date_from

        # List of downloaded file
        saved_files = []

        # Access Dashboard
        driver.get(base_link)
        self.log.insert(
            tk.END, f"{datetime.now().strftime('%H:%M')} - Download Outbound data from DB 141  \n")
        self.log.see("end")

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

            self.log.insert(
                tk.END, f"{datetime.now().strftime('%H:%M')} - Logged in as {self.username_db} \n")
            self.log.see("end")

            # Get current page information
            current_url = driver.current_url
            url_parts = current_url.split(":")

            current_page = url_parts[2]
            session = url_parts[3]

            # Outbound Data Page URL
            outbound_page = f"{base_link}:54:{session}::NO:RIR,54:P0_REP_FLAG:1"

            # Download data
            for i in range(0, self.loop):
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

                self.log.insert(
                    tk.END, f"{datetime.now().strftime('%H:%M')} - Downloading data {used_date.strftime('%d-%b-%Y')} \n")
                self.log.see("end")

                # Wait for file to be downloaded
                while not os.path.isfile(rf'{download_dir}\detail.csv'):
                    time.sleep(15)

                # Rename downloaded file
                result_filename = rf'{download_dir}\{used_date.strftime("%d-%b-%Y")}.csv'
                if os.path.isfile(rf'{download_dir}\detail.csv'):
                    os.rename(rf'{download_dir}\detail.csv', result_filename)
                    saved_files.append(result_filename)

                # Change date
                used_date += timedelta(days=1)

            # Close webdriver
            driver.quit()

            # Combine file if required
            combined_file = ""
            if self.is_combine == True:
                self.log.insert(
                    tk.END, f"{datetime.now().strftime('%H:%M')} - Combining data... \n")
                self.log.see("end")
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
                                 awb_column="CNOTE NO")
                apex_fn.send_to_apex()
                self.log.insert(
                    tk.END, f"{datetime.now().strftime('%H:%M')} - Uploading data to ApexDB... \n")
                self.log.see("end")

            self.log.insert(
                tk.END, f"{datetime.now().strftime('%H:%M')} - Process finished \n")
            self.log.see("end")
            showinfo(title="Message", message="Proses Selesai")
        except Exception as e:
            driver.quit()
            self.log.insert(
                tk.END, f"{datetime.now().strftime('%H:%M')} - Process failed \n")
            self.log.insert(
                tk.END, f"{datetime.now().strftime('%H:%M')} - Reason: {e} \n")
            self.log.see("end")
            showinfo(title="Error", message=e)

    def inbound_data141(self):
        # Read app configuration
        config = configparser.ConfigParser()
        config.read('config.ini')
        base_link = config.get('base_links', '141')

        # Initialize webdriver
        download_dir = rf"{self.working_dir}\{self.penarikan}"
        options = webdriver.EdgeOptions()
        options.use_chromium = True
        options.add_argument("start-maximized")
        options.add_argument("inprivate")
        options.add_experimental_option("prefs", {
            "download.default_directory": download_dir})
        driver = webdriver.Edge(service=Service(
            EdgeChromiumDriverManager().install()), options=options)
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
        self.log.insert(
            tk.END, f"{datetime.now().strftime('%H:%M')} - Download {process[self.mode]} data from DB 141  \n")
        self.log.see("end")

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

            self.log.insert(
                tk.END, f"{datetime.now().strftime('%H:%M')} - Logged in as {self.username_db} \n")
            self.log.see("end")

            # Read current page information
            current_url = driver.current_url
            url_parts = current_url.split(":")

            current_page = url_parts[2]
            session = url_parts[3]

            page_code = [
                # [Page Code, End Page Code, Result Code]
                ['1', '1', '64'],       # 0 - End-To-End
                ['24', '24', '64'],     # 1 - Summary
                ['72', '1', '75'],      # 2 - Intracity
            ]

            # Inbound Data Page URL
            inbound_page = f"{base_link}:{page_code[self.mode][0]}:{session}::NO:RIR,{page_code[self.mode][0]}:P0_REP_FLAG:{page_code[self.mode][1]}"

            for i in range(0, self.loop):
                # Go to Inbound Data Page
                driver.get(inbound_page)

                # Change data type
                if self.mode == 2:
                    wait.until(EC.presence_of_element_located(
                        (By.NAME, 'P0_REP_DATE')))
                    option_input = driver.find_element(By.NAME, 'P0_REP_DATE')
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

                # Intracity data doesn't need to be generated
                if self.mode != 2:
                    # Wait data to be generated
                    time.sleep(10)
                    wait.until(EC.invisibility_of_element_located(
                        (By.XPATH,
                         '//*[@id="report_R29576385550981813706"]/div/div[1]/table/tbody/tr/td[2]/button/img')
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

                # Start download proccess
                download_page = f"{base_link}:{page_code[self.mode][2]}:{session}:CSV::::"
                driver.get(download_page)

                self.log.insert(
                    tk.END, f"{datetime.now().strftime('%H:%M')} - Downloading data {used_date.strftime('%d-%b-%Y')} \n")
                self.log.see("end")

                # Wait for file to be downloaded
                while not os.path.isfile(rf'{download_dir}\detail.csv'):
                    time.sleep(15)

                # Rename downloaded file
                if os.path.isfile(rf'{download_dir}\detail.csv'):
                    os.rename(rf'{download_dir}\detail.csv',
                              rf'{download_dir}\{used_date.strftime("%d-%b-%Y")}.csv')
                    saved_files.append(
                        rf'{download_dir}\{used_date.strftime("%d-%b-%Y")}.csv')

                # Change date
                used_date += timedelta(days=1)

            # Close webdriver
            driver.quit()

            # Combine data if required
            combined_file = ""
            if self.is_combine == True:
                self.log.insert(
                    tk.END, f"{datetime.now().strftime('%H:%M')} - Combining data... \n")
                self.log.see("end")
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
                                 awb_column="CNOTE NO")
                apex_fn.get_awb()
                self.log.insert(
                    tk.END, f"{datetime.now().strftime('%H:%M')} - Uploading data to ApexDB... \n")
                self.log.see("end")

            self.log.insert(
                tk.END, f"{datetime.now().strftime('%H:%M')} - Process finished \n")
            self.log.see("end")
            showinfo(title="Message", message="Proses Selesai")
        except Exception as e:
            driver.quit()
            self.log.insert(
                tk.END, f"{datetime.now().strftime('%H:%M')} - Process failed \n")
            self.log.insert(
                tk.END, f"{datetime.now().strftime('%H:%M')} - Reason: {e} \n")
            self.log.see("end")
            showinfo(title="Error", message=e)
