import os
import time
import configparser
import tkinter as tk
from selenium import webdriver
from tkinter.messagebox import showinfo
from datetime import timedelta, datetime
from selenium.webdriver.common.by import By
from functions.db_apex_function import ApexDB
from selenium.webdriver.edge.service import Service
from functions.general_function import combine_files, check_file
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.microsoft import EdgeChromiumDriverManager


class DB117:
    def __init__(self, username_117, password_117, username_apex, password_apex, date_from, date_thru, loop, is_combine, is_apex, working_dir, apex_type, apex_file_name, penarikan, log):
        self.username_db = username_117
        self.password_db = password_117
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

    def inbound_data117(self):
        # Read app configuration
        config = configparser.ConfigParser()
        config.read('config.ini')
        base_link = config.get('base_links', '117')

        # Initialize webdriver
        options = webdriver.EdgeOptions()
        options.use_chromium = True
        options.add_argument("start-maximized")
        options.add_argument("inprivate")
        options.add_experimental_option("prefs", {
            "download.default_directory": rf"{self.working_dir}\{self.penarikan}"})
        driver = webdriver.Edge(service=Service(
            EdgeChromiumDriverManager().install()), options=options)
        wait = WebDriverWait(driver, 3600)

        # List of downloaded files
        saved_files = []

        try:
            # Access Dashboard
            driver.get(base_link)
            self.log.insert(
                tk.END, f"{datetime.now().strftime('%H:%M')} - Download Inbound data from DB 117  \n")
            self.log.see("end")

            # Read first date, Ascending order
            used_date = self.date_from

            time.sleep(5)

            # Login
            username_form = wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, '#P1011_USERNAME')))
            password_form = wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, '#P1011_PASSWORD')))
            login_button = wait.until(
                EC.presence_of_element_located((By.XPATH, '/html/body/form/div[2]/main/div/div/div/div/div[3]/button')))

            username_form.send_keys(self.username_db)
            password_form.send_keys(self.password_db)
            login_button.click()

            # Check if already logged in
            wait.until(EC.presence_of_element_located(
                (By.XPATH, '/html/body/form/header/div[1]/div[2]')))
            time.sleep(5)

            self.log.insert(
                tk.END, f"{datetime.now().strftime('%H:%M')}  - Logged in as {self.username_db} \n")
            self.log.see("end")

            # Read current page information
            current_url = driver.current_url
            url_parts = current_url.split("=")
            user_session = url_parts[1]

            # Inbound Data Page URL
            inbound_page = f"{base_link}/56?session={user_session}"

            for i in range(0, self.loop):
                is_verified = False

                # Verifiy data date before going into next iteration
                while is_verified == False:
                    # Go to Inbound Data Page
                    driver.get(inbound_page)

                    # Select Regional Dropdown
                    regional_btn = wait.until(EC.presence_of_element_located(
                        (By.XPATH, '//*[@id="P56_REGIONAL_DEST_lov_btn"]')))
                    regional_btn.click()

                    jtbnn_link = wait.until(EC.presence_of_element_located(
                        (By.XPATH, '//*[@id="PopupLov_56_P56_REGIONAL_DEST_dlg"]/div[2]/div/div[3]/ul/li')))
                    jtbnn_link.click()

                    # Select Branch Dropdown
                    branch_btn = wait.until(EC.presence_of_element_located(
                        (By.XPATH, '//*[@id="P56_BRANCH_DEST_lov_btn"]')))
                    branch_btn.click()

                    branch_ami_link = wait.until(EC.presence_of_element_located(
                        (By.XPATH, '//*[@id="PopupLov_56_P56_BRANCH_DEST_dlg"]/div[2]/div/div[3]/ul/li')))
                    branch_ami_link.click()

                    # Clear date input form
                    date_from_input = wait.until(EC.presence_of_element_located(
                        (By.XPATH, '//*[@id="P56_DATE1"]')))
                    date_from_input.clear()

                    date_thru_input = wait.until(EC.presence_of_element_located(
                        (By.XPATH, '//*[@id="P56_DATE2"]')))
                    date_thru_input.clear()

                    # Input date
                    date_from_input.send_keys(used_date.strftime('%d-%b-%Y'))
                    date_thru_input.send_keys(used_date.strftime('%d-%b-%Y'))

                    # Proceed
                    go_btn = wait.until(EC.presence_of_element_located(
                        (By.XPATH, '/html/body/form/div[1]/div[2]/div[2]/main/div[2]/div/div[2]/div/div/div[2]/div[2]/div[2]/div/div/div/div/div/div[1]/div/div/div[2]/div[2]/div/div[1]/div/div/div[1]/div/button')))
                    go_btn.click()

                    # Wait data to be generated
                    time.sleep(10)
                    wait.until(EC.invisibility_of_element_located(
                        (By.XPATH,
                         '/html/body/form/div[1]/div[2]/div[2]/main/div[2]/div/div[2]/div/div/div[2]/div[2]/div[2]/div/div/div/div/div/div[1]/div/div/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div[2]/div/div/div[1]/table/tbody/tr/td[2]/button')
                    ))

                    # Refresh page after data generated
                    time.sleep(5)
                    driver.refresh()

                    # Go into HAWB Page
                    hawb_link = wait.until(EC.presence_of_element_located(
                        (By.XPATH, '//*[@id="report_table_INB"]/tbody/tr/td[6]/a')))
                    hawb_link.click()

                    # Remove filter
                    filter_checkbox = wait.until(EC.presence_of_element_located(
                        (By.XPATH, '/html/body/form/div[1]/div[2]/div[2]/main/div[2]/div/div[2]/div/div/div[2]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/ul/li/span[1]/input')))
                    filter_checkbox.click()
                    time.sleep(5)

                    # Wait table data to be displayed
                    wait.until(EC.presence_of_element_located(
                        (By.XPATH, '/html/body/form/div[1]/div[2]/div[2]/main/div[2]/div/div[2]/div/div/div[2]/div[2]/div[2]/div/div[2]/div[2]/div[6]/div[1]/div/div[1]/table/tr/th[1]/a')))

                    # Open Action Dropdown
                    inbound_action_btn = wait.until(
                        EC.presence_of_element_located((By.ID, 'INB_actions_button')))
                    inbound_action_btn.click()

                    # Open download menu
                    download_menu = wait.until(EC.presence_of_element_located(
                        (By.XPATH, '/html/body/div[3]/div/ul/li[14]/div/span[1]/button')))
                    download_menu.click()

                    # Select Excel data format
                    excel_btn = wait.until(EC.presence_of_element_located(
                        (By.XPATH, '/html/body/div[5]/div[2]/div/ul/li[3]')))
                    excel_btn.click()

                    # Proceed to download
                    download_btn = wait.until(EC.presence_of_element_located(
                        (By.XPATH, '/html/body/div[5]/div[3]/div/button[2]')))
                    download_btn.click()

                    # Record download timestamp
                    download_time = datetime.now()
                    self.log.insert(
                        tk.END, f"{datetime.now().strftime('%H:%M')}  - Downloading data {used_date.strftime('%d-%b-%Y')} \n")
                    self.log.see("end")

                    # Wait for file to be downloaded, if in 5 minutes files not detected try to download again
                    while not os.path.isfile(rf"{self.working_dir}\{self.penarikan}\Inbound .xlsx"):
                        time.sleep(15)
                        if datetime.now() - download_time > timedelta(minutes=5):
                            download_btn.click()
                            download_time = datetime.now()

                    # Rename downloaded file
                    default_filepath = rf"{self.working_dir}\{self.penarikan}\Inbound .xlsx"
                    if os.path.isfile(default_filepath):
                        # Rename file
                        renamed_file = rf"{self.working_dir}\{self.penarikan}\{used_date.strftime('%d-%b-%Y')}.xlsx"
                        os.rename(default_filepath, renamed_file)

                        # Verifiy data date
                        is_verified = check_file(
                            file=renamed_file, current_date=used_date.date(), column_name='Manifest Inbound Date')

                        # Append file to list of downloaded file
                        if is_verified:
                            saved_files.append(renamed_file)

                # Change date
                used_date += timedelta(days=1)

            # Close webdriver
            driver.quit()

            # Combine data if required
            combined_file = ""
            if self.is_combine == True:
                self.log.insert(
                    tk.END, f"{datetime.now().strftime('%H:%M')}  - Combining data... \n")
                self.log.see("end")
                combine_files(files=saved_files,
                              start_date=self.date_from, end_date=self.date_thru, is_standalone=False)

            # Upload to apex
            if self.is_apex == True:
                self.log.insert(
                    tk.END, f"{datetime.now().strftime('%H:%M')} - Uploading data to ApexDB... \n")
                self.log.see("end")
                apex_fn = ApexDB(username=self.username_apex,
                                 password=self.password_apex,
                                 files=combined_file,
                                 name=self.apex_name,
                                 time="",
                                 apex_type=self.apex_type,
                                 working_dir=self.working_dir,
                                 awb_column="Hawb No")
                apex_fn.send_to_apex()

            self.log.insert(
                tk.END, f"{datetime.now().strftime('%H:%M')}  - Process finished \n")
            self.log.see("end")
            showinfo(title="Message", message="Proses Selesai")
        except Exception as e:
            driver.quit()
            self.log.insert(
                tk.END, f"{datetime.now().strftime('%H:%M')}  - Process failed \n")
            self.log.insert(
                tk.END, f"{datetime.now().strftime('%H:%M')}  - Reason: {e} \n")
            self.log.see("end")
            showinfo(title="Error", message=f'{e}')

    def outbound_data117(self):
        # Read app configuration
        config = configparser.ConfigParser()
        config.read('config.ini')
        base_link = config.get('base_links', '117')

        # Initialize webdriver
        options = webdriver.EdgeOptions()
        options.use_chromium = True
        options.add_argument("start-maximized")
        options.add_argument("inprivate")
        options.add_experimental_option("prefs", {
            "download.default_directory": rf"{self.working_dir}\{self.penarikan}"})
        driver = webdriver.Edge(service=Service(
            EdgeChromiumDriverManager().install()), options=options)
        wait = WebDriverWait(driver, 3600)

        # List of downloaded files
        saved_files = []

        try:
            # Access dashboard
            driver.get(base_link)
            self.log.insert(
                tk.END, f"{datetime.now().strftime('%H:%M')} - Download Outbound data from DB 117  \n")
            self.log.see("end")

            # Read first date, Ascending order
            used_date = self.date_from

            # There's a particular reason for this
            time.sleep(5)

            # Login
            username_form = wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, '#P1011_USERNAME')))
            password_form = wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, '#P1011_PASSWORD')))
            login_button = wait.until(
                EC.presence_of_element_located((By.XPATH, '/html/body/form/div[2]/main/div/div/div/div/div[3]/button')))

            username_form.send_keys(self.username_db)
            password_form.send_keys(self.password_db)
            login_button.click()

            # Check if already logged in
            wait.until(EC.presence_of_element_located(
                (By.XPATH, '/html/body/form/header/div[1]/div[2]')))
            time.sleep(5)

            self.log.insert(
                tk.END, f"{datetime.now().strftime('%H:%M')}  - Logged in as {self.username_db} \n")
            self.log.see("end")

            # Read current page Information
            current_url = driver.current_url
            url_parts = current_url.split("=")
            user_session = url_parts[1]

            # Outbound Data Page URL
            outbound_page = f"{base_link}/53?session={user_session}"

            for i in range(0, self.loop):
                is_verified = False

                # Verify datas date before proceeding to next iteration
                while is_verified == False:
                    # Go to Outbound Data Page
                    driver.get(outbound_page)

                    # Select Regional dropdown
                    regional_dropdown = wait.until(EC.presence_of_element_located(
                        (By.XPATH, '//*[@id="P53_REGIONAL"]')))
                    regional_dropdown.send_keys('JTBNN')
                    wait.until(EC.text_to_be_present_in_element(
                        (By.XPATH, '//*[@id="P53_REGIONAL"]'), 'JTBNN'))

                    # Select Branch dropdown
                    branch_dropdown = wait.until(EC.presence_of_element_located(
                        (By.XPATH, '//*[@id="P53_BRANCH"]/option[2]')))
                    branch_dropdown.click()

                    # Clear date input form
                    date_from_input = wait.until(EC.presence_of_element_located(
                        (By.XPATH, '//*[@id="P53_DATE1"]')))
                    date_from_input.clear()

                    date_thru_input = wait.until(EC.presence_of_element_located(
                        (By.XPATH, '//*[@id="P53_DATE2"]')))
                    date_thru_input.clear()

                    # Input date
                    date_from_input.send_keys(used_date.strftime('%d-%b-%Y'))
                    date_thru_input.send_keys(used_date.strftime('%d-%b-%Y'))

                    # Proceed
                    go_btn = wait.until(EC.presence_of_element_located(
                        (By.XPATH, '/html/body/form/div[1]/div[2]/div[2]/main/div[2]/div/div[1]/div/div/div/div[3]/div/button')))
                    go_btn.click()

                    # Wait data to be generated
                    time.sleep(10)
                    wait.until(EC.invisibility_of_element_located(
                        (By.XPATH, '/html/body/form/div[1]/div[2]/div[2]/main/div[2]/div/div[2]/div/div/div/div[2]/div[2]/div/div[1]/div/div[2]/div[2]/div[2]/div/div/div[1]/table/tbody/tr/td[2]/button')
                    ))
                    time.sleep(5)

                    # Refresh after data generated
                    driver.refresh()

                    # Go into HAWB Page
                    result_link = wait.until(EC.presence_of_element_located(
                        (By.XPATH, '/html/body/form/div[1]/div[2]/div[2]/main/div[2]/div/div[3]/div/div/div[2]/div[2]/div[2]/div/div/div[1]/table/tbody/tr/td[1]/a')))
                    result_link.click()

                    # Remove Shipment Type : Domestic
                    filter_checkbox = wait.until(EC.presence_of_element_located(
                        (By.XPATH, '/html/body/form/div[1]/div[2]/div[2]/main/div[2]/div/div/div/div/div[2]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/ul/li/span[1]/input')))
                    filter_checkbox.click()
                    time.sleep(5)

                    # Wait table data to be displayed
                    wait.until(EC.presence_of_element_located(
                        (By.XPATH, '/html/body/form/div[1]/div[2]/div[2]/main/div[2]/div/div/div/div/div[2]/div[2]/div[2]/div/div[2]/div[2]/div[6]/div[1]/div/div[1]/table/tr/th[1]/a')))

                    # Open Action Dropdown
                    inbound_action_btn = wait.until(
                        EC.presence_of_element_located((By.ID, 'INB_actions_button')))
                    inbound_action_btn.click()

                    # Open download menu
                    download_menu = wait.until(EC.presence_of_element_located(
                        (By.XPATH, '/html/body/div[3]/div/ul/li[14]/div/span[1]/button')))
                    download_menu.click()

                    # Select data file extension
                    excel_btn = wait.until(EC.presence_of_element_located(
                        (By.XPATH, '/html/body/div[5]/div[2]/div/ul/li[3]')))
                    excel_btn.click()

                    # Proceed to download
                    download_btn = wait.until(EC.presence_of_element_located(
                        (By.XPATH, '/html/body/div[5]/div[3]/div/button[2]')))
                    download_btn.click()

                    # Record the download timestamp
                    download_time = datetime.now()
                    self.log.insert(
                        tk.END, f"{datetime.now().strftime('%H:%M')}  - Downloading data {used_date.strftime('%d-%b-%Y')} \n")
                    self.log.see("end")

                    # Wait for file to be downloaded, if in 5 minutes files not detected try to download again
                    while not os.path.isfile(rf"{self.working_dir}\{self.penarikan}\Origin.xlsx"):
                        time.sleep(15)
                        if datetime.now() - download_time > timedelta(minutes=5):
                            download_btn.click()
                            download_time = datetime.now()

                    # Rename downloaded file
                    default_filepath = rf"{self.working_dir}\{self.penarikan}\Origin.xlsx"
                    if os.path.isfile(default_filepath):
                        # Rename file
                        renamed_file = rf"{self.working_dir}\{self.penarikan}\{used_date.strftime('%d-%b-%Y')}.xlsx"
                        os.rename(default_filepath, renamed_file)

                        # Verifiy data date
                        is_verified = check_file(
                            file=renamed_file, current_date=used_date.date(), column_name='Cnote Date')

                        # Append file to list of downloaded file
                        if is_verified:
                            saved_files.append(renamed_file)

                # Change date
                used_date += timedelta(days=1)

            # Close webdriver
            driver.quit()

            # Combine data if required
            combined_file = ""
            if self.is_combine == True:
                self.log.insert(
                    tk.END, f"{datetime.now().strftime('%H:%M')}  - Combining data... \n")
                self.log.see("end")
                combine_files(files=saved_files,
                              start_date=self.date_from, end_date=self.date_thru, is_standalone=False)

            # Upload to apex
            if self.is_apex == True:
                self.log.insert(
                    tk.END, f"{datetime.now().strftime('%H:%M')} - Uploading data to ApexDB... \n")
                self.log.see("end")
                apex_fn = ApexDB(username=self.username_apex,
                                 password=self.password_apex,
                                 files=combined_file,
                                 name=self.apex_name,
                                 time="",
                                 apex_type=self.apex_type,
                                 working_dir=self.working_dir,
                                 awb_column="Cnote No")
                apex_fn.send_to_apex()

            self.log.insert(
                tk.END, f"{datetime.now().strftime('%H:%M')}  - Process finished \n")
            self.log.see("end")
            showinfo(title="Message", message="Proses Selesai")
        except Exception as e:
            driver.quit()
            self.log.insert(
                tk.END, f"{datetime.now().strftime('%H:%M')}  - Process failed \n")
            self.log.insert(
                tk.END, f"{datetime.now().strftime('%H:%M')}  - Reason: {e} \n")
            self.log.see("end")
            showinfo(title="Error", message=f'{e}')
