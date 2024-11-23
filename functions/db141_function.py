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


def outbound_data141(username, password, date_thru, date_from, loop, combine, penarikan, working_dir, log):
    config = configparser.ConfigParser()
    config.read('config.ini')

    base_link = config.get('base_links', '141')

    download_dir = rf"{working_dir}\{penarikan}"
    options = webdriver.EdgeOptions()
    options.use_chromium = True
    options.add_argument("start-maximized")
    options.add_argument("inprivate")
    options.add_experimental_option("prefs", {
        "download.default_directory": download_dir})
    driver = webdriver.Edge(service=Service(
        EdgeChromiumDriverManager().install()), options=options)
    used_date = date_thru
    wait = WebDriverWait(driver, 3600)

    saved_files = []

    driver.get(base_link)

    try:
        username_form = wait.until(
            EC.presence_of_element_located((By.ID, 'P101_USERNAME')))

        password_form = wait.until(
            EC.presence_of_element_located((By.ID, 'P101_PASSWORD')))

        login_button = wait.until(
            EC.presence_of_element_located((By.ID, 'B12943410900979389915')))

        username_form.send_keys(username)
        password_form.send_keys(password)

        login_button.click()

        log.insert(
            tk.END, f"{datetime.now().strftime('%H:%M')} - Logged in as {username} \n")
        log.see("end")

        current_url = driver.current_url
        url_parts = current_url.split(":")

        current_page = url_parts[2]
        session = url_parts[3]

        # outbound_page = f"{base_link}:54:{session}::NO:RIR,1:P0_REP_FLAG:1"
        outbound_page = f"{base_link}:54:{session}::NO:RIR,54:P0_REP_FLAG:1"

        for i in range(0, loop):
            driver.get(outbound_page)

            wait.until(
                EC.presence_of_element_located((By.NAME, 'P0_DATE1')))
            date_from_input = driver.find_element(By.NAME, 'P0_DATE1')
            date_from_input.clear()
            wait.until(
                EC.presence_of_element_located((By.NAME, 'P0_DATE2')))
            date_thru_input = driver.find_element(By.NAME, 'P0_DATE2')
            date_thru_input.clear()

            date_from_input.send_keys(used_date.strftime('%d-%b-%Y'))
            date_thru_input.send_keys(used_date.strftime('%d-%b-%Y'))

            go_btn = wait.until(EC.presence_of_element_located(
                (By.ID, 'B12943525354637205515')))
            go_btn.click()

            time.sleep(10)

            wait.until(EC.invisibility_of_element_located(
                (By.XPATH, '//*[@id="loadingIcon"]')
            ))

            result_page = f'{base_link}:95:{session}::NO:RP,RIR,95::'
            driver.get(result_page)

            download_page = f"{base_link}:95:{session}:CSV::::"
            driver.get(download_page)

            log.insert(
                tk.END, f"{datetime.now().strftime('%H:%M')} - Downloading data {used_date.strftime('%d-%b-%Y')} \n")
            log.see("end")

            # Wait for file to be downloaded
            while not os.path.isfile(rf'{download_dir}\detail.csv'):
                time.sleep(15)

            # Rename downloaded file
            if os.path.isfile(rf'{download_dir}\detail.csv'):
                os.rename(rf'{download_dir}\detail.csv',
                          rf'{download_dir}\{used_date.strftime("%d-%b-%Y")}.csv')
                saved_files.append(
                    rf'{download_dir}\{used_date.strftime("%d-%b-%Y")}.csv')

            used_date -= timedelta(days=1)

        if combine == True:
            log.insert(
                tk.END, f"{datetime.now().strftime('%H:%M')} - Combining data... \n")
            log.see("end")
            combine_files(files=saved_files,
                          start_date=date_from, end_date=date_thru, is_standalone=False)

        driver.quit()
        log.insert(
            tk.END, f"{datetime.now().strftime('%H:%M')} - Process finished \n")
        log.see("end")
        showinfo(title="Message", message="Proses Selesai")
    except Exception as e:
        driver.quit()
        log.insert(
            tk.END, f"{datetime.now().strftime('%H:%M')} - Process failed \n")
        log.insert(
            tk.END, f"{datetime.now().strftime('%H:%M')} - Reason: {e} \n")
        log.see("end")
        showinfo(title="Error", message=e)


def inbound_data141(mode, username, password, date_thru, date_from, loop, combine, penarikan, working_dir, log):
    config = configparser.ConfigParser()
    config.read('config.ini')

    base_link = config.get('base_links', '141')

    download_dir = rf"{working_dir}\{penarikan}"
    options = webdriver.EdgeOptions()
    options.use_chromium = True
    options.add_argument("start-maximized")
    options.add_argument("inprivate")
    options.add_experimental_option("prefs", {
        "download.default_directory": download_dir})
    driver = webdriver.Edge(service=Service(
        EdgeChromiumDriverManager().install()), options=options)

    saved_files = []
    driver.get(base_link)

    used_date = date_thru
    wait = WebDriverWait(driver, 3600)

    try:
        # Login
        username_form = wait.until(
            EC.presence_of_element_located((By.ID, 'P101_USERNAME')))

        password_form = wait.until(
            EC.presence_of_element_located((By.ID, 'P101_PASSWORD')))

        login_button = wait.until(
            EC.presence_of_element_located((By.ID, 'B12943410900979389915')))

        username_form.send_keys(username)
        password_form.send_keys(password)

        login_button.click()

        log.insert(
            tk.END, f"{datetime.now().strftime('%H:%M')} - Logged in as {username} \n")
        log.see("end")

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

        # Go to Inbound Page
        inbound_page = f"{base_link}:{page_code[mode][0]}:{session}::NO:RIR,{page_code[mode][0]}:P0_REP_FLAG:{page_code[mode][1]}"

        for i in range(0, loop):
            driver.get(inbound_page)

            if mode == 2:
                wait.until(EC.presence_of_element_located(
                    (By.NAME, 'P0_REP_DATE')))
                option_input = driver.find_element(By.NAME, 'P0_REP_DATE')
                option_input.send_keys('AWB')

            wait.until(
                EC.presence_of_element_located((By.NAME, 'P0_DATE1')))
            date_from_input = driver.find_element(By.NAME, 'P0_DATE1')
            date_from_input.clear()
            wait.until(
                EC.presence_of_element_located((By.NAME, 'P0_DATE2')))
            date_thru_input = driver.find_element(By.NAME, 'P0_DATE2')
            date_thru_input.clear()

            date_from_input.send_keys(used_date.strftime('%d-%b-%Y'))
            date_thru_input.send_keys(used_date.strftime('%d-%b-%Y'))

            go_btn = wait.until(EC.presence_of_element_located(
                (By.ID, 'B12943525354637205515')))
            go_btn.click()

            time.sleep(10)

            wait.until(EC.invisibility_of_element_located(
                (By.XPATH, '//*[@id="loadingIcon"]')
            ))

            # Go to Result Page
            # End-To-End
            if mode == 0:
                result_page = f'{base_link}:64:{session}:IR_110406:NO:RP,64::'
            # Summary & Intracity
            else:
                result_page = f'{base_link}:{page_code[mode][2]}:{session}::NO:RP,RIR,{page_code[mode][2]}::'

            driver.get(result_page)

            download_page = f"{base_link}:{page_code[mode][2]}:{session}:CSV::::"
            driver.get(download_page)

            log.insert(
                tk.END, f"{datetime.now().strftime('%H:%M')} - Downloading data {used_date.strftime('%d-%b-%Y')} \n")
            log.see("end")

            # Wait for file to be downloaded
            while not os.path.isfile(rf'{download_dir}\detail.csv'):
                time.sleep(15)

            # Rename downloaded file
            if os.path.isfile(rf'{download_dir}\detail.csv'):
                os.rename(rf'{download_dir}\detail.csv',
                          rf'{download_dir}\{used_date.strftime("%d-%b-%Y")}.csv')
                saved_files.append(
                    rf'{download_dir}\{used_date.strftime("%d-%b-%Y")}.csv')

            used_date -= timedelta(days=1)

        if combine == True:
            log.insert(
                tk.END, f"{datetime.now().strftime('%H:%M')} - Combining data... \n")
            log.see("end")
            combine_files(files=saved_files,
                          start_date=date_from, end_date=date_thru, is_standalone=False)

        driver.quit()
        log.insert(
            tk.END, f"{datetime.now().strftime('%H:%M')} - Process finished \n")
        log.see("end")
        showinfo(title="Message", message="Proses Selesai")
    except Exception as e:
        driver.quit()
        log.insert(
            tk.END, f"{datetime.now().strftime('%H:%M')} - Process failed \n")
        log.insert(
            tk.END, f"{datetime.now().strftime('%H:%M')} - Reason: {e} \n")
        log.see("end")
        showinfo(title="Error", message=e)
