import os
import time
import tkinter as tk
from datetime import timedelta, datetime
from tkinter.messagebox import showinfo
from selenium.webdriver.common.by import By
from functions.general_function import combine_files
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from webdriver_manager.microsoft import EdgeChromiumDriverManager


def outbound_data(username, password, date_thru, date_from, loop, combine, penarikan, working_dir, log):
    options = webdriver.EdgeOptions()
    options.use_chromium = True
    options.add_argument("start-maximized")
    options.add_argument("inprivate")
    options.add_experimental_option("prefs", {
        "download.default_directory": rf"{working_dir}\{penarikan}"})
    driver = webdriver.Edge(service=Service(
        EdgeChromiumDriverManager().install()), options=options)

    saved_files = []

    try:
        driver.get(
            'http://10.18.2.51:8080/apex/f?p=141:LOGIN_DESKTOP:15219642399791:::::')

        used_date = date_thru
        wait = WebDriverWait(driver, 10)

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
            tk.END, f"{datetime.now().strftime('%H.%M %p')} - Logged in as {username} \n")

        current_url = driver.current_url
        url_parts = current_url.split(":")

        current_page = url_parts[3]
        session = url_parts[4]

        outbound_page = f"http://10.18.2.51:8080/apex/f?p=141:54:{session}::NO:RIR,1:P0_REP_FLAG:1"

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

            result_page = f'http://10.18.2.51:8080/apex/f?p=141:95:{session}::NO:RP,RIR,95::'
            driver.get(result_page)

            download_page = f"http://10.18.2.51:8080/apex/f?p=141:95:{session}:CSV::::"
            driver.get(download_page)

            log.insert(
                tk.END, f"{datetime.now().strftime('%H.%M %p')} - Downloading data {used_date.strftime('%d-%b-%Y')} \n")

            # Wait for file to be downloaded
            while not os.path.isfile(rf'{working_dir}\{penarikan}\detail.csv'):
                time.sleep(15)

            # Rename downloaded file
            if os.path.isfile(rf'{working_dir}\{penarikan}\detail.csv'):
                os.rename(rf'{working_dir}\{penarikan}\detail.csv',
                          rf'{working_dir}\{penarikan}\{used_date.strftime("%d-%b-%Y")}.csv')
                saved_files.append(
                    rf'{working_dir}\{penarikan}\{used_date.strftime("%d-%b-%Y")}.csv')

            used_date -= timedelta(days=1)

        if combine == True:
            log.insert(
                tk.END, f"{datetime.now().strftime('%H.%M %p')} - Combining data... \n")
            combine_files(files=saved_files,
                          start_date=date_from, end_date=date_thru, is_standalone=False)

        driver.quit()
        log.insert(
            tk.END, f"{datetime.now().strftime('%H.%M %p')} - Process finished \n")
        showinfo(title="Message", message="Proses Selesai")
    except Exception as e:
        driver.quit()
        log.insert(
            tk.END, f"{datetime.now().strftime('%H.%M %p')} - Process failed \n")
        log.insert(
            tk.END, f"{datetime.now().strftime('%H.%M %p')} - Reason: {e} \n")
        showinfo(title="Error", message=e)
