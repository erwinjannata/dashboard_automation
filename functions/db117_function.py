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


def inbound_data117(username, password, date_thru, date_from, loop, combine, penarikan, working_dir, log):
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
            'http://app.jne.co.id:7777/ords/f?p=117:LOGIN_DESKTOP:410809499327:::::')

        used_date = date_thru
        wait = WebDriverWait(driver, 3600)

        username_form = wait.until(
            EC.presence_of_element_located((By.ID, 'P101_USERNAME')))

        password_form = wait.until(
            EC.presence_of_element_located((By.ID, 'P101_PASSWORD')))

        login_button = wait.until(
            EC.presence_of_element_located((By.ID, 'B3779048980085325187')))

        username_form.send_keys(username)
        password_form.send_keys(password)

        login_button.click()

        log.insert(
            tk.END, f"{datetime.now().strftime('%H:%M')}  - Logged in as {username} \n")
        log.see("end")

        current_url = driver.current_url
        url_parts = current_url.split(":")

        current_page = url_parts[3]
        session = url_parts[4]

        inbound_page = f"http://app.jne.co.id:7777/ords/f?p=117:56:{session}:::::"

        for i in range(0, loop):
            driver.get(inbound_page)

            # Fill Regional
            regional_btn = wait.until(EC.presence_of_element_located(
                (By.ID, 'P56_REGIONAL_DEST_lov_btn')))
            regional_btn.click()

            jtbnn_link = wait.until(EC.presence_of_element_located(
                (By.CSS_SELECTOR, '#PopupLov_56_P56_REGIONAL_DEST_dlg > div.a-PopupLOV-results.a-TMV > div > div.a-TMV-w-scroll > ul > li')))
            jtbnn_link.click()

            # Fill Branch
            branch_btn = wait.until(EC.presence_of_element_located(
                (By.ID, 'P56_BRANCH_DEST_lov_btn')))
            branch_btn.click()

            branch_ami_link = wait.until(EC.presence_of_element_located(
                (By.CSS_SELECTOR, '#PopupLov_56_P56_BRANCH_DEST_dlg > div.a-PopupLOV-results.a-TMV > div > div.a-TMV-w-scroll > ul > li')))
            branch_ami_link.click()

            # Fill Date
            wait.until(EC.presence_of_element_located((By.NAME, 'P56_DATE1')))
            date_from_input = driver.find_element(By.NAME, 'P56_DATE1')
            date_from_input.clear()

            wait.until(EC.presence_of_element_located((By.NAME, 'P56_DATE2')))
            date_thru_input = driver.find_element(By.NAME, 'P56_DATE2')
            date_thru_input.clear()

            date_from_input.send_keys(used_date.strftime('%d-%b-%Y'))
            date_thru_input.send_keys(used_date.strftime('%d-%b-%Y'))

            go_btn = wait.until(EC.presence_of_element_located(
                (By.ID, 'B4053153599493451337')))
            go_btn.click()

            find_text = 'DONE'
            wait.until(EC.text_to_be_present_in_element(
                (By.XPATH, '//*[@id="R4053268519926012741"]/div[2]/div[2]'), find_text))

            driver.refresh()

            # Go into HAWB Page
            hawb_link = wait.until(EC.presence_of_element_located(
                (By.XPATH, '//*[@id="report_INB"]/div/div[1]/table/tbody/tr/td[5]/a')))
            hawb_link.click()

            inbound_action_btn = wait.until(
                EC.presence_of_element_located((By.ID, 'INB_actions_button')))
            inbound_action_btn.click()

            download_menu = wait.until(EC.presence_of_element_located(
                (By.XPATH, '//*[@id="INB_actions_menu_14i"]')))
            download_menu.click()

            excel_btn = wait.until(EC.presence_of_element_located(
                (By.XPATH, '//*[@id="INB_download_formats"]/li[3]')))
            excel_btn.click()

            data_only = wait.until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="INB_download_options"]/div[2]/div[1]/div/div[2]/label/span')))
            data_only.click()

            download_btn = wait.until(EC.presence_of_element_located(
                (By.XPATH, '//*[@id="t_PageBody"]/div[5]/div[3]/div/button[2]')))

            download_btn.click()
            download_time = datetime.now()

            log.insert(
                tk.END, f"{datetime.now().strftime('%H:%M')}  - Downloading data {used_date.strftime('%d-%b-%Y')} \n")
            log.see("end")

            # Wait for file to be downloaded, if in 5 minutes files not detected try to download again
            while not os.path.isfile(rf"{working_dir}\{penarikan}\Inbound .xlsx"):
                time.sleep(15)
                if datetime.now() - download_time > timedelta(minutes=5):
                    download_btn.click()
                    download_time = datetime.now()

            # Rename downloaded file
            if os.path.isfile(rf"{working_dir}\{penarikan}\Inbound .xlsx"):
                os.rename(rf"{working_dir}\{penarikan}\Inbound .xlsx",
                          rf"{working_dir}\{penarikan}\{used_date.strftime('%d-%b-%Y')}.xlsx")
                saved_files.append(
                    rf"{working_dir}\{penarikan}\{used_date.strftime('%d-%b-%Y')}.xlsx")

            used_date -= timedelta(days=1)

        if combine == True:
            log.insert(
                tk.END, f"{datetime.now().strftime('%H:%M')}  - Combining data... \n")
            log.see("end")
            combine_files(files=saved_files,
                          start_date=date_from, end_date=date_thru, is_standalone=False)

        driver.quit()
        log.insert(
            tk.END, f"{datetime.now().strftime('%H:%M')}  - Process finished \n")
        log.see("end")
        showinfo(title="Message", message="Proses Selesai")
    except Exception as e:
        driver.quit()
        log.insert(
            tk.END, f"{datetime.now().strftime('%H:%M')}  - Process failed \n")
        log.insert(
            tk.END, f"{datetime.now().strftime('%H:%M')}  - Reason: {e} \n")
        log.see("end")
        showinfo(title="Error", message=f'{e}')


def outbound_data117(username, password, date_thru, date_from, loop, combine, penarikan, working_dir, log):
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
            'http://app.jne.co.id:7777/ords/f?p=117:LOGIN_DESKTOP:410809499327:::::')

        used_date = date_thru
        wait = WebDriverWait(driver, 3600)

        username_form = wait.until(
            EC.presence_of_element_located((By.ID, 'P101_USERNAME')))

        password_form = wait.until(
            EC.presence_of_element_located((By.ID, 'P101_PASSWORD')))

        login_button = wait.until(
            EC.presence_of_element_located((By.ID, 'B3779048980085325187')))

        username_form.send_keys(username)
        password_form.send_keys(password)

        login_button.click()

        log.insert(
            tk.END, f"{datetime.now().strftime('%H:%M')}  - Logged in as {username} \n")
        log.see("end")

        current_url = driver.current_url
        url_parts = current_url.split(":")

        current_page = url_parts[3]
        session = url_parts[4]

        outbound_page = f"http://app.jne.co.id:7777/ords/f?p=117:53:{session}:::::"

        for i in range(0, loop):
            driver.get(outbound_page)

            # Fill Regional
            regional_btn = wait.until(EC.presence_of_element_located(
                (By.ID, 'P53_REGIONAL')))
            regional_btn.send_keys('JTBNN')
            wait.until(EC.text_to_be_present_in_element((By.ID, 'P53_REGIONAL'), 'JTBNN'))

            # Fill Branch
            branch_btn = wait.until(EC.presence_of_element_located(
                (By.XPATH, '//*[@id="P53_BRANCH"]/option[2]')))
            branch_btn.click()

            # Fill Date
            wait.until(EC.presence_of_element_located((By.NAME, 'P53_DATE1')))
            date_from_input = driver.find_element(By.NAME, 'P53_DATE1')
            date_from_input.clear()

            wait.until(EC.presence_of_element_located((By.NAME, 'P53_DATE2')))
            date_thru_input = driver.find_element(By.NAME, 'P53_DATE2')
            date_thru_input.clear()

            date_from_input.send_keys(used_date.strftime('%d-%b-%Y'))
            date_thru_input.send_keys(used_date.strftime('%d-%b-%Y'))

            go_btn = wait.until(EC.presence_of_element_located(
                (By.ID, 'B4017712943882114109')))
            go_btn.click()

            wait.until(EC.invisibility_of_element_located(
                (By.XPATH, '//*[@id="loadingIcon"]')
            ))
            
            # Go into HAWB Page
            result_page = f"http://app.jne.co.id:7777/ords/f?p=117:55:{session}:::RP,RIR,55:IREQ_SHIPMENT_TYPE:DOMESTIC"
            driver.get(result_page)

            # Remove Shipment Type : Domestic
            remove_button = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '#a_Collapsible1_INB_control_panel_content > ul > li > span.a-IRR-controls-cell.a-IRR-controls-cell--remove > button')))
            remove_button.click()

            wait.until_not(EC.presence_of_element_located((By.XPATH, '//*[@id="a_Collapsible1_INB_control_panel_content"]/ul')))

            inbound_action_btn = wait.until(
                EC.presence_of_element_located((By.ID, 'INB_actions_button')))
            inbound_action_btn.click()

            download_menu = wait.until(EC.presence_of_element_located(
                (By.XPATH, '//*[@id="INB_actions_menu_14i"]')))
            
            download_menu.click()

            excel_btn = wait.until(EC.presence_of_element_located(
                (By.XPATH, '//*[@id="INB_download_formats"]/li[3]')))
            excel_btn.click()

            data_only = wait.until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="INB_download_options"]/div[2]/div[1]/div/div[2]/label/span')))
            data_only.click()

            download_btn = wait.until(EC.presence_of_element_located(
                (By.XPATH, '//*[@id="t_PageBody"]/div[5]/div[3]/div/button[2]')))

            download_btn.click()
            download_time = datetime.now()

            log.insert(
                tk.END, f"{datetime.now().strftime('%H:%M')}  - Downloading data {used_date.strftime('%d-%b-%Y')} \n")
            log.see("end")

            # Wait for file to be downloaded, if in 5 minutes files not detected try to download again
            while not os.path.isfile(rf"{working_dir}\{penarikan}\Origin.xlsx"):
                time.sleep(15)
                if datetime.now() - download_time > timedelta(minutes=5):
                    download_btn.click()
                    download_time = datetime.now()

            # Rename downloaded file
            if os.path.isfile(rf"{working_dir}\{penarikan}\Origin.xlsx"):
                os.rename(rf"{working_dir}\{penarikan}\Origin.xlsx",
                          rf"{working_dir}\{penarikan}\{used_date.strftime('%d-%b-%Y')}.xlsx")
                saved_files.append(
                    rf"{working_dir}\{penarikan}\{used_date.strftime('%d-%b-%Y')}.xlsx")

            used_date -= timedelta(days=1)

        if combine == True:
            log.insert(
                tk.END, f"{datetime.now().strftime('%H:%M')}  - Combining data... \n")
            log.see("end")
            combine_files(files=saved_files,
                          start_date=date_from, end_date=date_thru, is_standalone=False)

        driver.quit()
        log.insert(
            tk.END, f"{datetime.now().strftime('%H:%M')}  - Process finished \n")
        log.see("end")
        showinfo(title="Message", message="Proses Selesai")
    except Exception as e:
        driver.quit()
        log.insert(
            tk.END, f"{datetime.now().strftime('%H:%M')}  - Process failed \n")
        log.insert(
            tk.END, f"{datetime.now().strftime('%H:%M')}  - Reason: {e} \n")
        log.see("end")
        showinfo(title="Error", message=f'{e}')