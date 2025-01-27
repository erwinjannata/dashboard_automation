import os
import configparser
import tkinter as tk
from datetime import datetime
from dotenv import load_dotenv
from functions.general_function import get_latest_version, download_latest_package
from tkinter.messagebox import showinfo


def check_update_function(log_box):
    # Fetch latest version in cloud
    load_dotenv()
    log_box.insert(
        tk.END, f"{datetime.now().strftime('%H:%M')} - Checking latest version... \n")
    log_box.see("end")
    latest_version = get_latest_version(repo_name=os.getenv("REPO_NAME"), repo_owner=os.getenv(
        "REPO_OWNER"), branch=os.getenv("REPO_BRANCH"), access_token=os.getenv("ACCESS_TOKEN"), file_path=os.getenv("FILE_PATH"), log_box=log_box)

    # Read current version in local
    config = configparser.ConfigParser()
    config.read('config.ini')
    current_version = config.get('app_version', 'version')

    if current_version == latest_version:
        showinfo(title="Info",
                 message="Versi aplikasi sudah terbaru")
    elif latest_version == None:
        return None
    else:
        log_box.insert(
            tk.END, f"{datetime.now().strftime('%H:%M')} - New version found, dowloading packages... \n")
        log_box.see("end")

        is_update_done = download_latest_package(repo_owner=os.getenv("REPO_OWNER"), repo_name=os.getenv("REPO_NAME"), branch=os.getenv(
            "REPO_BRANCH"), file_path=os.getenv("FILE_PATH_PACKAGES"), access_token=os.getenv("ACCESS_TOKEN"), log_box=log_box)

        if is_update_done:
            showinfo(title="Info",
                     message="Versi terbaru berhasil diunduh, silahkan mulai ulang aplikasi")
