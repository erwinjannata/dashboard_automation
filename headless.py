from functions.headless_141 import DB141
from functions.headless_117 import DB117
from datetime import datetime, timedelta
from pathlib import Path
import configparser
import shutil
import os

# Get Current Date
date: datetime = datetime.today() - timedelta(days=1)
jam_penarikan: str = datetime.now().strftime("%d-%b-%Y %H.%M %p")

# Webdriver Binary Location
home_drive: str = Path.home().drive
driver_location: str = rf"{home_drive}\Headless Dashboard Automation\webdriver\msedgedriver.exe"

# Read app configuration
config = configparser.ConfigParser()
config.read(rf'{home_drive}\Headless Dashboard Automation\config.ini')

# Download Directory & User Details
download_dir: str = config.get('directory', "download_dir_141")
cloud_dir: str = config.get('directory', 'cloud_141')
username_141: str = config.get('dashboard_user', "username_141")
password_141: str = config.get('dashboard_user', "password_141")

# Read current weekday and determine data to be downloaded
loop: int = 1
day: int = datetime.today().weekday()
is_combine: bool = False
# apex_filename: str = f"Outbound_{date.strftime('%d%b')}"

# Determine how many days to be downloaded
if day == 0:
    loop = 2
    is_combine = True
    # date2 = date - timedelta(days=1)
    # apex_filename = f"Outbound_{date2.strftime('%d')}-{(date.strftime('%d%b'))}"

# Initialize 141 Class
dashboard141 = DB141(
    link=config.get('base_links', '141'),
    mode=0,
    username_141=username_141,
    password_141=password_141,
    username_apex="",
    password_apex="",
    date_from=date,
    date_thru=date,
    loop=loop,
    is_combine=is_combine,
    is_apex=False,
    apex_type="",
    apex_file_name="",
    penarikan=jam_penarikan,
    working_dir=rf"{download_dir}\Outbound",
    driver=driver_location,
    cloud_dir=cloud_dir,
)

is_exist = os.path.exists(dashboard141.working_dir)

if is_exist == False:
    # Download Outbound 141 Data
    dashboard141.outbound_data141()

    # Download Inbound 141 Data
    dashboard141.working_dir = rf"{download_dir}\Inbound End-To-End"
    dashboard141.inbound_data141()

    # Copy Downloaded Files to ownCloud and delete original files
    shutil.copytree(download_dir, cloud_dir, dirs_exist_ok=True)
    shutil.rmtree(download_dir)


# ---------------------------------------------------------------------------

# # Inbound 117
# download_dir = config.get('directory', "download_dir_117")
# cloud_dir = config.get('directory', 'cloud_117')
# date_thru = date - timedelta(days=18)

# # Initialize 117 Class
# dashboard117 = DB117(
#     link=config.get('base_links', '117'),
#     username_117=config.get('dashboard_user', "username_117"),
#     password_117=config.get('dashboard_user', "password_117"),
#     username_apex="",
#     password_apex="",
#     date_from=date,
#     date_thru=date_thru,
#     loop=18,
#     is_combine=True,
#     is_apex=False,
#     apex_type="",
#     apex_file_name="",
#     penarikan=jam_penarikan,
#     working_dir=rf"{download_dir}\Inbound",
#     driver=driver_location,
# )

# is_exist = os.path.exists(dashboard117.working_dir)

# if is_exist == False:
#     dashboard117.inbound_data117()

#     # Copy Downloaded Files to ownCloud and delete original files
#     shutil.copytree(download_dir, cloud_dir, dirs_exist_ok=True)
#     shutil.rmtree(download_dir)
