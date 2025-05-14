from functions.headless_141 import DB141
from functions.headless_117 import DB117
from datetime import datetime, timedelta
from pathlib import Path

# Get Current Date
date: datetime = datetime.today() - timedelta(days=1)
jam_penarikan: str = datetime.now().strftime("%d-%b-%Y %H.%M %p")

# Webdriver Binary Location
driver_location: str = rf"{Path.cwd()}\chromedriver\chromedriver.exe"

# Read current weekday and determine data to be downloaded
loop: int = 1
day: int = datetime.today().weekday()

if day == 0:
    loop = 2

# Outbound 141
dashboard141 = DB141(
    mode=0,
    username_141="HAIRI",
    password_141="@HairiPM1",
    username_apex="",
    password_apex="",
    date_from=date,
    date_thru=date,
    loop=loop,
    is_combine=False,
    is_apex=False,
    apex_type="",
    apex_file_name="",
    penarikan=jam_penarikan,
    working_dir=rf"{Path.cwd()}\141\Outbound",
    driver=driver_location,
)
dashboard141.outbound_data141()

dashboard141.working_dir = rf"{Path.cwd()}\141\Inbound End-To-End"
dashboard141.inbound_data141()


# # Inbound 117
# dashboard117 = DB117(
#     mode=0,
#     username_117="IYUT",
#     password_117="123",
#     username_apex="",
#     password_apex="",
#     date_from=date += timedelta(days=1),
#     date_thru=date += timedelta(days=1),
#     loop=18,
#     is_combine=True,
#     is_apex=False,
#     apex_type="",
#     apex_file_name="",
#     penarikan=jam_penarikan,
#     working_dir=rf"{Path.cwd()}\117\Inbound",
#     driver=driver_location,
# )
# dashboard117.inbound_data117()
