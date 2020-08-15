import os, time
from datetime import datetime
from config import *
from google_api.sheets import Spreadsheet

LOG_PATH = os.path.join(SAVED_DIR_PATH, "Logs/ConanSandbox.log")

sheets = Spreadsheet(ADMIN_SPREADSHEET_ID, activeSheetId=ADMIN_LINT_SHEET_ID)

stats = os.stat(LOG_PATH)
last_ctime = this_ctime = stats[8]
while True:
    stats = os.stat(LOG_PATH)
    this_ctime = stats[8]
    if this_ctime != last_ctime:
        diff = this_ctime - last_ctime
        last_ctime = this_ctime
        if diff > 60:
            now = datetime.utcnow().strftime("%d-%b-%Y %H:%M:%S")
            sheets.insert_rows(startIndex=2)
            sheets.set_format(startRowIndex=2, endRowIndex=2, type='DATE_TIME', pattern='ddd dd-mmm-yyy')
            sheets.set_format(startColumnIndex=2, endColumnIndex=2, startRowIndex=2, endRowIndex=2, type='NUMBER', pattern='0')
            sheets.commit()
            sheets.update(range='Log write interval!A2:B2', values=[[now, diff]])
            print(f"[{now}]: {diff} seconds since last change. Resetting timer.")
    time.sleep(1)
