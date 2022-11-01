import os
import sys
import logging
from datetime import datetime
from exiles_api import ChatLogs
from google_api.sheets import Spreadsheet
from logger import get_logger
from config import (
    LOGS_SPREADSHEET_ID, LOGS_CHAT_SHEET_ID, CHAT_LOG_HOLD_BACK, SAVED_DIR_PATH,
    LOG_LEVEL_STDOUT, LOG_LEVEL_FILE
)

# catch unhandled exceptions
logger = get_logger('admin_process_chatlogs.log', log_level_stdout=LOG_LEVEL_STDOUT, log_level_file=LOG_LEVEL_FILE)


def handle_exception(exc_type, exc_value, exc_traceback):
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return

    logger.error("Uncaught exception", exc_info=(exc_type, exc_value, exc_traceback))


sys.excepthook = handle_exception

# save current time
now = datetime.utcnow()
if LOG_LEVEL_STDOUT > logging.INFO:
    print("Updating chat log sheets...")
logger.info("Updating chat log sheets...")


# instanciate the Spreadsheet object
logger.debug("Reading current data from the chatlog sheet.")
sheets = Spreadsheet(LOGS_SPREADSHEET_ID, activeSheetId=LOGS_CHAT_SHEET_ID)

values = [['Last Upload: ' + now.strftime("%d-%b-%Y %H:%M"), '', 'Hold Back Time: ' + str(CHAT_LOG_HOLD_BACK.days)],
          ['Date', 'Sender', 'Channel', 'Type', 'Chat message']]
sheets.update(range='Chat Log!A1:E2', values=values)

# make sure there'r three rows in the sheet
lastRow = sheets.get_properties()["gridProperties"]["rowCount"]
if lastRow == 2:
    sheets.insert_rows(startIndex=3, numRows=1)
sheets.set_frozen(rows=2)

# threshold beyond which no dates should be kept
ageThreshold = now - CHAT_LOG_HOLD_BACK

""" Read datetimes from spreadsheet """

# ensure that sheet is sorted from olderst (row 3) to latest (second to last row)
sheets.sort(sortCol=1, startRowIndex=3)
sheets.commit()
# read the datetimes of the chatlines that are already uploaded to the sheet
datetimes = [d[0] for d in sheets.read('Chat Log!A3:A', is_ordinal=True)]
lastRow = len(datetimes) + 2

""" Remove old data from spreadsheet """

logger.debug("Remove old data from chatlog sheet.")
rowsToDelete = 0
# remove the oldest entry while it's older than the calculated threshold and keep count
while(len(datetimes) > 1 and datetimes[0] < ageThreshold):
    rowsToDelete += 1
    del datetimes[0]

# remove as many lines as have been counted from the top of the sheet (excluding the header)
if rowsToDelete > 0:
    sheets.delete_rows(startIndex=3, numRows=rowsToDelete)
    sheets.commit()
    lastRow -= rowsToDelete

""" Read the required chatlog lines from the logfile """

logger.debug("Reading new entries from the logfile.")
# grab the newest date for comparison. If dates is empty or date is older use ageThreshold instead
newestDate = datetimes[-1] if datetimes and datetimes[-1] > ageThreshold else ageThreshold
# read only the chatfiles that are not already uploaded to the sheet. Order is from oldest to newest
LOGS_PATH = os.path.join(SAVED_DIR_PATH, "Logs")
logs = ChatLogs(LOGS_PATH, newestDate)
logs.get_lines()
# delete the oldest file if there are 3 (default) or more files and use the last edit date to rename the youngest
logs.cycle_log_files()

""" Write new data to spreadsheet """

logger.debug("Updating chatlog sheet with new lines.")
values = []
for line in logs.chat_lines:
    values.append(logs.get_chat_info(line))

numRows = len(values)
if numRows > 0:
    firstInsert = lastRow + 1
    lastRow += numRows
    pattern = 'ddd dd-mmm-yyy hh:mm:ss'
    range = 'Chat Log!A' + str(firstInsert) + ':E' + str(lastRow)
    sheets.insert_rows(startIndex=firstInsert, numRows=numRows, inheritFromBefore=True)
    sheets.set_format(startRowIndex=firstInsert, endRowIndex=lastRow, type='DATE_TIME', pattern=pattern)
    sheets.set_filter(startRowIndex=2, endRowIndex=lastRow)
    sheets.commit()
    sheets.update(range=range, values=values)

execTime = datetime.utcnow() - now
execTimeStr = str(execTime.seconds) + "." + str(execTime.microseconds)
if LOG_LEVEL_STDOUT > logging.INFO:
    print(f"Done! Required time: {execTimeStr} sec.")
logger.info(f"Done! Required time: {execTimeStr} sec.")
