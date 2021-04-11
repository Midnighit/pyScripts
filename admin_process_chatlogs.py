import os, sys
from datetime import datetime
from config import *
from exiles_api import ChatLogs
from google_api.sheets import Spreadsheet

# save current time
now = datetime.utcnow()
print("Updating chat log sheets...")

# instanciate the Spreadsheet object
sheets = Spreadsheet(LOGS_SPREADSHEET_ID, activeSheetId=LOGS_CHAT_SHEET_ID)

values = [['Last Upload: ' + now.strftime("%d-%b-%Y %H:%M"), '', 'Hold Back Time: ' + str(CHAT_LOG_HOLD_BACK.days)],
          ['Date', 'Sender', 'Recipient', 'Channel', 'Chat message']]
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

# grab the newest date for comparison. If dates is empty or date is older use ageThreshold instead
newestDate = datetimes[-1] if datetimes and datetimes[-1] > ageThreshold else ageThreshold
# newestDate = datetime(2021, 4, 10, 12)
# lastRow = 9
# read only the chatfiles that are not already uploaded to the sheet. Order is from oldest to newest
LOGS_PATH = os.path.join(SAVED_DIR_PATH, "Logs")
logs = ChatLogs(LOGS_PATH, newestDate)
logs.get_lines()

""" Write new data to spreadsheet """

values = []
for line in logs.chat_lines:
    values.append(logs.get_chat_info(line, date_format="%Y-%m-%d %H:%M:%S"))

numRows = len(values)
if numRows > 0:
    firstInsert = lastRow + 1
    lastRow += numRows
    range = 'Chat Log!A' + str(firstInsert) + ':E' + str(lastRow)
    sheets.insert_rows(startIndex=firstInsert, numRows=numRows, inheritFromBefore=False)
    sheets.set_format(startRowIndex=firstInsert, endRowIndex=lastRow, type='DATE_TIME', pattern='ddd dd-mmm-yyy hh:mm:ss')
    sheets.set_filter(startRowIndex=2, endRowIndex=lastRow)
    sheets.commit()
    sheets.update(range=range, values=values)

execTime = datetime.utcnow() - now
execTimeStr = str(execTime.seconds) + "." + str(execTime.microseconds)
print(f"Done!\nRequired time: {execTimeStr} sec.")
