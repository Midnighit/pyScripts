import sys
import logging
from sqlalchemy import desc
from datetime import datetime
from exiles_api import session, ServerPopulationRecordings as PopRecs
from google_api.sheets import Spreadsheet
from logger import get_logger
from config import (
    PLAYER_SPREADSHEET_ID, PLAYER_ACTIVITY_SHEET_ID, ACTIVITY_HOLD_BACK, MAX_POP, LOG_LEVEL_STDOUT, LOG_LEVEL_FILE
)

# catch unhandled exceptions
logger = get_logger('player_activity.log', log_level_stdout=LOG_LEVEL_STDOUT, log_level_file=LOG_LEVEL_FILE)


def handle_exception(exc_type, exc_value, exc_traceback):
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return

    logger.error("Uncaught exception", exc_info=(exc_type, exc_value, exc_traceback))


sys.excepthook = handle_exception

# save current time
now = datetime.utcnow()
if LOG_LEVEL_STDOUT > logging.INFO:
    print("Updating activity statistics sheet...")
logger.info("Updating activity statistics sheet...")


# instanciate the Spreadsheet object
sheets = Spreadsheet(PLAYER_SPREADSHEET_ID, activeSheetId=PLAYER_ACTIVITY_SHEET_ID)

# threshold beyond which no dates should be kept
ageThreshold = now - ACTIVITY_HOLD_BACK

""" Read data from spreadsheet and db """

logger.debug("Reading current data from the activity sheet.")
# read the values that are already uploaded to the sheet. Last row is kept empty
lastRow = sheets.get_properties()["gridProperties"]["rowCount"] - 1
# dates list is ordered from oldest (index 0) to latest
dates = [r[0] for r in sheets.read('Activity Statistics!A2:A' + str(lastRow), is_ordinal=True)] if lastRow >= 2 else []

# grab the newest date for comparison. If dates is empty or date is older use ageThreshold instead
newestDate = dates[-1] if len(dates) > 0 and dates[-1] > ageThreshold else ageThreshold

logger.debug("Reading new data from game.db.")
values = []
# read all the timestamps and population values from ServerPopulationRecordings starting with the newest
for record in session.query(PopRecs).order_by(desc(PopRecs.time_of_recording)).all():
    # while the newest date found in the db is newer than the newest date read from the sheet add it to values
    checkDate = datetime.utcfromtimestamp(record.time_of_recording)
    if checkDate > newestDate:
        values.insert(0, [checkDate.strftime("%Y-%m-%d %H:%M:%S"), int(record.population * MAX_POP)])
    else:
        break
session.close()

""" Write new data to spreadsheet """

logger.debug("Updating activity sheet with new lines.")
numRows = len(values)
if numRows > 0:
    firstInsert = lastRow + 1
    lastRow += numRows
    range = 'Activity Statistics!A' + str(firstInsert) + ':B' + str(lastRow)
    sheets.insert_rows(startIndex=firstInsert, numRows=numRows)
    sheets.set_format(startRowIndex=firstInsert, endRowIndex=lastRow, type='DATE_TIME', pattern='ddd dd-mmm-yyy')
    sheets.set_format(
        startColumnIndex=2,
        endColumnIndex=2,
        startRowIndex=firstInsert,
        endRowIndex=lastRow,
        type='NUMBER',
        pattern='0'
    )
    sheets.commit()
    sheets.update(range=range, values=values)

""" Remove old data from spreadsheet """

logger.debug("Delete old lines from activity sheet.")
rowsToDelete = 0
# remove the oldest entry while it's older than the calculated threshold and keep count
while(len(dates) > 1 and dates[0] < ageThreshold):
    rowsToDelete += 1
    del dates[0]

# remove as many lines as have been counted from the top of the sheet (excluding the header)
if rowsToDelete > 0:
    sheets.delete_rows(startIndex=2, numRows=rowsToDelete)
    sheets.commit()

execTime = datetime.utcnow() - now
execTimeStr = str(execTime.seconds) + "." + str(execTime.microseconds)
if LOG_LEVEL_STDOUT > logging.INFO:
    print(f"Done! Required time: {execTimeStr} sec.")
logger.info(f"Done! Required time: {execTimeStr} sec.")
