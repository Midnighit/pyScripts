import sys
import logging
from datetime import datetime, timedelta
from exiles_api import db_date, OwnersCache, ObjectsCache, Stats
from google_api.sheets import Spreadsheet
from logger import get_logger
from config import (
    RUINS_CLAN_ID, LOGS_SPREADSHEET_ID, LOGS_CHAT_SHEET_ID, INACTIVITY, PLAYER_SPREADSHEET_ID,
    PLAYER_STATISTICS_SHEET_ID, LOG_LEVEL_STDOUT, LOG_LEVEL_FILE
)

# catch unhandled exceptions
logger = get_logger('player_statistics.log', log_level_stdout=LOG_LEVEL_STDOUT, log_level_file=LOG_LEVEL_FILE)


def handle_exception(exc_type, exc_value, exc_traceback):
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return

    logger.error("Uncaught exception", exc_info=(exc_type, exc_value, exc_traceback))


sys.excepthook = handle_exception

# save current time
now = datetime.utcnow()
if LOG_LEVEL_STDOUT > logging.INFO:
    print("Updating statistics sheet...")
logger.info("Updating statistics sheet...")

# update the caches
logger.debug("Updating OwnersCache.")
OwnersCache.update(RUINS_CLAN_ID)
logger.debug("Updating ObjectsCache.")
ObjectsCache.update(RUINS_CLAN_ID)

# estimate db age by reading the last_login date of the first character in the characters table
if dbAge := db_date():
    dbAgeStr = "Database Date: " + dbAge.strftime("%d-%b-%Y %H:%M UTC")
else:
    execTime = datetime.utcnow() - now
    execTimeStr = str(execTime.seconds) + "." + str(execTime.microseconds)
    logger.info(f"Found no characters in db!\nRequired time: {execTimeStr} sec.")
    sys.exit(0)

# read the number of chatlines
logger.debug("Read number of chatlines from chatlog sheet.")
sheets = Spreadsheet(LOGS_SPREADSHEET_ID, activeSheetId=LOGS_CHAT_SHEET_ID)
lastRow = sheets.get_properties()["gridProperties"]["rowCount"] - 1
dates = [r[0] for r in sheets.read('Chat Log!A3:A' + str(lastRow), is_ordinal=True)] if lastRow >= 2 else []

num_lines = 0
threshold24h = dbAge - timedelta(hours=24)
for idx in range(len(dates)-1, 1, -1):
    if dates[idx] > dbAge:
        continue
    if dates[idx] > threshold24h:
        num_lines += 1
    else:
        break


def divby10k(num):
    """
    Little helper function that converts an integer number into a string and
    then moves the decimal point 4 digits forward equaling a division by 10,000
    If number already has a decimal point, that one is removed and digit appended
    """
    num = str(num)
    part = num.split(".")
    predec, dec = (part[0], "") if len(part) == 1 else (part[0], part[1])
    return float(("0." + predec.zfill(4) if len(predec) < 5 else predec[:-4] + "." + predec[-4:]) + dec)


# Get the statistics
logger.debug("Compile statistics from game.db.")
stats = Stats.get_tile_statistics(INACTIVITY)
value = []
value.append(dbAge.strftime("%d/%m/%Y %H:%M:%S"))
value.append(stats['numTiles'])
value.append(stats['numPlaceables'])
value.append(stats['numBuildingTiles'])
value.append(round(stats['meanTilesActiveCharsNoGuild']))
value.append(round(stats['medianTilesActiveCharsNoGuild']))
value.append(round(stats['meanTilesActiveGuilds']))
value.append(round(stats['medianTilesActiveGuilds']))
value.append(stats['numChars'])
value.append(stats['numActiveChars'])
value.append(stats['numInactiveChars'])
value.append(stats['numLogins'])
value.append(num_lines)
value.append(stats['numRuins'])
value.append(divby10k(stats['totalWealth']))
value.append(round(divby10k(stats['meanActiveCharsWealth']), 4))
value.append(divby10k(stats['medianActiveCharsWealth']))
values = [value]

# Write the statistics to the end of the spreadsheet
logger.debug("Update player statistics sheet.")
sheets = Spreadsheet(PLAYER_SPREADSHEET_ID, activeSheetId=PLAYER_STATISTICS_SHEET_ID)
lastRow = sheets.get_properties()["gridProperties"]["rowCount"]
range = 'Statistics!A' + str(lastRow) + ':Q' + str(lastRow)
sheets.insert_rows(startIndex=lastRow, numRows=1)
sheets.commit()
sheets.update(range=range, values=values)

execTime = datetime.utcnow() - now
execTimeStr = str(execTime.seconds) + "." + str(execTime.microseconds)
if LOG_LEVEL_STDOUT > logging.INFO:
    print(f"Done! Required time: {execTimeStr} sec.")
logger.info(f"Done! Required time: {execTimeStr} sec.")
