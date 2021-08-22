import sys, os
from config import *
from datetime import datetime
from statistics import mean, median
from exiles_api import *
from google_api.sheets import Spreadsheet

# save current time
now = datetime.utcnow()
print("Updating statistics sheet...")

# update the caches
OwnersCache.update(RUINS_CLAN_ID)
ObjectsCache.update(RUINS_CLAN_ID)

# estimate db age by reading the last_login date of the first character in the characters table
if dbAge := db_date():
    dbAgeStr = "Database Date: " + dbAge.strftime("%d-%b-%Y %H:%M UTC")
else:
    execTime = datetime.utcnow() - now
    execTimeStr = str(execTime.seconds) + "." + str(execTime.microseconds)
    print(f"Found no characters in db!\nRequired time: {execTimeStr} sec.")
    sys.exit(0)

# read the number of chatlines
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
        break;

# Get the statistics
stats = Stats.get_tile_statistics(INACTIVITY)
values = [[dbAge.strftime("%d/%m/%Y %H:%M:%S"), stats['numTiles'], stats['numBuildingTiles'], stats['numPlaceables'],
           round(stats['meanTilesActiveCharsNoGuild']), round(stats['medianTilesActiveCharsNoGuild']),
           round(stats['meanTilesActiveGuilds']), round(stats['medianTilesActiveGuilds']),
           stats['numChars'], stats['numActiveChars'], stats['numInactiveChars'],
           stats['numLogins'], num_lines, stats['numRuins']]]

# Write the statistics to the end of the spreadsheet
sheets = Spreadsheet(PLAYER_SPREADSHEET_ID, activeSheetId=PLAYER_STATISTICS_SHEET_ID)
lastRow = sheets.get_properties()["gridProperties"]["rowCount"]
range = 'Statistics!A' + str(lastRow) + ':N' + str(lastRow)
sheets.insert_rows(startIndex=lastRow, numRows=1)
sheets.commit()
sheets.update(range=range, values=values)

execTime = datetime.utcnow() - now
execTimeStr = str(execTime.seconds) + "." + str(execTime.microseconds)
print(f"Done!\nRequired time: {execTimeStr} sec.")
