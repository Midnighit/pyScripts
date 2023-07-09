import sys
import logging
from datetime import datetime
from operator import itemgetter
from exiles_api import db_date, OwnersCache, ObjectsCache, TilesManager, MembersManager
from google_api.sheets import Spreadsheet
from logger import get_logger
from config import (
    RUINS_CLAN_ID, PLAYER_SPREADSHEET_ID, PLAYER_TPM_SHEET_ID, INACTIVITY, BUILDING_TILE_MULT, PLACEBALE_TILE_MULT,
    OWNER_WHITELIST, HIDE_WHITELISTED_OWNERS, ALLOWANCE_INCLUDES_INACTIVES, ALLOWANCE_BASE, ALLOWANCE_CLAN,
    LOG_LEVEL_STDOUT, LOG_LEVEL_FILE, PLACEBALE_TILE_RATIO
)

# catch unhandled exceptions
logger = get_logger('player_tiles_per_member.log', log_level_stdout=LOG_LEVEL_STDOUT, log_level_file=LOG_LEVEL_FILE)


def handle_exception(exc_type, exc_value, exc_traceback):
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return

    logger.error("Uncaught exception", exc_info=(exc_type, exc_value, exc_traceback))


sys.excepthook = handle_exception

# save current time
now = datetime.utcnow()
if LOG_LEVEL_STDOUT > logging.INFO:
    print("Updating tiles per member sheet...")
logger.info("Updating tiles per member sheet...")

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

# instanciate the Spreadsheet object
sheets = Spreadsheet(PLAYER_SPREADSHEET_ID, activeSheetId=PLAYER_TPM_SHEET_ID)

# Create a new list of values to add to the sheet
date_str = now.strftime("%d-%b-%Y %H:%M UTC")
values = []

logger.debug("Gather tiles statistics.")
building_pieces, placeables = TilesManager.get_tiles_by_owner(BUILDING_TILE_MULT, PLACEBALE_TILE_MULT, do_round=False)
logger.debug("Gather member statistics.")
members = MembersManager.get_members(INACTIVITY)

# Compile the list for all guilds
logger.debug("Compile the list for all guilds")
for owner, data in members.items():
    # Whitelisted owners are ignored
    if owner in OWNER_WHITELIST and HIDE_WHITELISTED_OWNERS:
        continue
    # owners with no tiles are not listed
    if owner not in building_pieces or building_pieces[owner] + placeables[owner] == 0:
        continue
    # owners named "Ruins" are not shown
    if data['name'] == "Ruins":
        continue

    if ALLOWANCE_INCLUDES_INACTIVES:
        # allowedTiles is the base allowance + clan allowance per additional member
        allowedTiles = ALLOWANCE_BASE + (data['numMembers'] - 1) * ALLOWANCE_CLAN
        numMembers = data['numMembers']
        memberStr = str(data['numMembers'])
    else:
        # if there are no active players in the clan disregard
        if data['numActiveMembers'] == 0:
            continue
        allowedTiles = ALLOWANCE_BASE + (data['numActiveMembers'] - 1) * ALLOWANCE_CLAN
        numMembers = data['numActiveMembers']
        memberStr = str(data['numActiveMembers']) + ' / ' + str(data['numMembers'])
    # owners with no members are not listed
    if numMembers == 0:
        continue

    values.append([
        data['name'],                                                               # Owner Names
        memberStr,                                                                  # Members
        int(round(building_pieces[owner], 0)),                                      # Building pieces
        int(round(placeables[owner], 0)),                                           # Placeables (adjusted)
        int(round(placeables[owner] / PLACEBALE_TILE_MULT, 0)),                     # Placeables (actual)
        int(round(building_pieces[owner] + placeables[owner], 0)),                  # Tiles (total)
        int(round((building_pieces[owner] + placeables[owner]) / numMembers, 0)),   # Tiles per member
        # Allowance
        int(round(allowedTiles * PLACEBALE_TILE_RATIO, 0)),                         # Placeables (adjusted)
        allowedTiles                                                                # Tites (total)
    ])

# if there are any values, order them by tiles in descending order
logger.debug("Sort values for upload.")
if values:
    values.sort(key=itemgetter(2, 3), reverse=True)
# if there are no values, create a dummy row so freezing top two rows doesn't fail
else:
    values = [['no data', '', '', '', '', '', '', '', '']]

logger.debug("Update tiles per member sheet.")
# generate the headlines and add them to the values list
columnTwoHeader = "Members" if ALLOWANCE_INCLUDES_INACTIVES else "Members (active / total)"
values = [
    [
        'Last Upload: ' + date_str,
        '',
        dbAgeStr,
        '',
        '',
        '',
        '=IMAGE("https://i.imgur.com/v2UO6v5.png",1)',
        'Allowance'
    ], [
        'Owner Names',
        columnTwoHeader,
        'Building pieces',
        'Placeables (adjusted)',
        'Placeables (actual)',
        'Tiles (total)',
        'Tiles per member',
        'Placeables (adjusted)',
        'Tiles (total)'
    ]
] + values

# set the gridsize so it fits in all the values including the two headlines
lastRow = len(values)
sheets.set_grid_size(cols=9, rows=lastRow, frozen=2)
# set a basic filter starting from the second headline going up to the last row
sheets.set_filter(startRowIndex=2)
# merge and format the cells of the first headline
sheets.merge_cells(startColumnIndex=1, endColumnIndex=2, endRowIndex=1)
sheets.merge_cells(startColumnIndex=3, endColumnIndex=4, endRowIndex=1)
sheets.merge_cells(startColumnIndex=8, endColumnIndex=9, endRowIndex=1)
sheets.set_alignment(startColumnIndex=8, endColumnIndex=9, startRowIndex=1, endRowIndex=1, horizontalAlignment='CENTER')
# format the datalines
sheets.set_alignment(startColumnIndex=1, endColumnIndex=1, startRowIndex=3, horizontalAlignment='LEFT')
sheets.set_alignment(startColumnIndex=2, endColumnIndex=9, startRowIndex=3, horizontalAlignment='CENTER')
sheets.set_format(startColumnIndex=3, startRowIndex=3, type='NUMBER', pattern='#,##0')
# update the cells with the values
sheets.update('Tiles!A1:I' + str(lastRow), values)
sheets.commit()

execTime = datetime.utcnow() - now
execTimeStr = str(execTime.seconds) + "." + str(execTime.microseconds)
if LOG_LEVEL_STDOUT > logging.INFO:
    print(f"Done! Required time: {execTimeStr} sec.")
logger.info(f"Done! Required time: {execTimeStr} sec.")
