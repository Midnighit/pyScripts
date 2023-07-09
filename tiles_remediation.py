import sys
import logging
from datetime import datetime
from operator import itemgetter
from exiles_api import session, Characters, Guilds, TilesManager
from google_api.sheets import Spreadsheet
from logger import get_logger
from config import (
    TILES_MGMT_SPREADSHEET_ID, TILES_MGMT_SHEET_ID, OWNER_WHITELIST, BUILDING_TILE_MULT, PLACEBALE_TILE_MULT,
    ALLOWANCE_INCLUDES_INACTIVES, ALLOWANCE_BASE, ALLOWANCE_CLAN, INACTIVITY, LOG_LEVEL_STDOUT, LOG_LEVEL_FILE,
    PLACEBALE_TILE_RATIO
)

# catch unhandled exceptions
logger = get_logger('tiles_remediation.log', log_level_stdout=LOG_LEVEL_STDOUT, log_level_file=LOG_LEVEL_FILE)


def handle_exception(exc_type, exc_value, exc_traceback):
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return

    logger.error("Uncaught exception", exc_info=(exc_type, exc_value, exc_traceback))


sys.excepthook = handle_exception

# save current time
now = datetime.utcnow()
if LOG_LEVEL_STDOUT > logging.INFO:
    print("Updating tiles remediation sheet...")
logger.info("Updating tiles remediation sheet...")

# instanciate the Spreadsheet object
sheets = Spreadsheet(TILES_MGMT_SPREADSHEET_ID, activeSheetId=TILES_MGMT_SHEET_ID)

# Create a new list of values to add to the sheet
date_str = now.strftime("%d-%b-%Y")
values = []
logger.debug("Gather tiles statistics.")
building_pieces, placeables = TilesManager.get_tiles_by_owner(BUILDING_TILE_MULT, PLACEBALE_TILE_MULT, do_round=False)
# Compile the list for all guilds
for guild in session.query(Guilds).all():
    # Whitelisted owners are ignored
    if guild.id in OWNER_WHITELIST:
        continue
    # Ruins are ignored
    if guild.name == 'Ruins':
        continue
    # Discard guilds with no tiles
    if guild.id not in building_pieces:
        continue
    # building pices
    buildingPieces = building_pieces[guild.id]
    # the adjusted number of placeables
    placeablesAdjusted = placeables[guild.id]
    # Number of tiles taking bMult and pMult into account
    totalTiles = buildingPieces + placeablesAdjusted
    # list of guild members
    members = guild.members
    if ALLOWANCE_INCLUDES_INACTIVES:
        # allowedTilesTotal is the base allowance + clan allowance per additional member
        allowedTilesTotal = ALLOWANCE_BASE + (len(members) - 1) * ALLOWANCE_CLAN
        memberStr = len(members)
    else:
        # if there are no active players in the clan disregard
        if len(members.active(INACTIVITY)) == 0:
            continue
        allowedTilesTotal = ALLOWANCE_BASE + (len(members.active(INACTIVITY)) - 1) * ALLOWANCE_CLAN
        memberStr = str(len(members.active(INACTIVITY))) + ' / ' + str(len(members))

    # allowedPlaceables is a ratio of the total tiles
    allowedPlaceables = allowedTilesTotal * PLACEBALE_TILE_RATIO
    # if none of the allowances are exceeded disregard the guild
    if allowedTilesTotal >= round(totalTiles, 0) and allowedPlaceables >= round(placeablesAdjusted, 0):
        continue
    excessTotal = max(totalTiles - allowedTilesTotal, 0)
    excessPlaceables = max(placeablesAdjusted - allowedPlaceables, 0)

    allMembers = tuple(
        (m.name, (m.user.disc_user if m.user else ''), m.rank_name, m.last_login.strftime("%d-%b-%Y")) for m in members
    )
    allMemberNames = "\n".join((m[0] + (' (' + m[1] + ')' if m[1] else '') for m in allMembers))
    allMemberRanks = "\n".join((m[2] for m in allMembers))
    allMemberLogin = "\n".join((m[3] for m in allMembers))
    if guild.members.last_to_login and guild.members.last_to_login.user:
        disc_user = guild.members.last_to_login.user.disc_user
    else:
        disc_user = ''
    values.append([
        guild.name,                         # Owner
        disc_user,                          # Discord Name (last to login)
        allMemberNames,                     # (Char Name Discord Name)
        allMemberRanks,                     # (Rank)
        allMemberLogin,                     # (Last Login)
        memberStr,                          # Members (active / total)
        int(round(buildingPieces, 0)),      # Building Pieces
        int(round(placeablesAdjusted, 0)),  # Placeables (adjusted)
        int(round(totalTiles, 0)),          # Tiles (total)
        int(round(excessPlaceables, 0)),    # (excess placeables)
        int(round(excessTotal, 0)),         # (excess total)
        int(round(allowedPlaceables, 0)),   # (allowance placeables)
        allowedTilesTotal                   # (allowance total)
    ])

# Compile the list for all characters
logger.debug("Compile character tiles statistics.")
for character in session.query(Characters).all():
    # Whitelisted owners are ignored
    if character.id in OWNER_WHITELIST:
        continue
    # Discard characters that are in a guild or inactive
    if character.has_guild or character.is_inactive(INACTIVITY):
        continue
    # Discard characters with no tiles
    if character.id not in building_pieces:
        continue
    # building pices
    buildingPieces = building_pieces[character.id]
    # the adjusted number of placeables
    placeablesAdjusted = placeables[character.id]
    # Number of tiles taking bMult and pMult into account
    totalTiles = buildingPieces + placeablesAdjusted
    # allowedPlaceables is a ratio of the total tiles
    allowedPlaceables = round(ALLOWANCE_BASE * PLACEBALE_TILE_RATIO, 0)
    # if none of the allowances are exceeded disregard the character
    if ALLOWANCE_BASE >= round(totalTiles, 0) and allowedPlaceables >= round(placeablesAdjusted, 0):
        continue
    excessTotal = max(totalTiles - ALLOWANCE_BASE, 0)
    excessPlaceables = max(placeablesAdjusted - allowedPlaceables, 0)

    fullName = character.name + (
        ' (' + character.user.disc_user + ')' if character.user and character.user.disc_user else ''
    )
    disc_user = character.user.disc_user if character.user else ''
    values.append([
        character.name,                             # Owner
        disc_user,                                  # Discord Name (last to login)
        fullName,                                   # (Char Name Discord Name)
        character.rank_name,                        # (Rank)
        character.last_login.strftime("%d-%b-%Y"),  # (Last Login)
        1,                                          # Members (active / total)
        int(round(totalTiles, 0)),                  # Building Pieces
        int(round(placeablesAdjusted, 0)),          # Placeables (adjusted)
        int(round(totalTiles, 0)),                  # Tiles (total)
        int(round(excessPlaceables, 0)),            # (excess placeables)
        int(round(excessTotal, 0)),                 # (excess total)
        int(round(allowedPlaceables, 0)),           # (allowance placeables)
        ALLOWANCE_BASE                              # (allowance total)
    ])
session.close()

# order the values by tiles in descending order or add a token line if values are empty
logger.debug("Sort values for upload.")
if len(values) > 0:
    values.sort(key=itemgetter(6, 7), reverse=True)
else:
    values.append(["Nobody was over their tiles limit this week - yay!"] + [''] * 12)

logger.debug("Format and upload data to tiles remediation sheet.")
# combine headline with the new values
values = [[f"Calendar week: {str(now.isocalendar()[1])} ({date_str})"]] + values
# add enough empty lines at the top of the sheet to contain the new values
lastRow = len(values) + 1
sheets.insert_rows(startIndex=2, numRows=len(values), inheritFromBefore=False)
# format the headline
sheets.merge_cells(startRowIndex=2, endRowIndex=2)
sheets.set_bg_color(startRowIndex=2, endRowIndex=2, color="light_yellow")
sheets.set_alignment(startRowIndex=2, endRowIndex=2, horizontalAlignment='LEFT')
# format the datalines
sheets.set_dimension_group(startIndex=2, endIndex=lastRow, hidden=True)
sheets.set_bg_color(startRowIndex=3, endRowIndex=lastRow, color="white")
sheets.set_wrap(startColumnIndex=14, endColumnIndex=16, startRowIndex=3, endRowIndex=lastRow, wrapStrategy='WRAP')
sheets.set_format(
    startColumnIndex=7, endColumnIndex=13, startRowIndex=3, endRowIndex=lastRow,
    type='NUMBER', pattern='#,##0'
)
sheets.set_alignment(
    endColumnIndex=5, startRowIndex=3, endRowIndex=lastRow,
    horizontalAlignment='LEFT', verticalAlignment='MIDDLE'
)
sheets.set_alignment(
    startColumnIndex=6, endColumnIndex=13, startRowIndex=3, endRowIndex=lastRow,
    horizontalAlignment='CENTER', verticalAlignment='MIDDLE'
)
sheets.set_alignment(
    startColumnIndex=14, startRowIndex=3, endRowIndex=lastRow,
    horizontalAlignment='LEFT', verticalAlignment='MIDDLE'
)
# update the newly inserted cells with the values
sheets.commit()
sheets.update('Tiles!A2:N' + str(lastRow), values)

execTime = datetime.utcnow() - now
execTimeStr = str(execTime.seconds) + "." + str(execTime.microseconds)
if LOG_LEVEL_STDOUT > logging.INFO:
    print(f"Done! Required time: {execTimeStr} sec.")
logger.info(f"Done! Required time: {execTimeStr} sec.")
