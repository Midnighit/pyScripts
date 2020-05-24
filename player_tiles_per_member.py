from config import *
from datetime import datetime
from operator import itemgetter
from exiles_api.model import session, db_date, Characters, Guilds
from google_api.sheets import Spreadsheet

# save current time
now = datetime.utcnow()
print("Updating tiles per member sheet...")

# estimate db age by reading the last_login date of the first character in the characters table
if dbAge := db_date():
    dbAgeStr = "Database Date: " + dbAge.strftime("%d-%b-%Y %H:%M UTC")
else:
    execTime = datetime.utcnow() - now
    execTimeStr = str(execTime.seconds) + "." + str(execTime.microseconds)
    print(f"Found no characters in db!\nRequired time: {execTimeStr} sec.")
    sys.exit(0)

# instanciate the Spreadsheet object
sheets = Spreadsheet(PLAYER_SPREADSHEET_ID, activeSheetId=PLAYER_TPM_SHEET_ID)

# Create a new list of values to add to the sheet
date_str = now.strftime("%d-%b-%Y %H:%M UTC")
values = []
# Compile the list for all guilds
for guild in session.query(Guilds).all():
    # Whitelisted owners are ignored
    if guild.id in OWNER_WHITELIST and HIDE_WHITELISTED_OWNERS:
        continue
    # Number of tiles taking bMult and pMult into account
    numTiles = guild.num_tiles(bMult=BUILDING_TILE_MULT, pMult=PLACEBALE_TILE_MULT)
    if numTiles == 0:
        continue
    # list of guild members
    members = guild.members
    if ALLOWANCE_INCLUDES_INACTIVES:
        # allowedTiles is the base allowance + clan allowance per additional member
        allowedTiles = ALLOWANCE_BASE + (len(members) - 1) * ALLOWANCE_CLAN
        numMembers = len(members)
        memberStr = str(numMembers)
    else:
        # if there are no active players in the clan disregard
        if len(members.active(INACTIVITY)) == 0:
            continue
        allowedTiles = ALLOWANCE_BASE + (len(members.active(INACTIVITY)) - 1) * ALLOWANCE_CLAN
        numMembers = len(members.active(INACTIVITY))
        memberStr = str(numMembers) + ' / ' + str(len(members))
    tpm = numTiles / numMembers
    values.append([guild.name, memberStr, numTiles, tpm, allowedTiles])

# Compile the list for all characters
for character in session.query(Characters).all():
    # Whitelisted owners are ignored
    if character.id in OWNER_WHITELIST and HIDE_WHITELISTED_OWNERS:
        continue
    # Discard characters that are in a guild or inactive
    if character.has_guild or character.is_inactive(INACTIVITY):
        continue
    # Number of tiles taking bMult and pMult into account
    numTiles = character.num_tiles(bMult=BUILDING_TILE_MULT, pMult=PLACEBALE_TILE_MULT)
    if numTiles == 0:
        continue
    values.append([character.name, '1 / 1', numTiles, numTiles, ALLOWANCE_BASE])

# order the values by tiles in descending order
x = values.sort(key=itemgetter(2, 3), reverse=True)

# generate the headlines and add them to the values list
columnTwoHeader = "Members" if ALLOWANCE_INCLUDES_INACTIVES else "Members (active / total)"
values = [['Last Upload: ' + date_str, '', dbAgeStr],
          ['Owner Names', columnTwoHeader, 'Tiles', 'Tiles per member', 'Allowance']] + values

# set the gridsize so it fits in all the values including the two headlines
lastRow = len(values)
sheets.set_grid_size(cols=5, rows=lastRow, frozen=2)
# set a basic filter starting from the second headline going up to the last row
sheets.set_filter(startColumnIndex=1, endColumnIndex=5, startRowIndex=2, endRowIndex=lastRow)
# merge the cells of the first headline
sheets.merge_cells(startColumnIndex=1, endColumnIndex=2)
sheets.merge_cells(startColumnIndex=3, endColumnIndex=4)
# format the datalines
sheets.set_alignment(startColumnIndex=1, endColumnIndex=1, startRowIndex=3, endRowIndex=lastRow, horizontalAlignment = 'LEFT')
sheets.set_alignment(startColumnIndex=2, endColumnIndex=5, startRowIndex=3, endRowIndex=lastRow, horizontalAlignment = 'CENTER')
sheets.set_format(startColumnIndex=3, endColumnIndex=5, startRowIndex=3, endRowIndex=lastRow, type='NUMBER', pattern='#,##0')
# update the cells with the values
sheets.update('Tiles!A1:E' + str(lastRow), values)
sheets.commit()

execTime = datetime.utcnow() - now
execTimeStr = str(execTime.seconds) + "." + str(execTime.microseconds)
print(f"Done!\nRequired time: {execTimeStr} sec.")
