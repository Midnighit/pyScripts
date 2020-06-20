from config import *
from datetime import datetime
from operator import itemgetter
from exiles_api import *
from google_api.sheets import Spreadsheet

# save current time
now = datetime.utcnow()
print("Updating tiles per member sheet...")

# update the caches
OwnersCache.update()
ObjectsCache.update()

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

tm = TilesManager()
tiles = tm.get_tiles_by_owner(BUILDING_TILE_MULT, PLACEBALE_TILE_MULT)
mm = MembersManager()
members = mm.get_members(INACTIVITY)

# Compile the list for all guilds
for owner, data in members.items():
    # Whitelisted owners are ignored
    if owner in OWNER_WHITELIST and HIDE_WHITELISTED_OWNERS:
        continue
    # owners with no buildings are not listed
    if tiles[owner] == 0:
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
    tpm = tiles[owner] / numMembers
    values.append([data['name'], memberStr, tiles[owner], tpm, allowedTiles])

# order the values by tiles in descending order
values.sort(key=itemgetter(2, 3), reverse=True)

# generate the headlines and add them to the values list
columnTwoHeader = "Members" if ALLOWANCE_INCLUDES_INACTIVES else "Members (active / total)"
values = [['Last Upload: ' + date_str, '', dbAgeStr],
          ['Owner Names', columnTwoHeader, 'Tiles', 'Tiles per member', 'Allowance']] + values

# set the gridsize so it fits in all the values including the two headlines
lastRow = len(values)
sheets.set_grid_size(cols=5, rows=lastRow, frozen=2)
# set a basic filter starting from the second headline going up to the last row
sheets.set_filter(startRowIndex=2)
# merge the cells of the first headline
sheets.merge_cells(endRowIndex=1, endColumnIndex=2)
sheets.merge_cells(startColumnIndex=3, endRowIndex=1, endColumnIndex=4)
# format the datalines
sheets.set_alignment(startRowIndex=3, endColumnIndex=1, horizontalAlignment = 'LEFT')
sheets.set_alignment(startColumnIndex=2, startRowIndex=3, horizontalAlignment = 'CENTER')
sheets.set_format(startColumnIndex=3, startRowIndex=3, type='NUMBER', pattern='#,##0')
# update the cells with the values
sheets.update('Tiles!A1:E' + str(lastRow), values)
sheets.commit()

execTime = datetime.utcnow() - now
execTimeStr = str(execTime.seconds) + "." + str(execTime.microseconds)
print(f"Done!\nRequired time: {execTimeStr} sec.")
