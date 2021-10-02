import sys
from datetime import datetime
from operator import itemgetter
from math import ceil, inf
from config import ADMIN_SPREADSHEET_ID, ADMIN_TPM_SHEET_ID, BUILDING_TILE_MULT
from config import PLACEBALE_TILE_MULT, INACTIVITY, RUINS_CLAN_ID, ALLOWANCE_INCLUDES_INACTIVES
from exiles_api import db_date, session, TilesManager, MembersManager, OwnersCache
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
sheets = Spreadsheet(ADMIN_SPREADSHEET_ID, activeSheetId=ADMIN_TPM_SHEET_ID)

# Create a new list of values to add to the sheet
date_str = now.strftime("%d-%b-%Y %H:%M UTC")
values = []

tiles = TilesManager.get_tiles_consolidated(BUILDING_TILE_MULT, PLACEBALE_TILE_MULT)
members = MembersManager.get_members(INACTIVITY)

for object_id, ctd in tiles.items():
    # owner_id 0 is a game reserved id and can be ignored
    if ctd['owner_id'] == 0:
        continue
    # ensure that owner_id is in members
    if ctd['owner_id'] in members:
        name = members[ctd['owner_id']]['name']
        if name == 'Ruins' and ctd['owner_id'] != RUINS_CLAN_ID:
            owner = session.query(OwnersCache).get(ctd['owner_id'])
            if owner:
                name = owner.name + " (Ruins)"
    else:
        print("should never get here!")
        print(f"object_id: {object_id} / contents: {ctd}")
        print("Skipping object")
        continue

    if ALLOWANCE_INCLUDES_INACTIVES:
        num_members = num_members_str = members[ctd['owner_id']]['numMembers']
    else:
        num_members = members[ctd['owner_id']]['numActiveMembers']
        num_members_str = f"({num_members} / {members[ctd['owner_id']]['numMembers']})"
    if num_members > 0:
        tpm = round(ctd['sum_tiles'] / num_members)
    else:
        tpm = inf
    location = f"TeleportPlayer {ceil(ctd['x'])} {ceil(ctd['y'])} {ceil(ctd['z'])}"
    values.append([object_id, name, ctd['owner_id'], ctd['tiles'], num_members_str, tpm, ctd['class'], location])

# order the values by tiles in descending order
values.sort(key=itemgetter(3), reverse=True)
values.sort(key=itemgetter(2))
values.sort(key=itemgetter(5), reverse=True)

# generate the headlines and add them to the values list
columnTwoHeader = "Members" if ALLOWANCE_INCLUDES_INACTIVES else "Members (active / total)"
values = [
    ['Last Upload: ' + date_str, '', dbAgeStr],
    [
        'Object ID',
        'Owner Names',
        'Owner ID',
        'Tiles',
        columnTwoHeader,
        'Tiles per member',
        'Item Class',
        'Location'
    ]
] + values

# set the gridsize so it fits in all the values including the two headlines
lastRow = len(values)
sheets.set_grid_size(cols=8, rows=lastRow, frozen=2)
# ungroup everything
sheets.delete_dimension_group()
sheets.set_visibility(hidden=False)
# replace inf for first row
if values[2][5] == inf:
    values[2][5] = '∞'
# initialize some loop vars
prev_id = values[2][2]
prev_row = 2
multiline = False
# re-group by owner_ids
for row in range(3, lastRow):
    # compare owner_id with that of last row(s)
    if values[row][2] == prev_id:
        multiline = True
    else:
        if multiline:
            sheets.set_dimension_group(startIndex=prev_row + 1, endIndex=row, hidden=True)
        multiline = False
        prev_id = values[row][2]
        prev_row = row
    # replace all occurrences of math.inf with ∞
    if values[row][5] == inf:
        values[row][5] = '∞'
# ensure that last row is being grouped too if applicable
if multiline:
    sheets.set_dimension_group(startIndex=prev_row + 1, hidden=True)

# set a basic filter starting from the second headline going up to the last row
sheets.set_filter(startRowIndex=2)
# merge the cells of the first headline
sheets.merge_cells(endRowIndex=1, endColumnIndex=2)
sheets.merge_cells(startColumnIndex=3, endColumnIndex=5, endRowIndex=1)
# format the datalines
sheets.set_alignment(startRowIndex=3, endColumnIndex=3, horizontalAlignment='LEFT')
sheets.set_alignment(startRowIndex=3, startColumnIndex=4, endColumnIndex=4, horizontalAlignment='RIGHT')
sheets.set_alignment(startRowIndex=3, startColumnIndex=5, endColumnIndex=5, horizontalAlignment='CENTER')
sheets.set_alignment(startRowIndex=3, startColumnIndex=6, endColumnIndex=6, horizontalAlignment='RIGHT')
sheets.set_alignment(startRowIndex=3, startColumnIndex=7, endColumnIndex=8, horizontalAlignment='LEFT')
sheets.set_format(startColumnIndex=4, endColumnIndex=4, startRowIndex=3, type='NUMBER', pattern='#,##0')
sheets.set_format(startColumnIndex=6, endColumnIndex=6, startRowIndex=3, type='NUMBER', pattern='#,##0')
# update the cells with the values
sheets.commit()
sheets.update('Tiles per member!A1:H' + str(lastRow), values)
execTime = datetime.utcnow() - now
execTimeStr = str(execTime.seconds) + "." + str(execTime.microseconds)
print(f"Done!\nRequired time: {execTimeStr} sec.")
