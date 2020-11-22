from config import *
from datetime import datetime
from operator import itemgetter
from exiles_api import session, Characters, Guilds
from google_api.sheets import Spreadsheet

# instanciate the Spreadsheet object
sheets = Spreadsheet(TILES_MGMT_SPREADSHEET_ID, activeSheetId=TILES_MGMT_SHEET_ID)

# Create a new list of values to add to the sheet
now = datetime.utcnow()
date_str = now.strftime("%d-%b-%Y")
values = []
# Compile the list for all guilds
for guild in session.query(Guilds).all():
    # Whitelisted owners are ignored
    if guild.id in OWNER_WHITELIST:
        continue
    # Ruins are ignored
    if guild.name == 'Ruins':
        continue
    # Number of tiles taking bMult and pMult into account
    numTiles = guild.num_tiles(bMult=BUILDING_TILE_MULT, pMult=PLACEBALE_TILE_MULT)
    # list of guild members
    members = guild.members
    if ALLOWANCE_INCLUDES_INACTIVES:
        # allowedTiles is the base allowance + clan allowance per additional member
        allowedTiles = ALLOWANCE_BASE + (len(members) - 1) * ALLOWANCE_CLAN
        memberStr = len(members)
    else:
        # if there are no active players in the clan disregard
        if len(members.active(INACTIVITY)) == 0:
            continue
        allowedTiles = ALLOWANCE_BASE + (len(members.active(INACTIVITY)) - 1) * ALLOWANCE_CLAN
        memberStr = str(len(members.active(INACTIVITY))) + ' / ' + str(len(members))
    # if allowedTiles is greater than the absolute amount of tiles of the guild disregard
    if allowedTiles >= numTiles:
        continue
    excess = numTiles - allowedTiles

    allMembers = tuple((m.name, (m.user.disc_user if m.user else ''), m.rank_name, m.last_login.strftime("%d-%b-%Y")) for m in members)
    allMemberNames = "\n".join((m[0] + (' (' + m[1] + ')' if m[1] else '') for m in allMembers))
    allMemberRanks = "\n".join((m[2] for m in allMembers))
    allMemberLogin = "\n".join((m[3] for m in allMembers))
    disc_user = guild.members.last_to_login.user.disc_user if guild.members.last_to_login and guild.members.last_to_login.user else ''
    values.append([guild.name, disc_user, allMemberNames, allMemberRanks, allMemberLogin, memberStr, numTiles, excess, allowedTiles])

# Compile the list for all characters
for character in session.query(Characters).all():
    # Whitelisted owners are ignored
    if character.id in OWNER_WHITELIST:
        continue
    # Discard characters that are in a guild or inactive
    if character.has_guild or character.is_inactive(INACTIVITY):
        continue
    # Number of tiles taking bMult and pMult into account
    numTiles = character.num_tiles(bMult=BUILDING_TILE_MULT, pMult=PLACEBALE_TILE_MULT)
    allowedTiles = ALLOWANCE_BASE
    # if allowedTiles is greater than the absolute amount of tiles of the guild disregard
    if allowedTiles >= numTiles:
        continue
    excess = numTiles - allowedTiles
    fullName = character.name + (' (' + character.user.disc_user + ')' if character.user and character.user.disc_user else '')
    disc_user = character.user.disc_user if character.user else ''
    values.append([character.name, disc_user, fullName, character.rank_name, character.last_login.strftime("%d-%b-%Y"), 1, numTiles, excess, allowedTiles])
session.close()

# order the values by tiles in descending order or add a token line if values are empty
if len(values) > 0:
    values.sort(key=itemgetter(6, 7), reverse=True)
else:
    values.append(["Nobody was over their tiles limit this week - yay!", '', '', '', '', '', '', '', ''])

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
sheets.set_wrap(startColumnIndex=10, endColumnIndex=12, startRowIndex=3, endRowIndex=lastRow, wrapStrategy='WRAP')
sheets.set_format(startColumnIndex=7, endColumnIndex=9, startRowIndex=3, endRowIndex=lastRow, type='NUMBER', pattern='#,##0')
sheets.set_alignment(endColumnIndex=5, startRowIndex=3, endRowIndex=lastRow, horizontalAlignment = 'LEFT', verticalAlignment = 'MIDDLE')
sheets.set_alignment(startColumnIndex=6, endColumnIndex=9, startRowIndex=3, endRowIndex=lastRow, horizontalAlignment = 'CENTER', verticalAlignment = 'MIDDLE')
sheets.set_alignment(startColumnIndex=10, startRowIndex=3, endRowIndex=lastRow, horizontalAlignment = 'LEFT', verticalAlignment = 'MIDDLE')
# update the newly inserted cells with the values
sheets.commit()
sheets.update('Tiles!A2:J' + str(lastRow), values)
