from config import *
from datetime import datetime
from operator import itemgetter
from exiles_api.model import session, Characters, Guilds
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
    allMembers = tuple((m.name, m.player.disc_user, m.rank_name, m.last_login.strftime("%d-%b-%Y")) for m in members)
    allMemberNames = "\n".join((m[0] + (' (' + m[1] + ')' if m[1] else '') for m in allMembers))
    allMemberRanks = "\n".join((m[2] for m in allMembers))
    allMemberLogin = "\n".join((m[3] for m in allMembers))
    values.append([guild.name, guild.members.last_to_login().player.disc_user, allMemberNames, allMemberRanks, allMemberLogin, memberStr, numTiles, excess, allowedTiles])

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
    fullName = character.name + (' (' + character.player.disc_user + ')' if character.player.disc_user else '')
    values.append([character.name, character.player.disc_user, fullName, character.rank_name, character.last_login.strftime("%d-%b-%Y"), 1, numTiles, excess, allowedTiles])
session.close()

# order the values by tiles in descending order
values.sort(key=itemgetter(6, 7), reverse=True)

# combine headline with the new values
values = [[f"Calendar week: {str(now.isocalendar()[1])} ({date_str})"]] + values
# add enough empty lines at the top of the sheet to contain the new values
lastRow = len(values) + 1
sheets.insert_rows(startIndex=2, numRows=len(values), inheritFromBefore=False)
# format the headline
sheets.merge_cells(startColumnIndex=1, startRowIndex=2, endRowIndex=2)
sheets.set_bg_color(startColumnIndex=1, startRowIndex=2, endRowIndex=2, color="light_yellow")
sheets.set_alignment(startColumnIndex=1, startRowIndex=2, endRowIndex=2, horizontalAlignment='LEFT')
# format the datalines
sheets.set_dimension_group(startIndex=2, endIndex=lastRow, visibility=True)
sheets.set_bg_color(startColumnIndex=1, startRowIndex=3, endRowIndex=lastRow, color="white")
sheets.set_wrap(startColumnIndex=10, endColumnIndex=12, startRowIndex=3, endRowIndex=lastRow, wrapStrategy='WRAP')
sheets.set_format(startColumnIndex=7, endColumnIndex=9, startRowIndex=3, endRowIndex=lastRow, type='NUMBER', pattern='#,##0')
sheets.set_alignment(startColumnIndex=1, endColumnIndex=5, startRowIndex=3, endRowIndex=lastRow, horizontalAlignment = 'LEFT', verticalAlignment = 'MIDDLE')
sheets.set_alignment(startColumnIndex=6, endColumnIndex=9, startRowIndex=3, endRowIndex=lastRow, horizontalAlignment = 'CENTER', verticalAlignment = 'MIDDLE')
sheets.set_alignment(startColumnIndex=10, endColumnIndex=12, startRowIndex=3, endRowIndex=lastRow, horizontalAlignment = 'LEFT', verticalAlignment = 'MIDDLE')
# update the newly inserted cells with the values
sheets.update('Tiles!A2:J' + str(lastRow), values)
sheets.commit()
