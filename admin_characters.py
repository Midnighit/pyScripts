from config import *
from datetime import datetime
from exiles_api import *
from google_api.sheets import Spreadsheet

# save current time
now = datetime.utcnow()
print("Updating characters sheet...")

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

# instanciate the Spreadsheet object
sheets = Spreadsheet(ADMIN_SPREADSHEET_ID, activeSheetId=ADMIN_CHARACTERS_SHEET_ID)

# Create a new list of values to add to the sheet
date_str = now.strftime("%d-%b-%Y %H:%M UTC")
values = []

for c in session.query(Characters).order_by(Characters._last_login.desc()).all():
    guild_name = c.guild.name if c.guild else ''
    guild_id = c.guild.id if c.guild else ''
    disc_user = c.user.disc_user if c.user and c.user.disc_user else ''
    disc_id = c.user.disc_id if c.user and c.user.disc_id else ''
    values.append([c.name, c.id, guild_name, guild_id, c.level, disc_user, disc_id, c.account.funcom_id, c.slot, c.last_login.strftime("%d-%b-%Y %H:%M")])

# generate the headlines and add them to the values list
values = [['Last Upload: ' + date_str, '', dbAgeStr],
          ['Character Names', 'CharacterID', 'Guild Names', 'GuildID', 'lvl', 'Discord Name', 'DiscordID', 'FuncomID', 'Slot', 'Last Login (UTC)']] + values

# set the gridsize so it fits in all the values including the two headlines
lastRow = len(values)
sheets.set_grid_size(cols=10, rows=lastRow, frozen=2)
# set a basic filter starting from the second headline going up to the last row
sheets.set_filter(startRowIndex=2)
# format the datalines
sheets.set_alignment(startRowIndex=3, endColumnIndex=4, horizontalAlignment = 'LEFT')
sheets.set_alignment(startRowIndex=3, startColumnIndex=5, endColumnIndex=5, horizontalAlignment = 'RIGHT')
sheets.set_alignment(startRowIndex=3, startColumnIndex=6, endColumnIndex=9, horizontalAlignment = 'LEFT')
sheets.set_alignment(startRowIndex=3, startColumnIndex=10, endColumnIndex=10, horizontalAlignment = 'RIGHT')
sheets.set_format(startRowIndex=3, startColumnIndex=10, type='DATE', pattern='dd-mmm-yyyy hh:mm')
# update the cells with the values
sheets.update('Characters!A1:J' + str(lastRow), values)
sheets.commit()

execTime = datetime.utcnow() - now
execTimeStr = str(execTime.seconds) + "." + str(execTime.microseconds)
print(f"Done!\nRequired time: {execTimeStr} sec.")
