import sys
import logging
from datetime import datetime
from exiles_api import db_date, session, Guilds, Properties
from google_api.sheets import Spreadsheet
from logger import get_logger
from config import (
    ADMIN_SPREADSHEET_ID, ADMIN_CLANS_SHEET_ID, LOG_LEVEL_STDOUT, LOG_LEVEL_FILE, INACTIVITY, BUILDING_TILE_MULT,
    PLACEBALE_TILE_MULT
)

# catch unhandled exceptions
logger = get_logger('admin_clans.log', log_level_stdout=LOG_LEVEL_STDOUT, log_level_file=LOG_LEVEL_FILE)


def handle_exception(exc_type, exc_value, exc_traceback):
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return

    logger.error("Uncaught exception", exc_info=(exc_type, exc_value, exc_traceback))


sys.excepthook = handle_exception

# save current time
now = datetime.utcnow()
if LOG_LEVEL_STDOUT > logging.INFO:
    print("Updating clans sheet...")
logger.info("Updating clans sheet...")

# estimate db age by reading the last_login date of the first character in the characters table
if dbAge := db_date():
    dbAgeStr = "Database Date: " + dbAge.strftime("%d-%b-%Y %H:%M UTC")
else:
    execTime = datetime.utcnow() - now
    execTimeStr = str(execTime.seconds) + "." + str(execTime.microseconds)
    logger.info(f"Found no characters in db!\nRequired time: {execTimeStr} sec.")
    sys.exit(0)

# instanciate the Spreadsheet object
sheets = Spreadsheet(ADMIN_SPREADSHEET_ID, activeSheetId=ADMIN_CLANS_SHEET_ID)

# Create a new list of values to add to the sheet
date_str = now.strftime("%d-%b-%Y %H:%M UTC")
values = []


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


logger.debug("Compiling the character data.")
for g in session.query(Guilds).all():
    bronze_chars = Properties.get_pippi_money(guild_id=g.id, with_thespians=False, as_number=True)
    bronze_no_chars = Properties.get_pippi_money(guild_id=g.id, with_chars=False, as_number=True)
    # try to exclude admin/support chars with access to the cheat menu from the statistics

    values.append([
                    g.name,
                    g.id,
                    f"{len(g.members.active(INACTIVITY))} / {len(g.members)}",
                    divby10k(bronze_chars),
                    divby10k(bronze_no_chars),
                    g.num_tiles(BUILDING_TILE_MULT, PLACEBALE_TILE_MULT),
                    g.last_login.strftime("%d-%b-%Y %H:%M") if g.last_login else ''
                ])

# sort by last login with the last one being first in the list
values.sort(key=lambda x: x[6], reverse=True)

# generate the headlines and add them to the values list
values = [
            [
                'Last Upload: ' + date_str,
                '',
                dbAgeStr
            ],
            [
                'Guild Names',
                'GuildID',
                'Members (active / all)',
                'Pippi gold (chars)',
                'Pippi gold (thespians)',
                'Tiles',
                'Last Login (UTC)'
            ]
        ] + values

logger.debug("Uploading results to google sheet.")
# set the gridsize so it fits in all the values including the two headlines
lastRow = len(values)
sheets.set_grid_size(cols=7, rows=lastRow, frozen=2)
# set a basic filter starting from the second headline going up to the last row
sheets.set_filter(startRowIndex=2)
# format the datalines
sheets.set_alignment(startRowIndex=3, endColumnIndex=6, horizontalAlignment='LEFT')
sheets.set_alignment(startRowIndex=3, startColumnIndex=7, horizontalAlignment='RIGHT')
sheets.set_format(startRowIndex=3, startColumnIndex=7, type='DATE', pattern='dd-mmm-yyyy hh:mm')
# update the cells with the values
sheets.commit()
sheets.update('Clans!A1:G' + str(lastRow), values)

execTime = datetime.utcnow() - now
execTimeStr = str(execTime.seconds) + "." + str(execTime.microseconds)
if LOG_LEVEL_STDOUT > logging.INFO:
    print(f"Done! Required time: {execTimeStr} sec.")
logger.info(f"Done! Required time: {execTimeStr} sec.")
