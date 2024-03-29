import sys
import logging
from statistics import median, mean
from datetime import datetime
from exiles_api import db_date, session, Characters, Guilds, Properties
from google_api.sheets import Spreadsheet
from logger import get_logger
from config import ADMIN_SPREADSHEET_ID, ADMIN_CHARACTERS_SHEET_ID, LOG_LEVEL_STDOUT, LOG_LEVEL_FILE

# catch unhandled exceptions
logger = get_logger('admin_characters.log', log_level_stdout=LOG_LEVEL_STDOUT, log_level_file=LOG_LEVEL_FILE)


def handle_exception(exc_type, exc_value, exc_traceback):
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return

    logger.error("Uncaught exception", exc_info=(exc_type, exc_value, exc_traceback))


sys.excepthook = handle_exception

# save current time
now = datetime.utcnow()
if LOG_LEVEL_STDOUT > logging.INFO:
    print("Updating characters sheet...")
logger.info("Updating characters sheet...")

# estimate db age by reading the last_login date of the first character in the characters table
if dbAge := db_date():
    dbAgeStr = "Database Date: " + dbAge.strftime("%d-%b-%Y %H:%M UTC")
else:
    execTime = datetime.utcnow() - now
    execTimeStr = str(execTime.seconds) + "." + str(execTime.microseconds)
    logger.info(f"Found no characters in db!\nRequired time: {execTimeStr} sec.")
    sys.exit(0)

# instanciate the Spreadsheet object
sheets = Spreadsheet(ADMIN_SPREADSHEET_ID, activeSheetId=ADMIN_CHARACTERS_SHEET_ID)

# Create a new list of values to add to the sheet
date_str = now.strftime("%d-%b-%Y %H:%M UTC")
values = []
# Create a list of how much Pippi money each character has for additional statistics
wealth = []


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
for c in session.query(Characters).order_by(Characters._last_login.desc()).all():
    guild_name = c.guild.name if c.guild else ''
    guild_id = c.guild.id if c.guild else ''
    disc_user = c.user.disc_user if c.user and c.user.disc_user else ''
    disc_id = c.user.disc_id if c.user and c.user.disc_id else ''
    bronze = Properties.get_pippi_money(character_id=c.id, as_number=True)
    # try to exclude admin/support chars with access to the cheat menu from the statistics
    if c.slot == 'active' or c.slot in ('1', '2'):
        wealth.append(bronze)

    values.append([
                    c.name,
                    c.id,
                    guild_name,
                    guild_id,
                    c.level,
                    divby10k(bronze),
                    disc_user,
                    disc_id,
                    c.account.funcom_id,
                    c.slot,
                    c.last_login.strftime("%d-%b-%Y %H:%M")
                ])

guild_wealth = 0
for g in session.query(Guilds).all():
    guild_wealth += Properties.get_pippi_money(guild_id=g.id, with_chars=False, as_number=True)

total = divby10k(sum(wealth) + guild_wealth)
# generate the headlines and add them to the values list
values = [
            [
                'Last Upload: ' + date_str,
                '',
                dbAgeStr,
                '',
                '',
                (
                    f"Total Pippi gold: {divby10k(sum(wealth) + guild_wealth):,}  / "
                    f"Avrg Pippi gold per character: {divby10k(round(mean(wealth))):,}  / "
                    f"Median Pippi gold per character: {divby10k(round(median(wealth), 1)):,}"
                )
            ],
            [
                'Character Names',
                'CharacterID',
                'Guild Names',
                'GuildID',
                'lvl',
                'Pippi gold',
                'Discord Name',
                'DiscordID',
                'FuncomID',
                'Slot',
                'Last Login (UTC)'
            ]
        ] + values

logger.debug("Uploading results to google sheet.")
# set the gridsize so it fits in all the values including the two headlines
lastRow = len(values)
sheets.set_grid_size(cols=11, rows=lastRow, frozen=2)
# set a basic filter starting from the second headline going up to the last row
sheets.set_filter(startRowIndex=2)
# format the datalines
sheets.set_alignment(startRowIndex=3, endColumnIndex=4, horizontalAlignment='LEFT')
sheets.set_alignment(startRowIndex=3, startColumnIndex=5, endColumnIndex=5, horizontalAlignment='RIGHT')
sheets.set_alignment(startRowIndex=3, startColumnIndex=6, endColumnIndex=10, horizontalAlignment='LEFT')
sheets.set_alignment(startRowIndex=3, startColumnIndex=11, endColumnIndex=11, horizontalAlignment='RIGHT')
sheets.set_format(startRowIndex=3, startColumnIndex=11, type='DATE', pattern='dd-mmm-yyyy hh:mm')
# update the cells with the values
sheets.commit()
sheets.update('Characters!A1:K' + str(lastRow), values)

execTime = datetime.utcnow() - now
execTimeStr = str(execTime.seconds) + "." + str(execTime.microseconds)
if LOG_LEVEL_STDOUT > logging.INFO:
    print(f"Done! Required time: {execTimeStr} sec.")
logger.info(f"Done! Required time: {execTimeStr} sec.")
