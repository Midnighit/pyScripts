import re
import sys
import logging
import itertools
from datetime import datetime
from exiles_api import session, Users
from logger import get_logger
from config import WHITELIST_PATH, LOG_LEVEL_STDOUT, LOG_LEVEL_FILE

# catch unhandled exceptions
logger = get_logger('whitelistall.log', log_level_stdout=LOG_LEVEL_STDOUT, log_level_file=LOG_LEVEL_FILE)


def handle_exception(exc_type, exc_value, exc_traceback):
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return

    logger.error("Uncaught exception", exc_info=(exc_type, exc_value, exc_traceback))


sys.excepthook = handle_exception

# save current time
now = datetime.utcnow()
if LOG_LEVEL_STDOUT > logging.INFO:
    print("Whitelisting everyone in the supplemental database...")
logger.info("Whitelisting everyone in the supplemental database...")

try:
    with open(WHITELIST_PATH, 'r') as f:
        logger.debug("Opening existing whitelist.txt.")
        lines = f.readlines()
# if file doesn't exist create an empty list
except Exception:
    with open(WHITELIST_PATH, 'w') as f:
        logger.debug("Create new whitelist.txt.")
        pass
    lines = []

# split lines into id and name. Remove duplicates.
filtered = set()
names = {}
# define regular expression to filter out unprintable characters
logger.debug("Create index from whitelist.txt.")
control_chars = ''.join(map(chr, itertools.chain(range(0x00, 0x20), range(0x7f, 0xa0))))
control_char_re = re.compile('[%s]' % re.escape(control_chars))
for line in lines:
    if line != "\n" and "INVALID" not in line:
        # remove unprintable characters from the line
        res = control_char_re.sub('', line)
        res = res.split(':')
        id = res[0].strip()
        if id == '':
            continue
        if len(res) > 1:
            name = res[1].strip()
        else:
            name = 'Unknown'
        filtered.add(id)
        # if duplicate values exist, prioritize those containing a funcom_name
        if id not in names or names[id] == 'Unknown':
            names[id] = name

# go through the Users table and supplement missing users if any
logger.debug("Supplement whitelist from supplemental.db.")
for user in session.query(Users).all():
    if user.funcom_id not in filtered:
        filtered.add(user.funcom_id)
        names[user.funcom_id] = 'Unknown'

# create lines to write into new whitelist.txt
logger.debug("Create output list and sort it.")
wlist = []
for id in filtered:
    wlist.append(id + ':' + names[id] + '\n')
    wlist.sort()

# overwrite / write the new file with the contenst of wlist
logger.debug("Write list into new whitelist.txt")
with open(WHITELIST_PATH, 'w') as f:
    f.writelines(wlist)

execTime = datetime.utcnow() - now
execTimeStr = str(execTime.seconds) + "." + str(execTime.microseconds)
if LOG_LEVEL_STDOUT > logging.INFO:
    print(f"Done! Required time: {execTimeStr} sec.")
logger.info(f"Done! Required time: {execTimeStr} sec.")
