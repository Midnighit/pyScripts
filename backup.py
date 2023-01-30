import os, sys, logging, shutil
from datetime import datetime, timedelta
from logger import get_logger
from config import LOG_LEVEL_STDOUT, LOG_LEVEL_FILE, SAVED_DIR_PATH, BACKUP_DIR_PATH, FILES_TO_BACKUP

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
    print("Checking if backup is required...")
logger.info("Checking if backup is required...")

# try getting the date of the last entry of the logfile
logfile = os.path.join(SAVED_DIR_PATH, 'Logs', 'ConanSandbox.log')
try:
    with open(logfile, 'r') as f:
        lines = f.readlines()
        log_dt = datetime.strptime(lines[-1][1:20], '%Y.%m.%d-%H.%M.%S')

# if file doesn't, exist exit with error message
except Exception as e:
    logger.error("Error:", e)
    sys.exit()

# get the list of files and directories at the configured backup path
try:
    directories = os.listdir(BACKUP_DIR_PATH)

# if path doesn't, exist exit with error message
except Exception as e:
    logger.error("Error:", e)
    sys.exit()

# check if an entry for the date of the crash exists starting from the last directory
backupdir = None
directories.sort(reverse=True)
for directory in directories:
    try:
        dir_d = datetime.strptime(directory, '%Y.%m.%d')
    except Exception:
        # not a date, skip over
        continue
    
    # directory for the date in the logfile already exists
    if dir_d.date() == log_dt.date():
        backupdir = os.path.join(BACKUP_DIR_PATH, directory)
        break
    
    # if the last directory is older than the logfile, a new directory has to be created
    elif dir_d < log_dt:
        backupdir = os.path.join(BACKUP_DIR_PATH, log_dt.strftime('%Y.%m.%d'))
        os.mkdir(backupdir)
        break

# if no backup directory was either found or created, exit script
if not backupdir:
    logger.error(f"Couldn't find or create backup folder {log_dt.strftime('%Y.%m.%d')} in {BACKUP_DIR_PATH}.")
    sys.exit()

# get the list of files and directories at the path matching the date
directories = os.listdir(backupdir)
directories.sort(reverse=True)
dir_t = None
for directory in directories:
    try:
        dir_t = datetime.strptime(directory, '%H.%M.%S').time()
        break
    except Exception:
        # not a date, skip over
        continue

# if no directory with a valid time is in backupdir, one has to be created
if not dir_t:
    dir_t = log_dt.strftime('%H.%M.%S')
    backupdir = os.path.join(backupdir, dir_t)
    os.mkdir(backupdir)
# if the latest valid timestamp is not within 5 minutes of the time of the log, create a new directory as well
elif (log_dt - timedelta(minutes=5)) > datetime.combine(dir_d, dir_t):
    dir_t = log_dt.strftime('%H.%M.%S')
    backupdir = os.path.join(backupdir, dir_t)
    os.mkdir(backupdir)
# a valid timestamp was found an it's within 5 minutes of the time of the log, DSL backup worked correctly
else:
    backupdir = None

# if a new folder was created, copy all the relevant files there
if backupdir:
    for file in FILES_TO_BACKUP:
        try:
            shutil.copy2(file, backupdir)
            logger.info(f"Copying {os.path.basename(file)} to backup folder.")
        except Exception as error:
            # Errno 2 is 'No such file or directory'
            if error.args[0] == 2:
                pass
            else:
                logger.error(str(error))

execTime = datetime.utcnow() - now
execTimeStr = str(execTime.seconds) + "." + str(execTime.microseconds)
if LOG_LEVEL_STDOUT > logging.INFO:
    print(f"Done! Required time: {execTimeStr} sec.")
logger.info(f"Done! Required time: {execTimeStr} sec.")

