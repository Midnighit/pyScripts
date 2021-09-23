import os, logging, sys
from logging.handlers import RotatingFileHandler
from config import *

def get_logger(filename, log_level_stdout=logging.WARNING, log_level_file=logging.INFO, echo=True):
    logger = logging.getLogger(__name__)
    if not os.path.exists(SAVED_DIR_PATH + "/Logs"):
        os.mkdir(SAVED_DIR_PATH + "/Logs")

    logger.setLevel(logging.DEBUG)

    if echo:
        prn_handler = logging.StreamHandler(sys.stdout)
        prn_handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s: %(message)s"))
        prn_handler.setLevel(log_level_stdout)
        logger.addHandler(prn_handler)

    file_handler = RotatingFileHandler(
        SAVED_DIR_PATH + "/Logs/" + filename,
        maxBytes=1048576,
        backupCount=3,
        encoding="utf8"
    )
    file_handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s: %(message)s"))
    file_handler.setLevel(log_level_file)
    logger.addHandler(file_handler)
    return logger
