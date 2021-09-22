import os, logging, sys
from logging.handlers import RotatingFileHandler
from config import *

def get_logger(filename, loglevel=logging.INFO, echo=True):
    logger = logging.getLogger(__name__)
    if not os.path.exists(SAVED_DIR_PATH + '/Logs'):
        os.mkdir(SAVED_DIR_PATH + '/Logs')

    logger.setLevel(loglevel)

    if echo:
        prn_handler = logging.StreamHandler(sys.stdout)
        prn_handler.setFormatter(logging.Formatter('%(asctime)s %(levelname)s: %(message)s'))
        prn_handler.setLevel(logging.DEBUG)
        logger.addHandler(prn_handler)

    file_handler = RotatingFileHandler(SAVED_DIR_PATH + '/Logs/' + filename, maxBytes=1048576, backupCount=3)
    file_handler.setFormatter(logging.Formatter('%(asctime)s %(levelname)s: %(message)s'))
    file_handler.setLevel(logging.INFO)
    logger.addHandler(file_handler)
    return logger
