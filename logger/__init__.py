import os, logging
from logging.handlers import RotatingFileHandler

def get_logger(filename, loglevel=logging.INFO):
    logger = logging.getLogger(__name__)
    if not os.path.exists('logs'):
        os.mkdir('logs')
    file_handler = RotatingFileHandler('logs/' + filename, maxBytes=1048576, backupCount=3)
    file_handler.setFormatter(logging.Formatter('%(asctime)s %(levelname)s: %(message)s'))
    logger.setLevel(logging.INFO)
    logger.addHandler(file_handler)
    return logger
