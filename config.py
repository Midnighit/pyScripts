# pyScripts Config live

import os
import logging
from datetime import timedelta
from dotenv import load_dotenv

load_dotenv()

LOG_LEVEL_STDOUT = logging.INFO
LOG_LEVEL_FILE = logging.INFO

# Generic constants used by multiple sheets
INACTIVITY = timedelta(weeks=2)          # time before a character or clan is considered to be inactive.
LONG_INACTIVE = timedelta(days=30)       # Number of days before a character is deleted from the db.
PURGE = timedelta(days=5)                # Number of days before a base is being purged from the db.
EVENT_LOG_HOLD_BACK = timedelta(days=7)  # number of days to keep in the event log of the game.db
CHAT_LOG_HOLD_BACK = timedelta(days=14)  # number of days to keep in the chat log of the google sheet
RUINS_CLAN_ID = -20                      # id of the clan that all ownerless objects are moved into
MIN_DIST = 50000                         # Min. distance that for a row to be listed on the google sheet.
OWNER_WHITELIST = list(range(-19, 0)) + [154, 161, 163, 167, 178, 185, 247, 248, 554, 666, 726, 1293, 1368]
PREFABS = ['Prefab_Dwelling', 'Prefab_NordheimerBuilding', 'Prefab_Mound']
OBJECT_LIMITS = [
    {'max':  5, 'obj': ['Pippi_Glorb']},
    {'max':  6, 'obj': ['Pippi_Flaggi']},
    {'max': 10, 'obj': ['Tot_A_BasicNPC'], 'pm': 1},
    {'max':  3, 'obj': ['Tot_A_TraderNPC']},
    {'max':  1, 'obj': ['Crafting_Beehive']},
    {'max':  2, 'obj': ['Crafting_FishNet', 'Crafting_CrabPot']},
    {'max':  2, 'obj': ['Compost']},
    {'max':  1, 'obj': ['SvS_BP_Dust']},
    {'max':  2, 'obj': ['Crafting_Planter']},
    {'max':  0, 'obj': ['FeedingContainer']},
    {'max':  0, 'obj': ['DwellingNewWall_.']},
    {'max':  0, 'obj': ['Bedroll_Clean', 'Bedroll_Fiber', 'Bedroll_Turanian']},
    {'max':  2, 'obj': ['AnimalPen_Tier', 'AnimalPens', 'AnimalPen_Onestall']},
    {'max':  2, 'obj': PREFABS, 'pm':  1}
]

""" Player sheets """
# Tiles per member constants
PLAYER_SPREADSHEET_ID = os.getenv('PLAYER_SPREADSHEET_ID')
PLAYER_TPM_SHEET_ID = os.getenv('PLAYER_TPM_SHEET_ID')
PLAYER_STATISTICS_SHEET_ID = os.getenv('PLAYER_STATISTICS_SHEET_ID')
HIDE_WHITELISTED_OWNERS = True          # whether to list tiles for whitelisted owners or not

""" Need to be set in TERPBot as well """
ALLOWANCE_BASE = 500                    # Number of tiles a single player may have
ALLOWANCE_CLAN = 250                    # Number of tiles every active clan member over the first adds to the base tiles
ALLOWANCE_INCLUDES_INACTIVES = True     # True/False => all/only active members count towards allowance
BUILDING_TILE_MULT = 1                  # Multiplier for the building tiles
PLACEBALE_TILE_MULT = 3/5               # Multiplier for the placeable tiles
PLACEBALE_TILE_RATIO = 1/3              # Ratio of total tiles that may be placeables

# Activity constants
PLAYER_ACTIVITY_SPREADSHEET_ID = PLAYER_SPREADSHEET_ID
PLAYER_ACTIVITY_SHEET_ID = os.getenv('PLAYER_ACTIVITY_SHEET_ID')
ACTIVITY_HOLD_BACK = timedelta(weeks=1)  # duration that the activity chart should keep
MAX_POP = 70                             # the maximum number of players allowed on the server

""" Admin sheets """
ADMIN_SPREADSHEET_ID = os.getenv('ADMIN_SPREADSHEET_ID')
ADMIN_CHARACTERS_SHEET_ID = os.getenv('ADMIN_CHARACTERS_SHEET_ID')
ADMIN_CLANS_SHEET_ID = os.getenv('ADMIN_CLANS_SHEET_ID')
ADMIN_TPM_SHEET_ID = os.getenv('ADMIN_TPM_SHEET_ID')
ADMIN_LINT_SHEET_ID = os.getenv('ADMIN_LINT_SHEET_ID')

""" Log sheets """
LOGS_SPREADSHEET_ID = os.getenv('LOGS_SPREADSHEET_ID')
LOGS_CHAT_SHEET_ID = os.getenv('LOGS_CHAT_SHEET_ID')
LOGS_COMMANDS_SHEET_ID = os.getenv('LOGS_COMMANDS_SHEET_ID')

""" Misc sheets """
TILES_MGMT_SPREADSHEET_ID = os.getenv('TILES_MGMT_SPREADSHEET_ID')
TILES_MGMT_SHEET_ID = os.getenv('TILES_MGMT_SHEET_ID')

""" Exiles API constants """
SAVED_DIR_PATH = os.getenv('SAVED_DIR_PATH')
EXE_DIR_PATH = os.getenv("EXE_DIR_PATH")
BACKUP_DIR_PATH = os.getenv('BACKUP_DIR_PATH')
LOGS_DIR_PATH = os.getenv('LOGS_DIR_PATH')
CONFIG_DIR_PATH = os.getenv('CONFIG_DIR_PATH')
GAME_DB = "game.db"
BACKUP_DB = "backup.db"
USERS_DB = "supplemental.db"
GAME_DB_URI = "sqlite:///" + SAVED_DIR_PATH + "/" + GAME_DB
BACKUP_DB_URI = "sqlite:///" + SAVED_DIR_PATH + "/" + BACKUP_DB
USERS_DB_URI = "sqlite:///" + SAVED_DIR_PATH + "/" + USERS_DB
WHITELIST_PATH = SAVED_DIR_PATH + '/whitelist.txt'
ECHO = False
FILES_TO_BACKUP = [
    os.path.join(SAVED_DIR_PATH, GAME_DB),
    os.path.join(SAVED_DIR_PATH, BACKUP_DB),
    os.path.join(SAVED_DIR_PATH, USERS_DB),
    os.path.join(SAVED_DIR_PATH, 'blacklist.txt'),
    os.path.join(SAVED_DIR_PATH, 'whitelist.txt'),
    os.path.join(LOGS_DIR_PATH, 'ConanSandbox.log'),
    os.path.join(CONFIG_DIR_PATH, 'DedicatedServerLauncher.ini'),
    os.path.join(CONFIG_DIR_PATH, 'Engine.ini'),
    os.path.join(CONFIG_DIR_PATH, 'Game.ini'),
    os.path.join(CONFIG_DIR_PATH, 'ServerSettings.ini')
]

""" Reindex script parameters """
RI_DEST = [idx for idx in range(-1, -31, -1)]
RI_SOURCE = [103, 104, 106, 107, 108, 109, 112, 244, 8982]
