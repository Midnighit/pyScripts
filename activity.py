from config import *
from datetime import datetime
from exiles_api.model import session, ServerPopulationRecordings
from google_api.sheets import Spreadsheet

# instanciate the Spreadsheet object
sheets = Spreadsheet(ACTIVITY_SPREADSHEET_ID, active_sheet_id=ACTIVITY_SHEET_ID)

# read all the timestamps and population values from ServerPopulationRecordings
values = []
for record in session.query(ServerPopulationRecordings).all():
    values.append([datetime.utcfromtimestamp(record.time_of_recording).strftime("%Y-%m-%d %H:%M"), int(record.population * MAX_POP)])
lastRow = len(values) + 1
sheets.update('Activity Statistics!A2:B' + str(lastRow), values)
sheets.set_grid_size(cols=2, rows=lastRow, frozen=1)
sheets.set_format(startRowIndex=2, endRowIndex=lastRow, type='DATE_TIME', pattern='ddd dd-mmm-yyy')
sheets.set_format(startColumnIndex=2, endColumnIndex=2, startRowIndex=2, endRowIndex=lastRow, type='NUMBER', pattern='0')
sheets.commit()
