from google_api import *
from googleapiclient.discovery import build
from datetime import datetime, timedelta

class Spreadsheet:
    def __init__(self, id, activeSheetId=None):
        self.service = build('sheets', 'v4', credentials=credentials)
        self.id = id
        if not activeSheetId is None:
            self._active_sheet_id = activeSheetId
            self._active_sheet_name = self.get_sheet_name(activeSheetId)
        self.requests = []

    """ properties """

    @property
    def active_sheet_id(self):
        return self._active_sheet_name

    @property
    def active_sheet_id(self):
        return self._active_sheet_id

    @active_sheet_id.setter
    def active_sheet_id(self, value):
        if type(value) is str and value.isnumeric():
            name = self.get_sheet_name(int(value))
            if not name:
                raise ValueError(f"sheetId {value} not found in spreadsheet with id {self.id}.")
            self._active_sheet_name = name
            self._active_sheet_id = int(value)
        elif type(value) is int:
            name = self.get_sheet_name(value)
            if not name:
                raise ValueError(f"sheetId {str(value)} not found in spreadsheet with id {self.id}.")
            self._active_sheet_name = name
            self._active_sheet_id = value
        raise ValueError("sheetId must be an integer or numeric string.")

    @property
    def active_sheet_name(self):
        return self._active_sheet_name

    @active_sheet_name.setter
    def active_sheet_name(self, value):
        if type(value) is str:
            id = self.get_sheet_id(value)
            if not id:
                raise ValueError(f"sheetName {value} not found in spreadsheet with id {self.id}.")
            self._active_sheet_id = id
            self._active_sheet_name = value
        raise ValueError("sheetName must be a string.")

    """ internmal helper functions """

    def _get_sheet_id(self, sheetId):
        if type(sheetId) is int:
            return sheetId
        if sheetId is None and not self.active_sheet_id is None:
            return self.active_sheet_id
        if type(sheetId) is str and sheetId.isnumeric():
            return int(sheetId)
        if type(sheetId) is str:
            return self.get_sheet_id(sheetId)
        raise ValueError('No active_sheet_id assigned and no sheet_id given for _get_sheet_id.')

    def _get_color(self, r, g, b, a):
        return {
            "red": r if r <= 1 else r / 255,
            "green": g if g <= 1 else g / 255,
            "blue": b if b <= 1 else b / 255,
            "alpha": a if a <= 1 else a / 255
        }

    def _convert_ordinal_value(self, ordinal):
        if type(ordinal) is str:
            try:
                ordinal = float(ordinal)
            except:
                return ordinal
        if ordinal >= 60:
            ordinal -= 1  # Excel leap year bug, 1900 is not a leap year!
        return (datetime(1899, 12, 31) + timedelta(days=ordinal)).replace(microsecond=0)

    def _convert_ordinal_values(self, values):
        for row_id, row in enumerate(values):
            for cell_id, cell in enumerate(row):
                if type(cell) is float or type(cell) is int or type(cell) is str:
                    values[row_id][cell_id] = self._convert_ordinal_value(cell)

    def _create_batchUpdate_request(self, body):
        return self.service.spreadsheets().batchUpdate(spreadsheetId=self.id, body={'requests': body})

    """ externmal helper functions """

    def convert_R1_A1(self, col):
        R1 = ""
        while col > 0:
            col, remainder = divmod(col - 1, 26)
            R1 = chr(65 + remainder) + R1
        return R1

    def export(self, format, **kwargs):
        ''' Returns a download URL to export a sheet in the given format

        MANDATORY ARGUMENTS:
        spreadsheetId                   # spreadsheet or fileId
        format = 'pdf'                  # pdf/xlsx/ods/csv/tsv/zip(=html)

        DEFAULT ARGUMENTS
        scale = '1',                    # 1 = Normal 100%/2 = Fit to width/3 = Fit to height/4 = Fit to Page
        gridlines = 'false',            # true/false
        printnotes = 'false',           # true/false
        printtitle = 'false',           # true/false
        sheetnames = 'false',           # true/false
        attachment = 'true',            # true/false Opens a download dialog/displays the file in the browser

        OPTIONAL ARGUMENTS:
        sheetId = 123456789             # whole spreadsheet will be printed if not given
        portrait = 'false'              # true = Portrait/false = Landscape
        fitw = 'true'                   # true/false fit to window or actual size
        size = 'a4',                    # A3/A4/A5/B4/B5/letter/tabloid/legal/statement/executive/folio
        top_margin = '1.5'              # All four margins must be set!
        bottom_margin = '1.5'           # All four margins must be set!
        left_margin = '2.0'             # All four margins must be set!
        right_margin = '3.0'            # All four margins must be set!
        pageorder = '1'                 # 1 = Down, then over / 2 = Over, then down
        horizontal_alignment = 'CENTER' # LEFT/CENTER/RIGHT
        vertical_alignment = 'MIDDLE'   # TOP/MIDDLE/BOTTOM
        fzr = 'true'                    # repeat row headers
        fzc = 'true'                    # repeat column headers

        RANGE ARGUMENTS:
        range = 'MyNamedRange'          # Name of the actual range
        ir = 'false'                    # true/false (seems to be always false)
        ic = 'false'                    # same as ir
        r1 = '0'                        # Start Row number-1 (row 1 would be 0)
        c1 = '0'                        # Start Column number-1 (Column 1 would be 0)
        r2 = '15'                       # End Row number
        c2 = '6'                        # End Column number
        '''

        # Set default arguments
        if 'scale' not in kwargs:
            kwargs['scale'] = '1'
        if 'gridlines' not in kwargs:
            kwargs['gridlines'] = 'false'
        if 'printnotes' not in kwargs:
            kwargs['printnotes'] = 'false'
        if 'printtitle' not in kwargs:
            kwargs['printtitle'] = 'false'
        if 'sheetnames' not in kwargs:
            kwargs['sheetnames'] = 'false'
        if 'attachment' not in kwargs:
            kwargs['attachment'] = 'true'

        # make sure the format given is one of those allowed
        possible_formats = ['pdf', 'xlsx', 'ods', 'csv', 'tsv', 'zip']
        if format not in possible_formats:
            return format + ' is not a valid export format!'
        # the google api calls the sheetId 'gid', so we fix this if not passed that way already
        if 'sheetId' in kwargs:
            kwargs['gid'] = kwargs['sheetId']
            del kwargs['sheetId']

        # Create url
        url = 'https://docs.google.com/spreadsheets/d/' + self.id + '/export?format=' + format
        for k, v in kwargs.items():
            url = url + '&' + k + '=' + str(v)
        return url

    def commit(self):
        ''' Execute all the requests collected self.requests '''
        results = []
        for request in self.requests:
            results.append(request.execute())
        self.requests = []
        return results

    """ spreadsheets """

    def get_metadata(self):
        ''' Get sheets metadata '''
        return self.service.spreadsheets().get(spreadsheetId=self.id).execute()['sheets']

    def get_properties(self, sheetId=None, sheetName=None):
        ''' Get sheet properties '''
        if sheetId is None and sheetName is None and not self.active_sheet_id is None:
            sheetId = self.active_sheet_id
        elif type(sheetId) is str and sheetId.isnumeric():
            sheetId = int(sheetId)
        sheets = self.get_metadata()
        if type(sheetId) is int:
            for sheet in sheets:
                if sheet['properties']['sheetId'] == sheetId:
                    return sheet['properties']
        elif type(sheetName) is str:
            for sheet in sheets:
                if sheet['properties']['title'] == sheetName:
                    return sheet['properties']
        return False

    def get_sheet_id(self, sheetName):
        sheets = self.get_metadata()
        for sheet in sheets:
            if sheet['properties']['title'] == sheetName:
                return sheet['properties']['sheetId']
        return False

    def get_sheet_name(self, sheetId=None):
        if sheetId is None and not self.active_sheet_id is None:
            sheetId = self.active_sheet_id
        elif type(sheetId) is str and sheetId.isnumeric():
            sheetId = int(sheetId)
        sheets = self.get_metadata()
        for sheet in sheets:
            if sheet['properties']['sheetId'] == sheetId:
                return sheet['properties']['title']
        return False

    """ spreadsheets.values """

    def read(self, range=None, valueRenderOption='UNFORMATTED_VALUE', is_ordinal=False):
        ''' Read the cells within the given range'''
        if not range:
            range = self._active_sheet_name
        values = self.service.spreadsheets().values().get(
            spreadsheetId = self.id,
            range = range,
            valueRenderOption = valueRenderOption).execute().get('values', [])
        if is_ordinal:
            self._convert_ordinal_values(values)
        return values

    def update(self, range=None, values=[], valueInputOption='USER_ENTERED', majorDimension="ROWS", is_ordinal=False):
        ''' Update the cells within the given range '''
        if is_ordinal:
            self._convert_ordinal_values(values)
        if not range:
            range = self._active_sheet_name
        self.requests.append(self.service.spreadsheets().values().update(
            spreadsheetId=self.id,
            range=range,
            valueInputOption=valueInputOption,
            body={"range": range, "majorDimension": majorDimension, "values": values}
        ))

    def append(self, range=None, values=[], valueInputOption='USER_ENTERED', insertDataOption=None, majorDimension="ROWS", is_ordinal=False):
        ''' Append the cells at the given range '''
        if is_ordinal:
            self._convert_ordinal_values(values)
        if not range:
            range = self._active_sheet_name
        self.requests.append(self.service.spreadsheets().values().append(
            spreadsheetId=self.id,
            range=range,
            valueInputOption=valueInputOption,
            insertDataOption=insertDataOption,
            body={"range": range, "majorDimension": majorDimension, "values": values}
        ))

    """ spreadsheets.batchUpdate """

    def rename(self, sheetId=None, sheetName=None, spreadSheetName=None):
        ''' Renames a sheet or spreadsheet '''
        sheetId = self._get_sheet_id(sheetId)
        if type(spreadSheetName) is str:
            self.requests.append(self._create_batchUpdate_request({
                "updateSpreadsheetProperties": {
                    "fields": "title",
                    "properties": {"title": spreadSheetName}
                }
            }))
        if type(sheetName) is str:
            self.requests.append(self._create_batchUpdate_request({
                "updateSheetProperties": {
                    "fields": "title",
                    "properties": {"sheetId": sheetId, "title": sheetName}
                }
            }))

    def set_grid_size(self, sheetId=None, cols=0, rows=0, frozen=0):
        ''' Sets the gridsize of the sheet '''
        sheetId = self._get_sheet_id(sheetId)
        self.requests.append(self._create_batchUpdate_request({
            "updateSheetProperties": {
                "fields": "gridProperties.columnCount, gridProperties.rowCount, gridProperties.frozenRowCount",
                "properties": {
                    "sheetId": sheetId,
                    "gridProperties": {
                        "columnCount": cols,
                        "rowCount": rows,
                        "frozenRowCount": frozen
                    }
                }
            }
        }))

    def set_dimension_group(self, sheetId=None, startIndex=1, endIndex=1, visibility=None, dimension='ROWS'):
        sheetId = self._get_sheet_id(sheetId)
        self.requests.append(self._create_batchUpdate_request({
            "addDimensionGroup": {
                "range": {
                    "sheetId": sheetId,
                    "startIndex": startIndex,
                    "endIndex": endIndex,
                    "dimension": dimension
                }
            }
        }))
        if not visibility is None:
            self.set_visibility(sheetId, startIndex, endIndex, visibility, dimension)

    def delete_dimension_group(self, sheetId=None, startIndex=1, endIndex=1, dimension = 'ROWS'):
        sheetId = self._get_sheet_id(sheetId)
        self.requests.append(self._create_batchUpdate_request({
            "deleteDimensionGroup": {
                "range": {
                    "sheetId": sheetId,
                    "startIndex": startIndex,
                    "endIndex": endIndex,
                    "dimension": dimension
                }
            }
        }))

    def set_visibility(self, sheetId=None, startIndex=1, endIndex=1, visibility=False, dimension='ROWS'):
        sheetId = self._get_sheet_id(sheetId)
        self.requests.append(self._create_batchUpdate_request({
            "updateDimensionProperties": {
                "fields": 'hiddenByUser',
                "properties": {
                    "hiddenByUser": visibility
                },
                "range": {
                    "sheetId": sheetId,
                    "startIndex": startIndex,
                    "endIndex": endIndex,
                    "dimension": dimension
                }
            }
        }))

    def merge_cells(self, sheetId=None, startColumnIndex=1, startRowIndex=1, endColumnIndex=None, endRowIndex=None, mergeType = "MERGE_ROWS"):
        sheetId = self._get_sheet_id(sheetId)
        if not mergeType in ('MERGE_ROWS', 'MERGE_COLUMNS'):
            mergeType = 'MERGE_ROWS'
        self.requests.append(self._create_batchUpdate_request({
            "mergeCells": {
                "mergeType": mergeType,
                "range": {
                    "sheetId": sheetId,
                    "startRowIndex": startRowIndex - 1,
                    "startColumnIndex": startColumnIndex - 1,
                    "endRowIndex": endRowIndex,
                    "endColumnIndex": endColumnIndex,
                }
            }
        }))

    def unmerge_cells(self, sheetId=None, startColumnIndex=1, startRowIndex=1, endColumnIndex=None, endRowIndex=None):
        sheetId = self._get_sheet_id(sheetId)
        self.requests.append(self._create_batchUpdate_request({
            "unmergeCells": {
                "range": {
                    "sheetId": sheetId,
                    "startRowIndex": startRowIndex - 1,
                    "startColumnIndex": startColumnIndex - 1,
                    "endRowIndex": endRowIndex,
                    "endColumnIndex": endColumnIndex,
                }
            }
        }))

    def add_named_range(self, sheetId=None, name=None, namedRangeId=None, startColumnIndex=1, startRowIndex=1, endColumnIndex=None, endRowIndex=None):
        sheetId = self._get_sheet_id(sheetId)
        if name is None:
            return
        request = self._create_batchUpdate_request({
            "addNamedRange": {
                "namedRange": {
                    "name": name,
                    "range": {
                        "sheetId": sheetId
                    }
                }
            }
        })
        if namedRangeId:
            request["addNamedRange"]["namedRange"]["namedRangeId"] = str(namedRangeId)
        if startCol:
            request["addNamedRange"]["namedRange"]["range"]['startColumnIndex'] = startCol - 1
        if endCol:
            request["addNamedRange"]["namedRange"]["range"]['endColumnIndex'] = endCol
        if startRow:
            request["addNamedRange"]["namedRange"]["range"]['startRowIndex'] = startRow - 1
        if endRow:
            request["addNamedRange"]["namedRange"]["range"]['endRowIndex'] = endRow
        self.requests.append(request)

    def set_filter(self, sheetId=None, startColumnIndex=1, startRowIndex=1, endColumnIndex=None, endRowIndex=None):
        sheetId = self._get_sheet_id(sheetId)
        self.requests.append(self._create_batchUpdate_request({
            "setBasicFilter": {
                "filter": {
                    "range": {
                        "sheetId": sheetId,
                        "startRowIndex": startRowIndex - 1,
                        "startColumnIndex": startColumnIndex - 1,
                        "endRowIndex": endRowIndex,
                        "endColumnIndex": endColumnIndex,
                    }
                }
            }
        }))

    def set_bg_color(self, sheetId=None, startColumnIndex=1, startRowIndex=1, endColumnIndex=None, endRowIndex=None, color=None, red=0, green=0, blue=0, alpha=0):
        sheetId = self._get_sheet_id(sheetId)
        if color and color in COLORS:
            backgroundColor = self._get_color(COLORS[color][0], COLORS[color][1], COLORS[color][2], alpha)
        else:
            backgroundColor = self._get_color(red, green, blue, alpha)
        self.requests.append(self._create_batchUpdate_request({
            "repeatCell": {
                "fields": "userEnteredFormat.backgroundColor",
                "range": {
                    "sheetId": sheetId,
                    "startRowIndex": startRowIndex - 1,
                    "startColumnIndex": startColumnIndex - 1,
                    "endRowIndex": endRowIndex,
                    "endColumnIndex": endColumnIndex,
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": backgroundColor
                    }
                }
            }
        }))

    def set_color(self, sheetId=None, startColumnIndex=1, startRowIndex=1, endColumnIndex=None, endRowIndex=None, color=None, red=0, green=0, blue=0, alpha=0):
        sheetId = self._get_sheet_id(sheetId)
        if color and color in COLORS:
            color = self._get_color(COLORS[color][0], COLORS[color][1], COLORS[color][2], alpha)
        else:
            color = self._get_color(red, green, blue, alpha)
        self.requests.append(self._create_batchUpdate_request({
            "repeatCell": {
                "fields": "userEnteredFormat.textFormat.foregroundColor",
                "range": {
                    "sheetId": sheetId,
                    "startRowIndex": startRowIndex - 1,
                    "startColumnIndex": startColumnIndex - 1,
                    "endRowIndex": endRowIndex,
                    "endColumnIndex": endColumnIndex,
                },
                "cell": {
                    "userEnteredFormat": {
                        "textFormat": {
                            "foregroundColor": color
                        }
                    }
                }
            }
        }))

    def set_format(self, sheetId=None, startColumnIndex=1, startRowIndex=1, endColumnIndex=None, endRowIndex=None, type=None, pattern=None):
        sheetId = self._get_sheet_id(sheetId)
        numberFormat = {}
        if type in NUMBER_FORMATS:
            numberFormat['type'] = type
        if pattern:
            numberFormat['pattern'] = pattern
        self.requests.append(self._create_batchUpdate_request({
            "repeatCell": {
                "fields": "userEnteredFormat.numberFormat",
                "range": {
                    "sheetId": sheetId,
                    "startRowIndex": startRowIndex - 1,
                    "startColumnIndex": startColumnIndex - 1,
                    "endRowIndex": endRowIndex,
                    "endColumnIndex": endColumnIndex,
                },
                "cell": {
                    "userEnteredFormat": {
                        "numberFormat": numberFormat
                    }
                }
            }
        }))

    def set_alignment(self, sheetId=None, startColumnIndex=1, startRowIndex=1, endColumnIndex=None, endRowIndex=None, horizontalAlignment=None, verticalAlignment=None):
        sheetId = self._get_sheet_id(sheetId)
        userEnteredFormat = {}
        if horizontalAlignment in HORIZONTAL_ALIGNMENT:
            userEnteredFormat['horizontalAlignment'] = horizontalAlignment
        if verticalAlignment in VERTICAL_ALIGNMENT:
            userEnteredFormat['verticalAlignment'] = verticalAlignment
        self.requests.append(self._create_batchUpdate_request({
            "repeatCell": {
                "fields": "userEnteredFormat.horizontalAlignment, userEnteredFormat.verticalAlignment",
                "range": {
                    "sheetId": sheetId,
                    "startRowIndex": startRowIndex - 1,
                    "startColumnIndex": startColumnIndex - 1,
                    "endRowIndex": endRowIndex,
                    "endColumnIndex": endColumnIndex,
                },
                "cell": {
                    "userEnteredFormat": userEnteredFormat
                }
            }
        }))

    def set_wrap(self, sheetId=None, startColumnIndex=1, startRowIndex=1, endColumnIndex=None, endRowIndex=None, wrapStrategy='OVERFLOW_CELL'):
        sheetId = self._get_sheet_id(sheetId)
        userEnteredFormat = {}
        if wrapStrategy in WRAP_STRATEGY:
            userEnteredFormat['wrapStrategy'] = wrapStrategy
        self.requests.append(self._create_batchUpdate_request({
            "repeatCell": {
                "fields": "userEnteredFormat.wrapStrategy",
                "range": {
                    "sheetId": sheetId,
                    "startRowIndex": startRowIndex - 1,
                    "startColumnIndex": startColumnIndex - 1,
                    "endRowIndex": endRowIndex,
                    "endColumnIndex": endColumnIndex,
                },
                "cell": {
                    "userEnteredFormat": userEnteredFormat
                }
            }
        }))

    def set_borders(self):
        pass #ToDo

    def insert_rows(self, sheetId=None, startIndex=1, numRows=1, inheritFromBefore=True):
        sheetId = self._get_sheet_id(sheetId)
        self.requests.append(self._create_batchUpdate_request({
            "insertDimension": {
                "inheritFromBefore": inheritFromBefore,
                "range": {
                    "sheetId": sheetId,
                    "dimension": "ROWS",
                    "startIndex": startIndex - 1,
                    "endIndex": startIndex - 1 + numRows,
                }
            }
        }))

    def insert_columns(self, sheetId=None, startIndex=1, numColumns=1, inheritFromBefore=True):
        sheetId = self._get_sheet_id(sheetId)
        self.requests.append(self._create_batchUpdate_request({
            "insertDimension": {
                "inheritFromBefore": inheritFromBefore,
                "range": {
                    "sheetId": sheetId,
                    "dimension": "COLUMNS",
                    "startIndex": startIndex - 1,
                    "endIndex": startIndex - 1 + numColumns,
                }
            }
        }))

    def delete_rows(self, sheetId=None, startIndex=1, numRows=1):
        sheetId = self._get_sheet_id(sheetId)
        self.requests.append(self._create_batchUpdate_request({
            "deleteDimension": {
                "range": {
                    "sheetId": sheetId,
                    "dimension": "ROWS",
                    "startIndex": startIndex - 1,
                    "endIndex": startIndex - 1 + numRows,
                }
            }
        }))

    def delete_columns(self, sheetId=None, startIndex=1, numColumns=1):
        sheetId = self._get_sheet_id(sheetId)
        self.requests.append(self._create_batchUpdate_request({
            "deleteDimension": {
                "range": {
                    "sheetId": sheetId,
                    "dimension": "COLUMNS",
                    "startIndex": startIndex - 1,
                    "endIndex": startIndex - 1 + numColumns,
                }
            }
        }))
