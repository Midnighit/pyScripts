from google_api import *
from googleapiclient.discovery import build

class Spreadsheet:
    def __init__(self, spreadsheetId, active_sheet_id=None):
        self.service = build('sheets', 'v4', credentials = credentials)
        self.spreadsheetId = spreadsheetId
        if not active_sheet_id is None:
            self._active_sheet_id = active_sheet_id
        self.updates = {}

    """ Helper """
    @property
    def active_sheet_id(self):
        return self._active_sheet_id

    @active_sheet_id.setter
    def active_sheet_id(self, value):
        if type(value) is str and value.isnumeric():
            self._active_sheet_id = int(value)
        elif type(value) is str:
            self._active_sheet_id = self.get_sheet_id(value)

    def _get_sheet_id(self, sheetId):
        if sheetId is None and not self.active_sheet_id is None:
            return self.active_sheet_id
        if type(sheetId) is str and sheetId.isnumeric():
            return int(sheetId)
        if type(sheetId) is str:
            return self.get_sheet_id(sheetId)

    def _get_color(self, r, g, b, a):
        return {
            "red": r if r <= 1 else r / 255,
            "green": g if g <= 1 else g / 255,
            "blue": b if b <= 1 else b / 255,
            "alpha": a if a <= 1 else a / 255
        }

    def convertR1toA1(self, col):
        R1 = ""
        while col > 0:
            col, remainder = divmod(col - 1, 26)
            R1 = chr(65 + remainder) + R1
        return R1

    """ Non-Requests """

    def get_metadata(self):
        ''' Get sheets metadata '''
        return self.service.spreadsheets() \
            .get(spreadsheetId = self.spreadsheetId) \
            .execute()['sheets']

    def get_properties(self, sheetId = None):
        ''' Get sheet properties '''
        if sheetId is None and not self.active_sheet_id is None:
            sheetId = self.active_sheet_id
        elif type(sheetId) is str and sheetId.isnumeric():
            sheetId = int(sheetId)
        sheets = self.get_metadata()
        if type(sheetId) is int:
            for sheet in sheets:
                if sheet['properties']['sheetId'] == sheetId:
                    return sheet['properties']
        elif type(sheetId) is str:
            for sheet in sheets:
                if sheet['properties']['title'] == sheetId:
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

    def read(self, range, value_render_option = 'UNFORMATTED_VALUE'):
        ''' Read the cells within the given range'''
        return self.service.spreadsheets().values().get(
            spreadsheetId = self.spreadsheetId,
            range = range,
            valueRenderOption = value_render_option) \
            .execute() \
            .get('values', [])

    def update(self, range, values, valueInputOption = 'USER_ENTERED'):
        ''' Update the cells within the given range '''
        valueRangeBody = {"values": values}
        return self.service.spreadsheets().values().update(
            spreadsheetId = self.spreadsheetId,
            range = range,
            valueInputOption = valueInputOption,
            body = valueRangeBody) \
            .execute()

    def duplicate_sheet(sourceSheetId, destinationSpreadsheetId):
        ''' Copy a sheet into an existing Spreadsheet '''
        if type(sourceSheetId) is str and sourceSheetId.isnumeric():
            sourceSheetId = int(sourceSheetId)
        elif type(sourceSheetId) is str:
            sourceSheetId = self.get_sheet_id(sourceSheetId)
        body = {"destinationSpreadsheetId": destinationSpreadsheetId}
        response = self.service.spreadsheets().sheets().copyTo(
                spreadsheetId = self.spreadsheetId,
                sheetId = sourceSheetId,
                body = body) \
                .execute()
        return response

    # may need fixing with new object style!
    def create_spreadsheet(spreadsheetTitle, sheetTitles):
        ''' Creates an empty spreadsheet with some sheets '''
        spreadsheetBody = {
            "properties": {"title": spreadsheetTitle},
            "sheets": []
        }
        if type(sheetTitles) is str:
            sheetTitles = [sheetTitles]
        for title in sheetTitles:
            spreadsheetBody["sheets"].append({"properties": {"title": title} })
        return self.service.spreadsheets().create(body = spreadsheetBody).execute()

    def crop(self, sheetId = None, frozen = 0):
        sheetName = ''
        if sheetId is None and not self.active_sheet_id is None:
            sheetId = self.active_sheet_id
        elif type(sheetId) is str and sheetId.isnumeric():
            sheetId = int(sheetId)
            sheetName = self.get_sheet_name(sheetId)
        elif type(sheetId) is str:
            sheetName = sheetId
            sheetId = self.get_sheet_id(sheetName)
        if sheetName == '':
            sheetName = self.get_sheet_name(sheetId)
        results = self.get_properties(sheetId)
        rows = results['gridProperties']['rowCount']
        cols = results['gridProperties']['columnCount']
        range = sheetName + '!A1:' + self.convertR1toA1(cols) + str(rows)
        results = self.read(range)
        maxRows = 1
        maxColumns = 1
        idx = 0
        for currRow in results:
            columnsInCurrRow = len(currRow)
            if columnsInCurrRow > 0:
                maxRows = idx + 1
                maxColumns = max(maxColumns, columnsInCurrRow)
            idx = idx + 1
        self.set_grid_size(sheetId, maxColumns, maxRows, frozen)
        return [maxColumns, maxRows]

    def export(self, format, **kwargs):
        ''' Exports the Spreadsheet using a variable number of arguments

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
        url = 'https://docs.google.com/spreadsheets/d/' + self.spreadsheetId + '/export?format=' + format
        for k, v in kwargs.items():
            url = url + '&' + k + '=' + str(v)
        return url

    def check(self, sheetId = None):
        values = []
        if sheetId == None:
            sheets = self.get_metadata()
            for sheet in sheets:
                chk = self.check(sheet['properties']['sheetId'])
                if chk:
                    values = values + chk
        else:
            sheetName = ''
            if type(sheetId) is str and sheetId.isnumeric():
                sheetId = int(sheetId)
                sheetName = self.get_sheet_name(sheetId)
            elif type(sheetId) is str:
                sheetName = sheetId
                sheetId = self.get_sheet_id(sheetName)
            if sheetName == '':
                sheetName = self.get_sheet_name(sheetId)
            results = self.get_properties(sheetId)
            rows = results['gridProperties']['rowCount']
            cols = results['gridProperties']['columnCount']
            range = sheetName + '!A1:' + self.convertR1toA1(cols) + str(rows)
            results = self.read(range)
            rowIdx = 0
            for currRow in results:
                rowIdx = rowIdx + 1
                colIdx = 0
                for currCell in currRow:
                    colIdx = colIdx + 1
                    if currCell == '#ERROR!' or \
                        currCell == '#NAME?' or \
                        currCell == '#REF!':
                        values.append(sheetName + '!' + self.convertR1toA1(colIdx) + str(rowIdx))
        if values == []:
            return None
        return values

    """ Requests """

    def commit(self):
        ''' Execute all the requests collected in updates '''
        results = []
        for id, requests in self.updates.items():
            results.append(self.service.spreadsheets().batchUpdate(spreadsheetId = id, body = requests).execute())
        self.updates = {}

    def rename(self, arg1, arg2 = None):
        ''' Renames a sheet or spreadsheet '''
        if not self.spreadsheetId in self.updates:
            self.updates[self.spreadsheetId] = {'requests': []}
        # if only one argument is passed => rename whole spreadsheet
        if arg2 == None:
            self.updates[self.spreadsheetId]['requests'].append({
                "updateSpreadsheetProperties": {
                    "fields": "title",
                    "properties": {"title": arg1}
                }
            })
        # if two arguments are passed => rename just the sheet
        else:
            self.updates[self.spreadsheetId]['requests'].append({
                "updateSheetProperties": {
                    "fields": "title",
                    "properties": {"sheetId": arg1, "title": arg2}
                }
            })

    def set_grid_size(self, sheetId = None, cols = 0, rows = 0, frozen = 0):
        ''' Sets the gridsize of the sheet '''
        sheetId = self._get_sheet_id(sheetId)
        if not self.spreadsheetId in self.updates:
            self.updates[self.spreadsheetId] = {'requests': []}
        self.updates[self.spreadsheetId]['requests'].append({
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
        })

    def set_dimension_group(self, sheetId = None, startIndex = 1, endIndex = 1, visibility = None, dimension = 'ROWS'):
        sheetId = self._get_sheet_id(sheetId)
        if not self.spreadsheetId in self.updates:
            self.updates[self.spreadsheetId] = {'requests': []}
        self.updates[self.spreadsheetId]['requests'].append({
            "addDimensionGroup": {
                "range": {
                    "sheetId": sheetId,
                    "startIndex": startIndex,
                    "endIndex": endIndex,
                    "dimension": dimension
                }
            }
        })
        if not visibility is None:
            self.set_visibility(sheetId, startIndex, endIndex, visibility, dimension)

    def delete_dimension_group(self, sheetId = None, startIndex = 1, endIndex = 1, dimension = 'ROWS'):
        sheetId = self._get_sheet_id(sheetId)
        if not self.spreadsheetId in self.updates:
            self.updates[self.spreadsheetId] = {'requests': []}
        self.updates[self.spreadsheetId]['requests'].append({
            "deleteDimensionGroup": {
                "range": {
                    "sheetId": sheetId,
                    "startIndex": startIndex,
                    "endIndex": endIndex,
                    "dimension": dimension
                }
            }
        })

    def set_visibility(self, sheetId = None, startIndex = 1, endIndex = 1, visibility = False, dimension = 'ROWS'):
        sheetId = self._get_sheet_id(sheetId)
        if not self.spreadsheetId in self.updates:
            self.updates[self.spreadsheetId] = {'requests': []}
        self.updates[self.spreadsheetId]['requests'].append({
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
        })

    def merge_cells(self, sheetId = None, startColumnIndex = 1, startRowIndex = 1, endColumnIndex = None, endRowIndex = None, mergeType = "MERGE_ROWS"):
        sheetId = self._get_sheet_id(sheetId)
        if not self.spreadsheetId in self.updates:
            self.updates[self.spreadsheetId] = {'requests': []}
        if not mergeType in ('MERGE_ROWS', 'MERGE_COLUMNS'):
            mergeType = 'MERGE_ROWS'
        self.updates[self.spreadsheetId]['requests'].append({
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
        })

    def unmerge_cells(self, sheetId = None, startColumnIndex = 1, startRowIndex = 1, endColumnIndex = None, endRowIndex = None):
        sheetId = self._get_sheet_id(sheetId)
        if not self.spreadsheetId in self.updates:
            self.updates[self.spreadsheetId] = {'requests': []}
        self.updates[self.spreadsheetId]['requests'].append({
            "unmergeCells": {
                "range": {
                    "sheetId": sheetId,
                    "startRowIndex": startRowIndex - 1,
                    "startColumnIndex": startColumnIndex - 1,
                    "endRowIndex": endRowIndex,
                    "endColumnIndex": endColumnIndex,
                }
            }
        })

    def add_named_range(self, sheetId = None, name = None, namedRangeId = None, startColumnIndex = None, startRowIndex = None, endColumnIndex = None, endRowIndex = None):
        sheetId = self._get_sheet_id(sheetId)
        if name is None:
            return
        if not self.spreadsheetId in self.updates:
            self.updates[self.spreadsheetId] = {'requests': []}
        self.updates[self.spreadsheetId]['requests'].append({
            "addNamedRange": {
                "namedRange": {
                    "name": name,
                    "range": {
                        "sheetId": sheetId
                    }
                }
            }
        })
        idx = len(self.updates[self.spreadsheetId]['requests']) - 1
        if namedRangeId:
            self.updates[self.spreadsheetId]['requests'][idx]["addNamedRange"]["namedRange"]["namedRangeId"] = str(namedRangeId)
        if startCol:
            self.updates[self.spreadsheetId]['requests'][idx]["addNamedRange"]["namedRange"]["range"]['startColumnIndex'] = startCol - 1
        if endCol:
            self.updates[self.spreadsheetId]['requests'][idx]["addNamedRange"]["namedRange"]["range"]['endColumnIndex'] = endCol
        if startRow:
            self.updates[self.spreadsheetId]['requests'][idx]["addNamedRange"]["namedRange"]["range"]['startRowIndex'] = startRow - 1
        if endRow:
            self.updates[self.spreadsheetId]['requests'][idx]["addNamedRange"]["namedRange"]["range"]['endRowIndex'] = endRow

    def set_bg_color(self, sheetId = None, startColumnIndex = 1, startRowIndex = 1, endColumnIndex = None, endRowIndex = None, color = None, red=0, green=0, blue=0, alpha=0):
        sheetId = self._get_sheet_id(sheetId)
        if color and color in COLORS:
            backgroundColor = self._get_color(COLORS[color][0], COLORS[color][1], COLORS[color][2], alpha)
        else:
            backgroundColor = self._get_color(red, green, blue, alpha)
        if not self.spreadsheetId in self.updates:
            self.updates[self.spreadsheetId] = {'requests': []}
        self.updates[self.spreadsheetId]['requests'].append({
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
        })

    def set_color(self, sheetId = None, startColumnIndex = 1, startRowIndex = 1, endColumnIndex = None, endRowIndex = None, color = None, red=0, green=0, blue=0, alpha=0):
        sheetId = self._get_sheet_id(sheetId)
        if color and color in COLORS:
            color = self._get_color(COLORS[color][0], COLORS[color][1], COLORS[color][2], alpha)
        else:
            color = self._get_color(red, green, blue, alpha)
        if not self.spreadsheetId in self.updates:
            self.updates[self.spreadsheetId] = {'requests': []}
        self.updates[self.spreadsheetId]['requests'].append({
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
        })

    def set_format(self, sheetId = None, startColumnIndex = 1, startRowIndex = 1, endColumnIndex = None, endRowIndex = None, type = None, pattern = None):
        sheetId = self._get_sheet_id(sheetId)
        numberFormat = {}
        if type in NUMBER_FORMATS:
            numberFormat['type'] = type
        if pattern:
            numberFormat['pattern'] = pattern
        if not self.spreadsheetId in self.updates:
            self.updates[self.spreadsheetId] = {'requests': []}
        self.updates[self.spreadsheetId]['requests'].append({
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
        })

    def set_alignment(self, sheetId = None, startColumnIndex = 1, startRowIndex = 1, endColumnIndex = None, endRowIndex = None, horizontalAlignment = None, verticalAlignment = None):
        sheetId = self._get_sheet_id(sheetId)
        userEnteredFormat = {}
        if horizontalAlignment in HORIZONTAL_ALIGNMENT:
            userEnteredFormat['horizontalAlignment'] = horizontalAlignment
        if verticalAlignment in VERTICAL_ALIGNMENT:
            userEnteredFormat['verticalAlignment'] = verticalAlignment
        if not self.spreadsheetId in self.updates:
            self.updates[self.spreadsheetId] = {'requests': []}
        self.updates[self.spreadsheetId]['requests'].append({
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
        })

    def set_wrap(self, sheetId = None, startColumnIndex = 1, startRowIndex = 1, endColumnIndex = None, endRowIndex = None, wrapStrategy = 'OVERFLOW_CELL'):
        sheetId = self._get_sheet_id(sheetId)
        userEnteredFormat = {}
        if wrapStrategy in WRAP_STRATEGY:
            userEnteredFormat['wrapStrategy'] = wrapStrategy
        if not self.spreadsheetId in self.updates:
            self.updates[self.spreadsheetId] = {'requests': []}
        self.updates[self.spreadsheetId]['requests'].append({
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
        })

    def set_borders(self):
        pass

    def insert_rows(self, sheetId = None, startIndex = 1, numRows = 1, inheritFromBefore = True):
        sheetId = self._get_sheet_id(sheetId)
        if not self.spreadsheetId in self.updates:
            self.updates[self.spreadsheetId] = {'requests': []}
        self.updates[self.spreadsheetId]['requests'].append({
            "insertDimension": {
                "inheritFromBefore": inheritFromBefore,
                "range": {
                    "sheetId": sheetId,
                    "dimension": "ROWS",
                    "startIndex": startIndex - 1,
                    "endIndex": startIndex - 1 + numRows,
                }
            }
        })

    def insert_columns(self, sheetId = None, startIndex = 1, numColumns = 1, inheritFromBefore = True):
        sheetId = self._get_sheet_id(sheetId)
        if not self.spreadsheetId in self.updates:
            self.updates[self.spreadsheetId] = {'requests': []}
        self.updates[self.spreadsheetId]['requests'].append({
            "insertDimension": {
                "inheritFromBefore": inheritFromBefore,
                "range": {
                    "sheetId": sheetId,
                    "dimension": "COLUMNS",
                    "startIndex": startIndex - 1,
                    "endIndex": startIndex - 1 + numColumns,
                }
            }
        })

    def delete_rows(self, sheetId = None, startIndex = 1, numRows = 1):
        sheetId = self._get_sheet_id(sheetId)
        if not self.spreadsheetId in self.updates:
            self.updates[self.spreadsheetId] = {'requests': []}
        self.updates[self.spreadsheetId]['requests'].append({
            "deleteDimension": {
                "range": {
                    "sheetId": sheetId,
                    "dimension": "ROWS",
                    "startIndex": startIndex - 1,
                    "endIndex": startIndex - 1 + numRows,
                }
            }
        })

    def delete_columns(self, sheetId = None, startIndex = 1, numColumns = 1):
        sheetId = self._get_sheet_id(sheetId)
        if not self.spreadsheetId in self.updates:
            self.updates[self.spreadsheetId] = {'requests': []}
        self.updates[self.spreadsheetId]['requests'].append({
            "deleteDimension": {
                "range": {
                    "sheetId": sheetId,
                    "dimension": "COLUMNS",
                    "startIndex": startIndex - 1,
                    "endIndex": startIndex - 1 + numColumns,
                }
            }
        })
