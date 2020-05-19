from google.oauth2 import service_account

SERVICE_ACCOUNT_FILE = 'google_api/client_secret.json'
SCOPES = ['https://www.googleapis.com/auth/drive']
credentials = service_account.Credentials \
    .from_service_account_file(SERVICE_ACCOUNT_FILE, scopes = SCOPES) \
    .with_subject('ce-info@ce-info.iam.gserviceaccount.com')

COLORS = {
    'black': (0, 0, 0),
    'grey': (0.5, 0.5, 0.5),
    'white': (1, 1, 1),
    'red': (1, 0, 0),
    'light_red': (1, 0.6, 0.6),
    'dark_red': (0.6, 0, 0),
    'green': (0, 1, 0),
    'light_green': (0.6, 1, 0.6),
    'dark_green': (0, 0.6, 0),
    'blue': (0, 0, 1),
    'light_blue': (0.6, 0.6, 1),
    'dark_blue': (0, 0, 0.6),
    'cyan': (0, 1, 1),
    'light_cyan': (0.6, 0.6, 1),
    'dark_cyan': (0, 0.6, 0.6),
    'magenta': (1, 0, 1),
    'light_magenta': (1, 0.6, 1),
    'dark_magenta': (0.6, 0, 0.6),
    'yellow': (1, 1, 0),
    'light_yellow': (1, 1, 0.6),
    'dark_yellow': (0.6, 0.6, 0)
}

NUMBER_FORMATS = {
    'TEXT',
    'NUMBER',
    'PERCENT',
    'CURRENCY',
    'DATE',
    'TIME',
    'DATE_TIME',
    'SCIENTIFIC'
}

HORIZONTAL_ALIGNMENT = {
    'LEFT',
    'CENTER',
    'RIGHT'
}

VERTICAL_ALIGNMENT = {
    'TOP',
    'MIDDLE',
    'BOTTOM'
}

WRAP_STRATEGY = {
    'OVERFLOW_CELL',
    'LEGACY_WRAP',
    'CLIP',
    'WRAP'
}
