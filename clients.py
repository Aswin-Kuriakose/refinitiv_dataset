import gspread
from oauth2client.service_account import ServiceAccountCredentials

def getGSpreadClient():
    scope = ['https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive.file', "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(
        'company-profile-url-200.json', scope)
    client = gspread.authorize(creds)
    return client