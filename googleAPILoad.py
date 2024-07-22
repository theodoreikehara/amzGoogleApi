import os
import csv

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
SPREADSHEET_ID = '10PIE7TnXFJ7WAQhCzujtDsXOdLE_YJUGPbNwuuS3oTY'
CSV_FILE = 'Week45.csv'
SHEET_NAME = 'Week45'  # Update this to the name of the sheet where you want to upload the data

def read_csv(file_path):
    """Reads a CSV file and returns its content as a 2D list."""
    with open(file_path, mode='r') as file:
        reader = csv.reader(file)
        data = list(reader)
    return data

def sheet_exists(sheet_service, sheet_name):
    """Checks if a sheet with the specified name exists."""
    response = sheet_service.get(spreadsheetId=SPREADSHEET_ID).execute()
    sheets = response.get('sheets', '')
    for sheet in sheets:
        if sheet['properties']['title'] == sheet_name:
            return True
    return False

def create_sheet(sheet_service, sheet_name):
    """Creates a new sheet with the specified name."""
    body = {
        'requests': [
            {
                'addSheet': {
                    'properties': {
                        'title': sheet_name
                    }
                }
            }
        ]
    }
    sheet_service.batchUpdate(spreadsheetId=SPREADSHEET_ID, body=body).execute()

def main():
    """Loads data from a CSV file into a Google Sheet."""
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('client_secret.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()

        # Check if sheet exists
        if sheet_exists(sheet, SHEET_NAME):
            print(f"Sheet '{SHEET_NAME}' already exists. No data will be written.")
            return
        else:
            # Create the sheet and then upload the data
            create_sheet(sheet, SHEET_NAME)
            
            # Read CSV file and upload the data to Google Sheet
            csv_data = read_csv(CSV_FILE)
            sheet_range = f'{SHEET_NAME}!A1'
            body = {'values': csv_data}
            result = sheet.values().update(spreadsheetId=SPREADSHEET_ID, range=sheet_range, valueInputOption="USER_ENTERED", body=body).execute()
            print('{0} cells updated.'.format(result.get('updatedCells')))

    except HttpError as err:
        print(err)

if __name__ == '__main__':
    main()