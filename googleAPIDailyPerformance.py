import os
from datetime import datetime, timedelta

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# This is found in the url of the spread sheet
SPREADSHEET_ID = input('Enter GoogleSheets ID: ')
PARANTSHEET = input('Enter Name of desired sheets: ')

# # Format the date
# # Get the sheet date
# # Get today's date
# today = datetime.today()
# formatted_date = today.strftime("%m/%d %a")
# 
# # Split the date into its individual components
# month, rest = formatted_date.split("/")
# day, weekday = rest.split(" ")
# 
# # Remove leading zeros from month and day
# month = month.lstrip("0")
# day = day.lstrip("0")
# 
# # Rejoin the components together
# date = f"{month}/{day} {weekday}"
# print(date)
# 
# # this is the main sheet
# PARANTSHEET = date

creds = None
MAX_LENGTH = 18
PLACEHOLDER = ''
# The file token.json stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.
# Check if the credentials are available and valid
if not creds or not creds.valid:
    try:
        # Try to refresh the credentials if they have expired
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
    except google.auth.exceptions.RefreshError:
        # If a RefreshError occurs, delete the token.json file
        if os.path.exists('token.json'):
            os.remove('token.json')
            creds = None

    # If credentials are still not available, start the login flow
    if not creds:
        flow = InstalledAppFlow.from_client_secrets_file('client_secret.json', SCOPES)
        creds = flow.run_local_server(port=0)
        # Save the new credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

def getSheet(sheetName, cellRange):
    try:
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        rangeVar = sheetName + cellRange
        result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=rangeVar).execute()
        values = result.get('values', [])

        # Ensure the retrieved list always has the same length and flatten it
        flat_values = []
        if values:
            for row in values:
                while len(row) < MAX_LENGTH:
                    row.append(PLACEHOLDER)
                flat_values.extend(row)

        return flat_values  # Return the flattened list
        
    except HttpError as err:
        print(err)

def flatList(masterList):
    return [name for sublist in masterList for inner_list in sublist for name in inner_list]

def evaluate_performance(data):

    rating = 0

    # Extract times
    planned_rts_time = datetime.strptime(data[7], '%H:%M')
    route_finished_time = datetime.strptime(data[8], '%H:%M')
    actual_end_work = datetime.strptime(data[10], '%H:%M')
    second_last_field = data[-2]
    last_field = data[-1]

    # Calculate time difference in seconds
    # end_work - rts_time
    time_difference = (actual_end_work - planned_rts_time).total_seconds()

    print('time diff: ', time_difference)
    
    # Conditions
    if time_difference <= -1200 and last_field:
        rating = 5
    elif (time_difference <= -1200):
        rating = 4
    elif (abs(time_difference) <= 1200):
        rating = 3
    elif (time_difference <= 2400):
        rating = 2
    else:
        rating = 1

    if second_last_field and rating > 1:
        rating = rating - 1
    if last_field and rating < 5:
        rating = rating + 1

    if rating == 5:
        return "FANTASTIC PLUS"
    elif rating == 4:
        return "FANTASTIC"
    elif rating == 3:
        return "GREAT"
    elif rating == 2:
        return "FAIR"
    else:
        return "POOR"
    

def getMasterList():
    # Initial values
    masterList = []
    row_number = 8  # Starting row

    while True:
        current_range = f'!A{row_number}:R{row_number}'
        current_data = getSheet(PARANTSHEET, current_range)
        
        # If the first cell (name) in the current row is empty, break the loop
        if not current_data or not current_data[2]: 
            break

        try:
            print(current_data)
            print(evaluate_performance(current_data))
            performance = evaluate_performance(current_data)
        except:
            print('skip')

        try:
            service = build('sheets', 'v4', credentials=creds)
            sheet = service.spreadsheets()

            if performance == 'FANTASTIC PLUS':
                sheet.values().update(spreadsheetId=SPREADSHEET_ID, range=f"{PARANTSHEET}!L{row_number}", valueInputOption="USER_ENTERED", body={"values": [['1']]}).execute()
            elif performance == 'FANTASTIC':
                sheet.values().update(spreadsheetId=SPREADSHEET_ID, range=f"{PARANTSHEET}!M{row_number}", valueInputOption="USER_ENTERED", body={"values": [['1']]}).execute()
            elif performance == 'GREAT':
                sheet.values().update(spreadsheetId=SPREADSHEET_ID, range=f"{PARANTSHEET}!N{row_number}", valueInputOption="USER_ENTERED", body={"values": [['1']]}).execute()
            elif performance == 'FAIR':
                sheet.values().update(spreadsheetId=SPREADSHEET_ID, range=f"{PARANTSHEET}!O{row_number}", valueInputOption="USER_ENTERED", body={"values": [['1']]}).execute()
            elif performance == 'POOR':
                sheet.values().update(spreadsheetId=SPREADSHEET_ID, range=f"{PARANTSHEET}!P{row_number}", valueInputOption="USER_ENTERED", body={"values": [['1']]}).execute()
            else:
                print('ERROR WRITING!!!!!')

        except HttpError as err:
            print(err)

        masterList.append(current_data)
        row_number += 1

    return masterList

def main():
    getMasterList()
    # print(evaluate_performance(newList))

if __name__ == '__main__':
    main()
