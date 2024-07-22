import os
from datetime import datetime, timedelta
import numpy as np
import time
import csv

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# This is found in the url of the spread sheet

# Week
SPREADSHEET_ID = '1Mw5vaJiwi0b-OEEFG7MuE44EYPkHZ-li3cpYqN5azNk'
BEGIN_DATE = 'Sun 11/05'
CSV_NAME = 'week45MSCO.csv'

creds = None
MAX_LENGTH = 18
PLACEHOLDER = ''
# The file token.json stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.
if os.path.exists('token.json'):
    creds = Credentials.from_authorized_user_file('token.json', SCOPES)
# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file('client_secret.json', SCOPES)
        creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open('token.json', 'w') as token:
        token.write(creds.to_json())

def generate_sheet_id_dates(start_date_str):
    # Get the current year
    current_year = datetime.now().year
    
    # Parse the input date string and set the year to the current year
    # The format here is expecting a day abbreviation followed by the month and day
    start_date = datetime.strptime(start_date_str, '%a %m/%d').replace(year=current_year)
    
    # Generate the list of the next 7 dates in the format 'DayAbbr Month/Day'
    sheet_ids = [
        (start_date + timedelta(days=i)).strftime('%a %m/%d')
        for i in range(7)
    ]
    
    return sheet_ids

# def generate_sheet_id_dates(start_date_str):
#     # Get the current year
#     current_year = datetime.now().year
#     
#     # Parse the input date string and set the year to the current year
#     start_date = datetime.strptime(start_date_str, '%m/%d %a').replace(year=current_year)
#     
#     # Generate the list of the next 14 dates
#     sheet_ids = [
#         "{}/{} {}".format(
#             (start_date + timedelta(days=i)).month, 
#             (start_date + timedelta(days=i)).day, 
#             (start_date + timedelta(days=i)).strftime('%a')
#         ) 
#         for i in range(7)
#     ]
#     
#     return sheet_ids

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

        shortList = [flat_values[2], flat_values[3], flat_values[11], flat_values[12], flat_values[13], flat_values[14], flat_values[15]]

        return shortList  # Return the flattened list
        
    except HttpError as err:
        print(err)
    
def getMasterList(sheet):
    # Initial values
    masterList = []
    row_number = 8  # Starting row

    while True:
        current_range = f'!A{row_number}:R{row_number}'
        try:
            current_data = getSheet(sheet, current_range)
        except:
            break
        
        # If the first cell (name) in the current row is empty, break the loop
        if not current_data or not current_data[0]: 
            break

        print(current_data)

        masterList.append(current_data)
        row_number += 1

    return masterList

def get_rating(scores):
    categories = ["Fantastic Plus", "Fantastic", "Great", "Fair", "Poor"]
    values = [5, 4, 3, 2, 1]

    total_score = sum([a*b for a, b in zip(scores, values)])
    total_ratings = sum(scores)

    if total_ratings == 0:  # to avoid division by zero
        return "No Ratings"

    average = total_score // total_ratings

    # Return the category based on the average value
    return categories[values.index(average)]

def sanitize_data(data):
    return data.replace('"', '').replace(',', '').replace('\n', ' ')

def main():
    start_time = time.time()  # start runtime

    # gets the 1 weeks sheet
    SHEET_ID = generate_sheet_id_dates(BEGIN_DATE)

    driver_data = {}  # To store cumulative data for each driver
    
    count = 0
    for i in SHEET_ID:
        if not count == 0:
            print('Waiting 1 min...')
            time.sleep(60)
        count += 1
        print(i)
        daily_data = getMasterList(i)
        for row in daily_data:
            driver_name = row[1]
            scores = [int(val) if val.isdigit() else 0 for val in row[2:]]
            if driver_name not in driver_data:
                driver_data[driver_name] = scores
            else:
                driver_data[driver_name] = [sum(x) for x in zip(driver_data[driver_name], scores)]

    # Writing to CSV
    with open(CSV_NAME, 'w', newline='') as csvfile:
        csvwriter = csv.writer(csvfile)
        
        # Writing the header row
        csvwriter.writerow(["Name", "Rating"])
        
        # Writing data for each driver
        for driver_name, scores in driver_data.items():
            rating = get_rating(scores)
            
            # Sanitize data before writing
            safe_driver_name = sanitize_data(driver_name)
            safe_rating = sanitize_data(rating)
            
            csvwriter.writerow([safe_driver_name, safe_rating])
    
    runtime_seconds = time.time() - start_time  # Calculate the total runtime in seconds
    runtime_minutes = int(runtime_seconds // 60)
    runtime_seconds %= 60
    print(f"\nRuntime: {runtime_minutes} minutes, {runtime_seconds:.2f} seconds")

if __name__ == '__main__':
    main()