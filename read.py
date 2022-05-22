from googleapiclient.discovery import build
from google.oauth2 import service_account
import re

def main():
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    SERVICE_ACCOUNT_FILE = 'keys.json'

    creds = None
    creds = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=SCOPES)

    # The ID and range of a sample spreadsheet.
    ssId = '120xqIQG1wcTEhl4_XR2GFQjgNzm2xMnxyNh3Rd-Sric'
    #SAMPLE_RANGE_NAME = 'Class Data!A2:E'

    minRange = "Purchases!A1:A1"
    maxRange = "Purchases!A1:D1000"

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=ssId,
                                range=maxRange).execute()
    values = result.get('values', [])

    test = [["test values"], ["I rewrote this one"]]

    request = sheet.values().update(spreadsheetId=ssId, range="Purchases!A2", valueInputOption="USER_ENTERED", body={"values": test}).execute()

    print(getLastRow(ssId, maxRange, service))

    print(request)

def getLastRow(ssId, maxRange, service):

    #empty table
    table = {
        'majorDimension': 'ROWS',
        'values': []
    }

    #append the empty table
    request = service.spreadsheets().values().append(
    spreadsheetId=ssId,
    range=maxRange,
    valueInputOption='USER_ENTERED',
    insertDataOption='INSERT_ROWS',
    body=table)

    result = request.execute()

    # get last row index
    p = re.compile('^.*![A-Z]+\d+:[A-Z]+(\d+)$')
    match = p.match(result['tableRange'])
    lastrow = match.group(1)

    # lookup the data on the last row
    result = service.spreadsheets().values().get(
        spreadsheetId=ssId,
        range=f'Purchases!A{lastrow}:Z{lastrow}'
    ).execute()

    return lastrow

def addNew(ssId, maxRange, service, info):
    service.spreadsheets().values().append(
    spreadsheetId=ssId,
    range=maxRange,
    valueInputOption='USER_ENTERED',
    insertDataOption='INSERT_ROWS',
    body=info)

if __name__ == "__main__":
    main()

