
from __future__ import print_function
import httplib2
import os

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

#try:
#    import argparse
#    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
#except ImportError:
#    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/spreadsheets.readonly'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Google Sheets API Python Quickstart'


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'sheets.googleapis.com-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

def main(SID):
    """Shows basic usage of the Sheets API.

    Creates a Sheets API service object and prints the names and majors of
    students in a sample spreadsheet:
    https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
    """
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http,
                              discoveryServiceUrl=discoveryUrl)

    spreadsheetId = '1k5OgXFL_o99gbgqD_MJt6LuggL4KRGBI27SIW45-FgQ'
    sheet_metadata = service.spreadsheets().get(spreadsheetId='1k5OgXFL_o99gbgqD_MJt6LuggL4KRGBI27SIW45-FgQ').execute()
    sheets = sheet_metadata.get('sheets', '')
    groups = []
    sortedgroups = []

    for sheet in sheets:
        title = sheet.get("properties", {}).get("title", "Sheet1")
        groups.append(title)

    for rangeName in groups:
        result = service.spreadsheets().values().get(spreadsheetId=spreadsheetId, range=rangeName).execute()
        values = result.get('values', [])
        try:
            SIDindex = values[0].index("SID")
            for row in values:
                if str(SID) == row[SIDindex]:
                    sortedgroups.append(rangeName)
        except:
            pass # Ignores sheets that do not have SID as a column
    return sortedgroups

def remove_from_all(SID):
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http,
                              discoveryServiceUrl=discoveryUrl)

    spreadsheetId = '1k5OgXFL_o99gbgqD_MJt6LuggL4KRGBI27SIW45-FgQ'
    sheet_metadata = service.spreadsheets().get(spreadsheetId='1k5OgXFL_o99gbgqD_MJt6LuggL4KRGBI27SIW45-FgQ').execute()
    sheets = sheet_metadata.get('sheets', '')
    groups = []
    for sheet in sheets:
        title = sheet.get("properties", {}).get("title", "Sheet1")
        groups.append(title)

    for group in groups:
        remove_from_group(SID, group)

def remove_from_group(SID, group):
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http,
                              discoveryServiceUrl=discoveryUrl)

    spreadsheetId = '1k5OgXFL_o99gbgqD_MJt6LuggL4KRGBI27SIW45-FgQ'
    sheet_metadata = service.spreadsheets().get(spreadsheetId='1k5OgXFL_o99gbgqD_MJt6LuggL4KRGBI27SIW45-FgQ').execute()
    sheets = sheet_metadata.get('sheets', '')
    groups = []
    sheetIds = []
    requests = []
    for sheet in sheets:
        title = sheet.get("properties", {}).get("title", "Sheet1")
        sheetId = sheet.get("properties", {}).get("sheetId")
        groups.append(title)
        sheetIds.append(sheetId)
    sheetorder = 0

    for rangeName in groups:
        if group == rangeName:
            roworder = 0
            result = service.spreadsheets().values().get(spreadsheetId=spreadsheetId, range=rangeName).execute()
            values = result.get('values', [])
            try:
                SIDindex = values[0].index("SID")
                for row in values:
                    if str(SID) == row[SIDindex]:
                        requests.append({
                            'deleteDimension': {
                                'range': {
                                    'sheetId': sheetIds[sheetorder],
                                    'dimension': 'ROWS',
                                    'startIndex': roworder,
                                    'endIndex': roworder + 1,
                                }
                            }
                        }
                    )
                    roworder += 1
            except:
                pass # Ignores sheets that do not have SID as a column
        sheetorder += 1
    if requests:
        body = {'requests': requests}
        response = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheetId, body=body).execute()

if __name__ == '__main__':
    main()
