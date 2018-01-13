
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

def sheets():
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http,
                              discoveryServiceUrl=discoveryUrl)

    spreadsheetId = '1k5OgXFL_o99gbgqD_MJt6LuggL4KRGBI27SIW45-FgQ'
    sheet_metadata = service.spreadsheets().get(spreadsheetId='1k5OgXFL_o99gbgqD_MJt6LuggL4KRGBI27SIW45-FgQ').execute()
    sheets = sheet_metadata.get('sheets', '')

    requests = []
    # Change the spreadsheet's title.
    requests.append({
        'updateSpreadsheetProperties': {
            'properties': {'title': 'Roster2'},
            'fields': 'title'
        }
    })

    body = {'requests': requests}
    response = service.spreadsheets().batchUpdate(spreadsheetId='1k5OgXFL_o99gbgqD_MJt6LuggL4KRGBI27SIW45-FgQ', body=body).execute()
    find_replace_response = response.get('replies')[1].get('findReplace')
    print('{0} replacements made.'.format(
        find_replace_response.get('occurrencesChanged')))

if __name__ == '__main__':
    main()
