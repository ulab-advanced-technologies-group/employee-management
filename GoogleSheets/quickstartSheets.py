
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

def main():
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
    return service

def get_sheets(service):
    spreadsheetId = '1k5OgXFL_o99gbgqD_MJt6LuggL4KRGBI27SIW45-FgQ'
    sheet_metadata = service.spreadsheets().get(spreadsheetId='1k5OgXFL_o99gbgqD_MJt6LuggL4KRGBI27SIW45-FgQ').execute()
    sheets = sheet_metadata.get('sheets', '')
    return sheets

def groups(SID):
    service = main()
    sheets = get_sheets(service)
    sortedgroups = []
    mainroster = sheets[0].get("properties", {}).get("title", "Sheet1") # Assumes mainroster is 1st sheet
    result = service.spreadsheets().values().get(spreadsheetId='1k5OgXFL_o99gbgqD_MJt6LuggL4KRGBI27SIW45-FgQ', range=mainroster).execute()
    values = result.get('values', [])
    try:
        SIDindex = values[0].index("SID")
        for row_index in range(len(values)):
            if str(SID) == values[row_index][SIDindex]:
                SIDrow = row_index
                break
        groupIndexStart = 5
        for group_index in range(groupIndexStart, len(values[SIDrow])):
            cell = values[SIDrow][group_index]
            if cell == 'y':
                sortedgroups.append(values[0][group_index])
    except:
        pass # Ignores sheets
    return sortedgroups

def create_group(title, parent=''):
    service = main()
    requests = []
    values = [
    ["SID", "Role", "Subgroups", "Parent"], # Additional rows
    ]
    requests.append({
        "addSheet": {
            "properties": {
              "title": title,
            }
            }
        }
    )
    values1= [
    ["SID", "Role", "Subgroups", "Parent"]# Additional rows
    ]
    if parent and type(parent) == str:
        values2 = [["", "", "", parent]]
    data = [
    {
        'range': title + "!A1:D1",
        'values': values1
    },
    {
        'range': title + "!A2:D2",
        'values': values2
    }
    # Additional ranges to update ...
    ]
    body2 = {
    'valueInputOption': "USER_ENTERED",
    'data': data
    }
    if requests:
        body = {'requests': requests}
        response = service.spreadsheets().batchUpdate(spreadsheetId='1k5OgXFL_o99gbgqD_MJt6LuggL4KRGBI27SIW45-FgQ', body=body).execute()
        sheetId = response.get('replies')[0].get('addSheet').get('properties').get('sheetId')
        column = service.spreadsheets().values().batchUpdate(spreadsheetId='1k5OgXFL_o99gbgqD_MJt6LuggL4KRGBI27SIW45-FgQ', body=body2).execute()

def remove_group(title):
    sheetId = get_sheetid(title)
    service = main()
    result = service.spreadsheets().values().get(spreadsheetId='1k5OgXFL_o99gbgqD_MJt6LuggL4KRGBI27SIW45-FgQ', range=title).execute()
    values = result.get('values', [])

    subgroup_index = values[0].index("Subgroups")
    for row_index in range(1,len(values)):
        subgroup_title = values[row_index][subgroup_index]
        if subgroup_title:
            remove_group(subgroup_title)
    parent_index = values[0].index("Parent")
    parent_title = values[1][parent_index]
    parent_sheet = service.spreadsheets().values().get(spreadsheetId='1k5OgXFL_o99gbgqD_MJt6LuggL4KRGBI27SIW45-FgQ', range=parent_title).execute()
    parent_values = parent_sheet.get('values', [])
    # List comprehension to get the values in the subgroup column of the parent if there is a subgroup value in that row.
    parent_subgroup_column = []
    removed = False
    count = 2
    for i in range(1, len(parent_values)):
        row = parent_values[i]
        if len(row) > 2:
            subgroup = row[2]
            if subgroup != title:
                parent_subgroup_column.append(subgroup)
                removed = True
            count += 1
    if removed:
        parent_subgroup_column.append('')
    print(parent_subgroup_column)

    body = {
            'valueInputOption': "USER_ENTERED",
            "data": [
                {
                    "range": parent_title + "!C2:C" + str(count),
                    "majorDimension": "COLUMNS",
                    "values": [
                        parent_subgroup_column
                    ]
                }
            ]
    }
    if sheetId != -1:
        delsheetrequest = [{
            "deleteSheet": {
                "sheetId": sheetId,
                }
            }
        ]
        delmaincolumnreq = [{
            "deleteDimension": {

            }

            }
        ]
        body2 = {"requests": delsheetrequest}
        deleteSubgroups = service.spreadsheets().values().batchUpdate(spreadsheetId='1k5OgXFL_o99gbgqD_MJt6LuggL4KRGBI27SIW45-FgQ', body=body).execute()
        deleteCurgroup = service.spreadsheets().batchUpdate(spreadsheetId='1k5OgXFL_o99gbgqD_MJt6LuggL4KRGBI27SIW45-FgQ', body=body2).execute()

# remove column in mainroster

# Returns the sheet id for the given sheet title. Returns -1 if no sheet title matches.
def get_sheetid(sheet_title):
    service = main()
    sheet_metadata = service.spreadsheets().get(spreadsheetId='1k5OgXFL_o99gbgqD_MJt6LuggL4KRGBI27SIW45-FgQ').execute()
    sheets = sheet_metadata.get('sheets', '')
    for sheet in sheets:
        title = sheet.get("properties", {}).get("title", "Sheet1")
        if title == sheet_title:
            sheetId = sheet.get("properties", {}).get("sheetId")
            return sheetId
    return -1

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
