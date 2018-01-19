
from __future__ import print_function
import httplib2
import os

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

# try:
#     import argparse
#     flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
# except ImportError:
#     flags = None

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
        num_rows = len(values)
        for row_index in range(1, num_rows):
            if str(SID) == values[row_index][SIDindex]:
                SIDrow = row_index
                break
        # Later revision: iterate through top row until first instance starting with 'ulab' and use that index for group start instead
        groupIndexStart =  group_start_index()
        for group_index in range(groupIndexStart, len(values[SIDrow])):
            cell = values[SIDrow][group_index]
            if cell == 'y':
                sortedgroups.append(values[0][group_index])
    except:
        pass # Ignores sheets
    return sortedgroups

def person_from_group(group):
    persons = []
    service = main()
    sheets = get_sheets(service)
    mainroster = sheets[0].get("properties", {}).get("title", "Sheet1") # Assumes mainroster is 1st sheet
    result = service.spreadsheets().values().get(spreadsheetId='1k5OgXFL_o99gbgqD_MJt6LuggL4KRGBI27SIW45-FgQ', range=mainroster).execute()
    values = result.get('values', [])
    try:
        SIDindex = values[0].index("SID")
        group_index = values[0].index(group)
        num_rows = len(values)
        num_cols = len(values[0])
        for row_index in range(1, num_rows):
            if values[row_index][group_index] == 'y':
                persons.append(values[row_index][SIDindex])
    except:
        pass # Ignores sheets
    return persons

# Input n should not be zero-indexed.
def num_to_letter(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string

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

    #add group to parent's subgroup
    if parent and type(parent) == str and parent != 'mainroster':
        parent_sheet = service.spreadsheets().values().get(spreadsheetId='1k5OgXFL_o99gbgqD_MJt6LuggL4KRGBI27SIW45-FgQ', range=parent).execute()
        parent_values = parent_sheet.get('values', [])
        parent_subgroup_column = []
        count = 2
        for i in range(1, len(parent_values)):
            row = parent_values[i]
            if len(row) > 2:
                subgroup = row[2]
                parent_subgroup_column.append(subgroup)
                count += 1
        parent_subgroup_column.append(title)
        body3 = {
                'valueInputOption': "USER_ENTERED",
                "data": [
                    {
                        "range": parent + "!C2:C" + str(count + 1),
                        "majorDimension": "COLUMNS",
                        "values": [
                            parent_subgroup_column
                        ]
                    }
                ]
        }
        addSubgroup = service.spreadsheets().values().batchUpdate(spreadsheetId='1k5OgXFL_o99gbgqD_MJt6LuggL4KRGBI27SIW45-FgQ', body=body3).execute()
    #add group to mainroster column
    mainrosterid = get_sheetid('mainroster')
    mainroster = service.spreadsheets().values().get(spreadsheetId='1k5OgXFL_o99gbgqD_MJt6LuggL4KRGBI27SIW45-FgQ', range='mainroster').execute()
    values = mainroster.get('values', [])
    num_rows = len(values)
    num_cols = len(values[0])
    default_values = []
    default_values.append(title)
    letter = num_to_letter(num_cols + 1)
    for i in range(num_rows - 1):
        default_values.append('n')
    rangeName = '!' + letter + '1:' + letter + str(num_rows)
    body4 = {
            'valueInputOption': "USER_ENTERED",
            "data": [
                {
                    "range": 'mainroster' + rangeName,
                    "majorDimension": "COLUMNS",
                    "values": [
                        default_values
                    ]
                }
            ]
    }
    addMainColumn = service.spreadsheets().values().batchUpdate(spreadsheetId='1k5OgXFL_o99gbgqD_MJt6LuggL4KRGBI27SIW45-FgQ', body=body4).execute()

def total_num_groups() :
    return len(get_all_groups())

def get_all_groups() :
    service = main()
    groups = []
    sheets = get_sheets(service)
    mainroster = sheets[0].get("properties", {}).get("title", "Sheet1") # Assumes mainroster is 1st sheet
    result = service.spreadsheets().values().get(spreadsheetId='1k5OgXFL_o99gbgqD_MJt6LuggL4KRGBI27SIW45-FgQ', range=mainroster).execute()
    values = result.get('values', [])
    try :
        groupIndexStart = group_start_index()
        for group_index in range(groupIndexStart, len(values[0])) :
            groups.append(values[0][group_index])
    except:
        pass # Ignores sheets
    return groups

def group_start_index() :
    service = main()
    sheets = get_sheets(service)
    mainroster = sheets[0].get("properties", {}).get("title", "Sheet1") # Assumes mainroster is 1st sheet
    result = service.spreadsheets().values().get(spreadsheetId='1k5OgXFL_o99gbgqD_MJt6LuggL4KRGBI27SIW45-FgQ', range=mainroster).execute()
    values = result.get('values', [])
    try :
        for group_index in range(0, len(values[0])) :
            if values[0][group_index].find("ulab") != -1 :
                break
    except :
        pass # Ignores Sheets

    return group_index

def remove_group(title) :
    sheetId = get_sheetid(title)
    if sheetId != -1 :
        service = main()
        result = service.spreadsheets().values().get(spreadsheetId='1k5OgXFL_o99gbgqD_MJt6LuggL4KRGBI27SIW45-FgQ', range=title).execute()
        values = result.get('values', [])

        subgroup_index = values[0].index("Subgroups")
        for row_index in range(1,len(values)) :
            subgroup_title = values[row_index][subgroup_index]
            if subgroup_title :
                remove_group(subgroup_title)
        parent_index = values[0].index("Parent")
        parent_title = values[1][parent_index]
        if parent_title != 'mainroster':
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
            if removed :
                parent_subgroup_column.append('')
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
            deleteSubgroups = service.spreadsheets().values().batchUpdate(spreadsheetId='1k5OgXFL_o99gbgqD_MJt6LuggL4KRGBI27SIW45-FgQ', body=body).execute()

        delsheetrequest = [{
            "deleteSheet": {
                "sheetId": sheetId,
                }
            }
        ]
        mainrosterid = get_sheetid('mainroster')
        mainroster = service.spreadsheets().values().get(spreadsheetId='1k5OgXFL_o99gbgqD_MJt6LuggL4KRGBI27SIW45-FgQ', range='mainroster').execute()
        title_rows = mainroster.get('values', [])[0]
        column_index = title_rows.index(title)
        delmaincolumnreq = [{
            "deleteDimension": {
                "range": {
                  "sheetId": mainrosterid,
                  "dimension": "COLUMNS",
                  "startIndex": column_index,
                  "endIndex":   column_index + 1
                }
            }
            }
        ]
        body2 = {"requests": delsheetrequest}
        body3 = {"requests": delmaincolumnreq}

        deleteCurgroup = service.spreadsheets().batchUpdate(spreadsheetId='1k5OgXFL_o99gbgqD_MJt6LuggL4KRGBI27SIW45-FgQ', body=body2).execute()
        deleteMainColumn = service.spreadsheets().batchUpdate(spreadsheetId='1k5OgXFL_o99gbgqD_MJt6LuggL4KRGBI27SIW45-FgQ', body=body3).execute()

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

def add_person_to_mainroster(firstname, lastname, calnet, SID, email, phonenumber, supervisor):
    service = main()
    mainroster = service.spreadsheets().values().get(spreadsheetId='1k5OgXFL_o99gbgqD_MJt6LuggL4KRGBI27SIW45-FgQ', range='mainroster').execute()
    values = mainroster.get('values', [])
    num_rows = len(values)
    num_cols = len(values[0])
    default_values = [firstname, lastname, calnet, SID, email, phonenumber, supervisor]
    letter = num_to_letter(num_cols)
    for i in range(num_cols - 7): # change number depending column before groups
        default_values.append('n')
    rangeName = '!A' + str(num_rows + 1) + ':' + letter + str(num_rows + 1)
    body = {
            'valueInputOption': "USER_ENTERED",
            "data": [
                {
                    "range": 'mainroster' + rangeName,
                    "majorDimension": "ROWS",
                    "values": [
                        default_values
                    ]
                }
            ]
    }
    addMainrow = service.spreadsheets().values().batchUpdate(spreadsheetId='1k5OgXFL_o99gbgqD_MJt6LuggL4KRGBI27SIW45-FgQ', body=body).execute()

def del_person_from_ulab(SID):
    # remove from group function ----
    service = main()
    mainrosterid = get_sheetid('mainroster')
    mainroster = service.spreadsheets().values().get(spreadsheetId='1k5OgXFL_o99gbgqD_MJt6LuggL4KRGBI27SIW45-FgQ', range='mainroster').execute()
    values = mainroster.get('values', [])
    roworder = 0
    requests = []
    try:
        SIDindex = values[0].index("SID")
        for row in values:
            if str(SID) == row[SIDindex]:
                requests.append({
                    'deleteDimension': {
                        'range': {
                            'sheetId': mainrosterid,
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
    if requests:
        body = {'requests': requests}
        response = service.spreadsheets().batchUpdate(spreadsheetId='1k5OgXFL_o99gbgqD_MJt6LuggL4KRGBI27SIW45-FgQ', body=body).execute()

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
