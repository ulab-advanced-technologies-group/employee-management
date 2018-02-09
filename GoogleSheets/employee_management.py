from __future__ import print_function
import httplib2
import os

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

from Drive import drive

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Google Sheets API Python Quickstart'
ROSTER = 'mainroster'
ROOT_GROUP = 'ulab'

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
    sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_Id).execute()
    sheets = sheet_metadata.get('sheets', '')
    return sheets

spreadsheet_Id = '1k5OgXFL_o99gbgqD_MJt6LuggL4KRGBI27SIW45-FgQ'
service = main()


# Returns a list of the names of the groups that belong to the person with this SID.
def group_names(SID):
    person = get_person(SID)
    if not person:
        print("Please provide a valid person in the organization.")
        return []
    return list(person.groups)

# Returns a list of Group objects that belong to the given SID.
def groups(SID):
    return [get_group(group_name) for group_name in group_names(SID)]

# Returns a list of SIDs of people who are in the provided group.
def person_from_group(group):
    persons = []
    mainroster = service.spreadsheets().values().get(spreadsheetId=spreadsheet_Id, range=ROSTER).execute()
    values = mainroster.get('values', [])
    try:
        SIDindex = 0
        group_index = values[0].index(group)
        num_rows = len(values)
        num_cols = len(values[0])
        for row_index in range(1, num_rows):
            if values[row_index][group_index] == 'y':
                persons.append(values[row_index][SIDindex])
    except IndexError as e:
        pass
    return persons

# Input n should not be zero-indexed.
def num_to_letter(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string

# Creates a new group with the provided group name if it does not already exist.
# Attaches the group to the provided parent. Creates a group with the name of the
# parent concatenated with the subgroup's name with a dash. To create
# ulab-physics-astro, pass in astro with the parent as ulab-physics.
def create_group(group_name, parent_name='ulab'):
    parent = get_group(parent_name)
    name = parent_name + '-' + group_name
    new_group = get_group(name)
    if new_group:
        print("A group with this name already exists.")
        return False
    if not parent:
        print("Please specify a valid parent for this new group.")
        return False
    new_group = Group(name=name, parent=parent)
    new_group.save_group()
    # Parent got a new subgroup, so we need to save this as well.
    parent.save_group()

    # drive.create_new_directory(name, {}, drive.get_group_id(parent_name))
    return True

######### For Demo purposes

def create_folder_structure(group):
    if group.parent == None:
        drive.create_new_directory(group.name)
    else:
        drive.create_new_directory(group.name, drive.get_group_id(group.parent.name))
    if group.subgroups != set():
        for subgroup in group.subgroups:
            create_folder_structure(get_group(subgroup))
    else:
        folders = {'Test Folder 1', 'Test Folder 2', 'Test Folder 3'}
        for folder in folders:
            drive.create_new_directory(folder, drive.get_group_id(group.name))
    return True

#########

# def get_email_from_SID(SID):
#     mainroster = service.spreadsheets().values().get(spreadsheetId=spreadsheet_Id, range=ROSTER).execute()
#     values = mainroster.get('values', [])
#     SID_index = values[0].index('SID')
#     email_index = values[0].index('email')
#
#     for row_index in range(1, len(values)):
#         if str(SID) == values[row_index][SID_index]:
#             return values[row_index][email_index]


# Returns the total number of groups in the organization.
def total_num_groups() :
    return len(get_all_group_names())

# Returns a list of the names of all the current groups in the organization.
def get_all_group_names() :
    groups = []
    mainroster = service.spreadsheets().values().get(spreadsheetId=spreadsheet_Id, range=ROSTER).execute()
    values = mainroster.get('values', [])
    try :
        groupIndexStart = group_start_index()
        for group_index in range(groupIndexStart, len(values[0])) :
            groups.append(values[0][group_index])
    except:
        pass # Ignores sheets
    return groups

# Returns the column index of the root group on the main roster. This group indicates the beginning of the groups.
def group_start_index() :
    mainroster = service.spreadsheets().values().get(spreadsheetId=spreadsheet_Id, range=ROSTER).execute()
    values = mainroster.get('values', [])
    try :
        for group_index in range(0, len(values[0])) :
            if values[0][group_index].find(ROOT_GROUP) != -1 :
                break
    except :
        pass # Ignores Sheets
    return group_index

# Removes the group with the given title. Does not allow removal of the root group.
def remove_group(group_name):
    if group_name == ROOT_GROUP:
        print("The root group cannot be removed.")
        return False
    else:
        group = get_group(group_name)
        if not group:
            print("Please provide a valid group.")
            return False
        group.remove_group()
        parent = group.parent
        # Commit the changes made to the parent group.
        parent.save_group()

        drive.delete.directory(group_name)

        return True

# Returns the sheet id for the given sheet title. Returns -1 if no sheet title matches.
def get_sheetid(sheet_title):
    sheets = get_sheets(service)
    for sheet in sheets:
        title = sheet.get("properties", {}).get("title", "Sheet1")
        if title == sheet_title:
            sheetId = sheet.get("properties", {}).get("sheetId")
            return sheetId
    return -1

def check_fields(fields):
    # if len(fields) != len(Person.FIELDS):
    #     return False
    # else:
    required_fields = set([Person.SID, Person.FIRST_NAME, Person.LAST_NAME, Person.EMAIL, Person.SUPERVISOR, Person.USERNAME])
    for field_name in fields:
        if field_name in required_fields and not fields[field_name]:
            # Missing some required information.
            return False
    return True

# Call this function with fields being a dictionary with all of Person.FIELDS included as keys.
def add_person_to_mainroster(fields):
    # Checks to make sure we have the required fields.
    person_fields = fields.copy()
    if not check_fields(person_fields):
        print("Fields are not inputted properly.")
        return False

    # get_person takes in an integer SID.
    new_person = get_person(int(person_fields[Person.SID]))
    if new_person:
        print("This person already exists.")
        return False
    new_person = Person(person_fields)
    new_person.save_person()
    return True

def add_person_to_group(SID, role, group):
    person = get_person(SID)
    group_id = drive.get_group_id(group)
    if not person:
        print("Please specify a proper person.")
        return False
    group = get_group(group)
    if not group:
        print("Please specify a proper group.")
        return False
    group.add_person_to_group(person, role)
    drive.add_permissions(person.person_fields[Person.emailAddress], group.name)
    person.save_person()
    return True
    # # Failure Cases
    # # 1. SID isn't in ulab
    # mainroster = service.spreadsheets().values().get(spreadsheetId=spreadsheet_Id, range='mainroster').execute()
    # main_values = mainroster.get('values', [])
    # main_SID_index = main_values[0].index('SID')
    # SID_list_main = []
    # for row_index in range(1, len(main_values)) :
    #     SID_list_main.append(main_values[row_index][main_SID_index])
    # if str(SID) not in SID_list_main :
    #     print('Person does not exist in ulab. Add them to ulab using the add_person_to_mainroster method.')
    #     return
    # # 2. Person already exists in the group (SID and role are same)
    # result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_Id, range=group).execute()
    # group_values = result.get('values', [])
    # group_SID_index = group_values[0].index('SID')
    # group_roles_index = group_values[0].index('Role')
    # for row_index in range (1, len(group_values)) :
    #     if group_values[row_index][group_SID_index] == str(SID) and group_values[row_index][group_roles_index] == str(role) :
    #         print('Person already exists in this group with the requested SID and role.')
    #         return
    # SID_column = []
    # count = 2

    # if group_values[1][group_SID_index] != '':
    #     for row_index in range(1, len(group_values)):
    #         SIDvalue = group_values[row_index][group_SID_index]
    #         SID_column.append(SIDvalue)
    #         count += 1

    # SID_column.append(SID)
    # body = {
    #     'valueInputOption': "USER_ENTERED",
    #     "data": [
    #         {
    #             "range": group + "!A2:A" + str(count),
    #             "majorDimension": "COLUMNS",
    #             "values": [SID_column]
    #         }
    #     ]
    # }
    # #### Replacing Role Column
    # role_column = []
    # count = 2

    # if group_values[1][group_roles_index] != '':
    #     for row_index in range(1, len(group_values)) :
    #         roles = group_values[row_index][group_roles_index]
    #         role_column.append(roles)
    #         count += 1

    # role_column.append(role)
    # body2 = {
    #     'valueInputOption': "USER_ENTERED",
    #     "data": [
    #         {
    #             "range": group + "!B2:B" + str(count),
    #             "majorDimension": "COLUMNS",
    #             "values": [role_column]
    #         }
    #     ]
    # }

    # # replace 'n' w/ 'y'
    # main_group_index = main_values[0].index(group)
    # num_cols = len(main_values[0])
    # letter = num_to_letter(num_cols)
    # for row_index in range(1, len(main_values)):
    #     if str(SID) == main_values[row_index][main_SID_index]:
    #         row = row_index
    #         main_values[row_index][main_group_index] = 'y'
    #         rangeName = '!A' + str(row_index) + ':' + letter + str(row_index)
    #         break

    # body3 = {
    #     'valueInputOption': "USER_ENTERED",
    #     "data": [
    #         {
    #             "range": 'mainroster',
    #             "majorDimension": "ROWS",
    #             "values": main_values
    #         }
    #     ]
    # }

    # replaceSIDs = service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet_Id, body=body).execute()
    # replaceRoles = service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet_Id, body=body2).execute()
    # replaceMain = service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet_Id, body=body3).execute()

def del_person_from_group(SID, group):
    person = get_person(SID)
    if not person:
        print("Please specify a proper person.")
        return False
    group = get_group(group)
    if not group:
        print("Please specify a proper group.")
        return False
    group.remove_person_from_group(person)
    person.save_person()
    return True

    # result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_Id, range=group).execute()
    # values = result.get('values', [])
    # SID_index = values[0].index('SID')
    # role_index = values[0].index('Role')
    # subgroup_index = values[0].index("Subgroups")
    # #### Recursive Call
    # for row_index in range(1,len(values)) :
    #     if values[1][subgroup_index] != '' : # Checks if there is a subgroup
    #         subgroup_title = values[row_index][subgroup_index]
    #         print(subgroup_title)
    #         del_person_from_group(SID, subgroup_title)
    # #### Replacing SID column
    # SID_column = []
    # count = 1
    # for row_index in range(1, len(values)):
    #     SIDvalue = values[row_index][SID_index]
    #     SID_column.append(SIDvalue)
    #     count += 1

    # if str(SID) in SID_column:
    #     SID_column.remove(str(SID))
    #     SID_column.append('')
    # else :
    #     print('A person of this SID does not exist in the group.')
    # body = {
    #     'valueInputOption': "USER_ENTERED",
    #     "data": [
    #         {
    #             "range": group + "!A2:A" + str(count),
    #             "majorDimension": "COLUMNS",
    #             "values": [SID_column]
    #         }
    #     ]
    # }
    # #### Replacing Role Column
    # role_column = []
    # count = 1
    # if values[1][role_index] != '':
    #     for row_index in range(1, len(values)) :
    #         roles = values[row_index][role_index]
    #         role_column.append(roles)
    #         SIDvalue = values[row_index][SID_index]
    #         if SIDvalue == str(SID) :
    #             role = values[row_index][role_index]
    #             role_column.remove(role)
    #             role_column.append('')
    #         count += 1

    # body2 = {
    #     'valueInputOption': "USER_ENTERED",
    #     "data": [
    #         {
    #             "range": group + "!B2:B" + str(count),
    #             "majorDimension": "COLUMNS",
    #             "values": [role_column]
    #         }
    #     ]
    # }

    # # replace 'y' w/ 'n'
    # mainroster = service.spreadsheets().values().get(spreadsheetId=spreadsheet_Id, range='mainroster').execute()
    # main_values = mainroster.get('values', [])
    # main_SID_index = main_values[0].index('SID')
    # main_group_index = main_values[0].index(group)
    # num_cols = len(main_values[0])
    # letter = num_to_letter(num_cols)
    # for row_index in range(1, len(main_values)):
    #     if str(SID) == main_values[row_index][main_SID_index]:
    #         row = row_index
    #         main_values[row_index][main_group_index] = 'n'
    #         rangeName = '!A' + str(row_index) + ':' + letter + str(row_index)
    #         break

    # body3 = {
    #     'valueInputOption': "USER_ENTERED",
    #     "data": [
    #         {
    #             "range": 'mainroster',
    #             "majorDimension": "ROWS",
    #             "values": main_values
    #         }
    #     ]
    # }

    # replaceSIDs = service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet_Id, body=body).execute()
    # replaceRoles = service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet_Id, body=body2).execute()
    # replaceMain = service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet_Id, body=body3).execute()

def del_person_from_ulab(SID):
    person = get_person(SID)
    if not person:
        print("This person does not exist.")
        return False
    return person.remove_person()


    # del_person_from_group(SID, 'ulab')
    # mainrosterid = get_sheetid('mainroster')
    # mainroster = service.spreadsheets().values().get(spreadsheetId=spreadsheet_Id, range='mainroster').execute()
    # values = mainroster.get('values', [])
    # roworder = 0
    # requests = []
    # try:
    #     SIDindex = values[0].index("SID")
    #     for row in values:
    #         if str(SID) == row[SIDindex]:
    #             requests.append({
    #                 'deleteDimension': {
    #                     'range': {
    #                         'sheetId': mainrosterid,
    #                         'dimension': 'ROWS',
    #                         'startIndex': roworder,
    #                         'endIndex': roworder + 1,
    #                     }
    #                 }
    #             }
    #         )
    #         roworder += 1
    # except:
    #     pass # Ignores sheets that do not have SID as a column
    # if requests:
    #     body = {'requests': requests}
    #     response = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_Id, body=body).execute()

# Returns the Group object corresponding to the given group name. If there
# is no group of this name, returns None. This function is used in the Group class as well.
# Needs to create a Group instance according to the constructor defined in the Group class and return that instance.
# The parent field of Group is another Group and the subgroups field is a list of subgroup names.
# So, to get the Group <group_name> you would also need to get the parent group.
# (Just call get_group on the parent group name for this.)
def get_group(group_name, parent=None):
    if get_sheetid(group_name) == -1:
        return None
    group_sheet = service.spreadsheets().values().get(spreadsheetId=spreadsheet_Id, range=group_name).execute()
    values = group_sheet.get('values', [])
    if not values:
        return None
    SID_index = values[0].index('SID')
    role_index = values[0].index('Role')
    subgroup_index = values[0].index("Subgroups")
    parent_index = values[0].index("Parent")
    persons = {}
    subgroups = set()
    for row in range(1, len(values)):
        sid = values[row][SID_index]
        role = values[row][role_index]
        subgroup = ""
        if len(values[row]) > 2:
            subgroup = str(values[row][subgroup_index])
        if subgroup:
            subgroups.add(subgroup)
        if sid:
            persons[sid] = role
    if group_name == ROOT_GROUP:
        group = Group(group_name, persons, subgroups, True)
    else:
        parent_name = values[1][parent_index]
        if parent and isinstance(parent, Group):
            group = Group(group_name, persons, subgroups, True, parent)
        else:
            group = Group(group_name, persons, subgroups, True, get_group(parent_name))
    return group


# Returns the Person object corresponding to the given SID. If there is no person with this SID, returns None.
# Needs access to the spreadsheets. Needs to create a Person instance according to the constructor defined in the Person class.
# For now just create the Person instance with basic field information (name, sid, email, etc.). We can add other fields later.
def get_person(SID):
    mainroster = service.spreadsheets().values().get(spreadsheetId=spreadsheet_Id, range=ROSTER).execute()
    values = mainroster.get('values', [])
    if not values:
        return None
    person_fields = {}
    for row_index in range(1, len(values)):
        # SID is at the first column.
        if str(SID) == values[row_index][0]:
            # To accommodate all the necessary fields, make a list of fields instead in the Person class.
            # Replace with dictionary.get() calls with specified default values. Assumes that the order of the fields
            # in Person.FIELDS matches the order of the columns
            for i in range(len(Person.FIELDS)):
                field = values[row_index][i]
                # Remove leading and trailing whitespace.
                stripped_field = field.strip()
                list_fields = set([Person.CLASSES, Person.MAJORS, Person.ACCESSES, Person.ACCOUNTS])
                field_name = Person.FIELDS[i]
                if stripped_field:
                    if field_name in list_fields:
                        person_fields[field_name] = stripped_field.split(",")
                    else:
                        person_fields[field_name] = stripped_field
            groups = set()
            groupIndexStart =  group_start_index()
            for group_index in range(groupIndexStart, len(values[row_index])):
                cell = values[row_index][group_index]
                if cell == 'y':
                    groups.add(values[0][group_index])
            return Person(person_fields, True, groups)
    return None

# Returns a list of person objects given a list of SIDs. More efficient than calling get_person n times.
def batch_get_persons(SIDs):
    mainroster = service.spreadsheets().values().get(spreadsheetId=spreadsheet_Id, range=ROSTER).execute()
    values = mainroster.get('values', [])
    if not values:
        return []
    res = []
    for SID in SIDs:
        person_fields = {}
        for row_index in range(1, len(values)):
            # SID is at the first column.
            if str(SID) == values[row_index][0]:
                # To accommodate all the necessary fields, make a list of fields instead in the Person class.
                # Replace with dictionary.get() calls with specified default values. Assumes that the order of the fields
                # in Person.FIELDS matches the order of the columns
                for i in range(len(Person.FIELDS)):
                    field = values[row_index][i]
                    # Remove leading and trailing whitespace.
                    stripped_field = field.strip()
                    field_name = Person.FIELDS[i]
                    if stripped_field:
                        person_fields[field_name] = stripped_field
                groups = set()
                groupIndexStart =  group_start_index()
                for group_index in range(groupIndexStart, len(values[row_index])):
                    cell = values[row_index][group_index]
                    if cell == 'y':
                        groups.add(values[0][group_index])
                res.append(Person(person_fields, True, groups))
                break
    return res

"""
During tree traversals of the group structure, Group objects will not be
created until needed. The subgroups field of Groups will store string names
of subgroups and when thbe actual Group object is needed, we call get_group on
that string name.

After making modifications to a group, use the save() function to commit the changes
to the spreadsheet.
"""

class Group:

    def __init__(self, name, people={}, subgroups=set(), exists=False, parent=None):
        self.name = name
        self.subgroups = subgroups
        # Stored as a dictionary of SIDs to roles, as Person objects here
        # will take up a considerable amount of space. This dictionary
        # should only be populated at leaf node groups.
        self.people = people.copy()
        self.exists = exists
        self.parent = parent
        self.googlegroup = 'example@googlegroups.com'

    # Returns whether this group is a leaf or not. The group is a leaf only if there are no subgroups.
    def isLeaf(self):
        return not bool(self.subgroups)

    # Takes in a string name of the subgroup we are adding as a child of this group.
    def add_subgroup(self, subgroup):
        if subgroup not in self.subgroups:
            self.subgroups.add(subgroup)

    def remove_subgroup(self, subgroup):
        if subgroup in self.subgroups:
            self.subgroups.remove(subgroup)

    # Returns a list of this group's subgroups as Group objects.
    def get_subgroups(self):
        if self.isLeaf():
            return []
        subgroups = []
        for group_name in self.subgroups:
            subgroups.append(get_group(group_name, self))
        return subgroups


    # Adds a person to the group if they are not already in it. People can only be added to groups
    # that are leaves in the ULAB organization. Traverse the path back to the root and add the person
    # to each ancestor group's respective column on mainroster. Returns true if the person is added successfully.
    # Needs to call save person after this function call to commit the person to their groups.
    def add_person_to_group(self, person, role):
        if not person or not isinstance(person, Person):
            print("Please provide a proper person.")
            return False
        if not self.isLeaf():
            print("Please specify a more specific subgroup to add this person to.")
            return False
        if person.person_fields[Person.SID] in self.people:
            print("This person already exists in the group.")
            return True
        else:
            self.people[person.person_fields[Person.SID]] = role
            group = self
            while group.parent != None:
                if group.name not in person.groups:
                    person.groups.add(group.name)
                group = group.parent
            drive.add_permissions(person.person_fields[Person.emailAddress], self.name)
            self.save_group()
            return True

    # Removes a person from this group if they are in the group. If this group is a leaf, the person is only removed from this leaf group.
    # If this group is an internal node, we remove the person from all the subgroups as well. Returns true if the remove is successful and
    # returns false otherwise. Needs to call save person on the person removed from this group.
    def remove_person_from_group(self, person):
        if not person or not isinstance(person, Person):
            print("Please provide a proper person.")
            return False
        drive.remove_permissions(person.person_fields[Person.emailAddress], self.name)
        if self.isLeaf():
            if person.person_fields[Person.SID] not in self.people:
                print("This person does not exist in the group.")
                return True
            else:
                if self.name in person.groups:
                    person.groups.remove(self.name)
                self.people.pop(person.person_fields[Person.SID], None)
                self.save_group()
                return True
        else:
            # Remove this group from the set of group names for the person.
            if self.name in person.groups:
                person.groups.remove(self.name)
            # Recursively remove the person from subgroups.
            for subgroup in self.get_subgroups():
                if subgroup and isinstance(subgroup, Group):
                    subgroup.remove_person_from_group(person)
            return True

    # Still need to call save group on the parent after this call in order to commit the removal of the subgroup.
    def remove_group(self):
        if self.name == ROOT_GROUP:
            print("Root group cannot be removed.")
            return
        if get_sheetid(self.name) == -1:
            print("Group has no corresponding sheet.")
            return
        group_sheet = service.spreadsheets().values().get(spreadsheetId=spreadsheet_Id, range=self.name).execute()
        values = group_sheet.get('values', [])
        for subgroup in self.get_subgroups():
            if subgroup and isinstance(subgroup, Group):
                subgroup.remove_group()
        self.parent.remove_subgroup(self.name)
        delsheetrequest = [
            {
            "deleteSheet": {
                "sheetId": get_sheetid(self.name),
                }
            }
        ]
        mainroster = service.spreadsheets().values().get(spreadsheetId=spreadsheet_Id, range=ROSTER).execute()
        title_row = mainroster.get('values', [])[0]
        column_index = title_row.index(self.name)
        delmaincolumnreq = [{
                "deleteDimension": {
                    "range": {
                      "sheetId": get_sheetid(ROSTER),
                      "dimension": "COLUMNS",
                      "startIndex": column_index,
                      "endIndex":   column_index + 1
                    }
                }
            }
        ]
        delete_group_body = {"requests": delsheetrequest}
        delete_group_column = {"requests": delmaincolumnreq}

        delete_group_response = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_Id, body=delete_group_body).execute()
        delete_group_column_response = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_Id, body=delete_group_column).execute()


    # This function will be called to save changes made to a group. It is essentially a commit
    # after the user does a series of transactions on the Group object.
    def save_group(self):
        # If the group is not already created, then we make a new sheet for it and fill it in.
        if not self.exists:
            requests = []
            title_row = ["SID", "Role", "Subgroups", "Parent"]
            empty_values = ['', '', '', self.parent.name]
            requests.append({
                    "addSheet": {
                        "properties": {
                          "title": self.name,
                        }
                    }
                }
            )
            data = [
                {
                    'range': self.name + "!A1:D2",
                    'majorDimension': 'ROWS',
                    'values': [title_row, empty_values]
                }
            ]
            new_group_body = {
                'valueInputOption': "USER_ENTERED",
                'data': data
            }
            if requests:
                body = {'requests': requests}
                add_sheet_response = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_Id, body=body).execute()
                add_group_template_response = service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet_Id, body=new_group_body).execute()
            if self.parent:
                self.parent.add_subgroup(self.name)

            # Add group to mainroster column.
            mainroster = service.spreadsheets().values().get(spreadsheetId=spreadsheet_Id, range=ROSTER).execute()
            values = mainroster.get('values', [])
            num_rows = len(values)
            num_cols = len(values[0])
            # Fill the default values with the student number of n's to indicate no one is a part of this group yet.
            default_values = [self.name]
            letter = num_to_letter(num_cols + 1)
            for i in range(num_rows - 1):
                default_values.append('n')

            rangeName = 'mainroster!' + letter + '1:' + letter + str(num_rows)
            new_group_column_body = {
                    'valueInputOption': "USER_ENTERED",
                    "data": [
                        {
                            "range": rangeName,
                            "majorDimension": "COLUMNS",
                            "values": [
                                default_values
                            ]
                        }
                    ]
            }
            new_group_column_response = service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet_Id, body=new_group_column_body).execute()

            addemptycolumnreq = [{
                "appendDimension": {
                    "sheetId": get_sheetid(ROSTER),
                    "dimension": "COLUMNS",
                    "length": 1
                    }
                }
            ]
            append_empty_body = {"requests": addemptycolumnreq}
            append_empty_column_response = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_Id, body=append_empty_body).execute()

        SID_column = ['SID']
        role_column = ['Role']
        subgroup_column = ['Subgroups'] + list(self.subgroups)
        parent = ['Parent', self.parent.name if self.parent else '']
        for SID in self.people:
            SID_column.append(SID)
            role_column.append(self.people[SID])
        # To account for cases where either we delete a subgroup, or delete a person from this group.
        subgroup_column += ['']
        SID_column += ['']
        role_column += ['']
        longest_col = max([len(SID_column), len(role_column), len(subgroup_column), len(parent)])
        body = {
            'valueInputOption': "USER_ENTERED",
            "data": [
                {
                    "range": self.name + "!A1:D" + str(longest_col),
                    "majorDimension": "COLUMNS",
                    "values": [SID_column, role_column, subgroup_column, parent]
                }
            ]
        }
        overWriteGroup = service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet_Id, body=body).execute()

    def __repr__(self):
        return " ".join([word.capitalize() for word in self.name.split("-")])

"""
After making modifications to a person, use the save() function to commit the changes
to the spreadsheet.
"""

class Person:
    # Define fields for the Person type here.
    SID  = 'SID'
    FIRST_NAME = 'first_name'
    MIDDLE_NAME = 'middle_name'
    LAST_NAME = 'last_name'
    USERNAME = 'calnet'
    SUPERVISOR = 'supervisor'
    EMAIL = 'email'
    PHONE_NUMBER = 'phone_no'
    MAJORS = 'majors'
    BACK_ID = 'back_id'
    DIETARY_PREFERENCES = 'diet_preference'
    EXPECTED_GRADUATION = 'expected_grad'
    CLASSES = 'classes'
    GOALS = 'goals'
    ACCOUNTS = 'accounts'
    ACCESSES = 'accesses'

    FIELDS = [SID, FIRST_NAME, MIDDLE_NAME, LAST_NAME, USERNAME, SUPERVISOR, EMAIL, PHONE_NUMBER, MAJORS, BACK_ID, DIETARY_PREFERENCES,
                EXPECTED_GRADUATION, CLASSES, GOALS, ACCOUNTS, ACCESSES]

    HUMAN_FRIENDLY_FIELDS = {SID: 'SID', FIRST_NAME: 'First Name', MIDDLE_NAME: 'Middle Name', LAST_NAME: 'Last Name',
                            USERNAME: 'CalNet', SUPERVISOR: 'Supervisor', EMAIL: 'Email', PHONE_NUMBER: 'Phone Number',
                            MAJORS: 'Majors', BACK_ID: 'Back ID', DIETARY_PREFERENCES: 'Dietary Preferences',
                            EXPECTED_GRADUATION: 'Expected Graduation', CLASSES: 'Classes', GOALS: 'Goals',
                            ACCOUNTS: 'Accounts', ACCESSES: 'Accesses'}

    def __init__(self, person_fields={}, exists=False, groups=set([ROOT_GROUP])):
        self.person_fields = {}
        self.person_fields[Person.SID] = person_fields.get(Person.SID, "-1")
        self.person_fields[Person.FIRST_NAME] = person_fields.get(Person.FIRST_NAME, '')
        self.person_fields[Person.MIDDLE_NAME] = person_fields.get(Person.MIDDLE_NAME, '')
        self.person_fields[Person.LAST_NAME] = person_fields.get(Person.LAST_NAME, '')
        self.person_fields[Person.USERNAME] = person_fields.get(Person.USERNAME, '')
        self.person_fields[Person.SUPERVISOR] = person_fields.get(Person.SUPERVISOR, "-1")
        self.person_fields[Person.EMAIL] = person_fields.get(Person.EMAIL, 'defaultemail@email.com')
        self.person_fields[Person.PHONE_NUMBER] = person_fields.get(Person.PHONE_NUMBER, 'XXX-XXX-XXXX')
        self.person_fields[Person.MAJORS] = person_fields.get(Person.MAJORS, [])
        self.person_fields[Person.BACK_ID] = person_fields.get(Person.BACK_ID, "-1")
        self.person_fields[Person.DIETARY_PREFERENCES] = person_fields.get(Person.DIETARY_PREFERENCES, 'No Dietary Preference')
        self.person_fields[Person.EXPECTED_GRADUATION] = person_fields.get(Person.EXPECTED_GRADUATION, 'Default Graduation')
        self.person_fields[Person.CLASSES] = person_fields.get(Person.CLASSES, [])
        self.person_fields[Person.GOALS] = person_fields.get(Person.GOALS, '')
        self.person_fields[Person.ACCOUNTS] = person_fields.get(Person.ACCOUNTS, [])
        self.person_fields[Person.ACCESSES] = person_fields.get(Person.ACCESSES, [])
        # Stores the string names of the groups that this person is a part of. Using the get_group() in employee-management.py
        # to retrieve the Group object for this group. Since order does not matter, we will use a set.
        self.groups = groups
        self.exists = exists

    # Adds the string name of a group to this person.
    def add_group(self, group) :
        self.groups.append(group)

    def addRole(Role) :
        self.Roles.append(Role)

    def deleteRole(Role) :
        self.Roles.remove(Role)

    def adjustFirstName(Name) :
        self.FirstName = Name

    def adjustLastName(Name) :
        self.LastName = Name

    def adjustMiddleName(Name) :
        self.MiddleName = Name

    def adjustSID(newSID) :
        self.SID = newSID

    def adjustEmail(newEmail):
        self.Email = newEmail

    def adjustPhoneNumber(newPhoneNumber):
        self.PhoneNumber = newPhoneNumber

    def dietaryPreferences(DietaryPreferences):
        self.DietaryPreferences = DietaryPreferences

    def linkSchedule(schedule):
        self.schedule = schedule

    def remove_person(self):
        ulab = get_group(ROOT_GROUP)
        ulab.remove_person_from_group(self)
        mainroster = service.spreadsheets().values().get(spreadsheetId=spreadsheet_Id, range=ROSTER).execute()
        values = mainroster.get('values', [])
        roworder = 1
        requests = []
        SIDindex = 0
        for row_index in range(1, len(values)):
            if str(self.person_fields[Person.SID]) == values[row_index][SIDindex]:
                requests.append({
                    'deleteDimension': {
                        'range': {
                            'sheetId': get_sheetid(ROSTER),
                            'dimension': 'ROWS',
                            'startIndex': roworder,
                            'endIndex': roworder + 1,
                        }
                    }
                }
            )
            roworder += 1
        if requests:
            delete_person_body = {'requests': requests}
            delete_person_response = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_Id, body=delete_person_body).execute()

    # Returns a list of the Person object fields, formatted to be saved to the spreadsheet.
    def _collect_fields(self):
        values = []
        for field_name in Person.FIELDS:
            field = self.person_fields[field_name]
            if type(field) == list:
                value = ",".join(field)
                values.append(value)
            else:
                values.append(field)
        return values

    def save_person(self):
        mainroster = service.spreadsheets().values().get(spreadsheetId=spreadsheet_Id, range=ROSTER).execute()
        values = mainroster.get('values', [])
        title_row = values[0]
        num_rows = len(values)
        num_cols = len(values[0])
        default_values = self._collect_fields()

        group_start = group_start_index()
        letter = num_to_letter(num_cols)
        for i in range(group_start, num_cols):
            if title_row[i] in self.groups:
                default_values.append('y')
            else:
                default_values.append('n')
        row_number = 0
        if not self.exists:
            row_number = num_rows + 1
        else:
            for row_index in range(1, len(values)):
                if str(self.person_fields[Person.SID]) == values[row_index][0]:
                    row_number = row_index + 1
                    break

        rangeName = '!A' + str(row_number) + ':' + letter + str(row_number)
        update_person_body = {
                'valueInputOption': "USER_ENTERED",
                "data": [
                    {
                        "range": ROSTER + rangeName,
                        "majorDimension": "ROWS",
                        "values": [
                            default_values
                        ]
                    }
                ]
        }
        update_person_response = service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet_Id, body=update_person_body).execute()

if __name__ == '__main__':
    main()
