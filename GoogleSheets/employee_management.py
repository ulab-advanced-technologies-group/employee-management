from __future__ import print_function
import httplib2
import os

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage
import sys

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
    try:
        sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_Id).execute()
    except Exception as e:
        print("Sheet metadata could not be accessed. Check your connection and rerun the tool. If the issue persists, please contact the Employee Management team in ATG.")
        print(e)
        sys.exit(1)
    sheets = sheet_metadata.get('sheets', '')
    return sheets

spreadsheet_Id = '1k5OgXFL_o99gbgqD_MJt6LuggL4KRGBI27SIW45-FgQ'
service = main()

# Wrapper check for several of the functions below which does some type checking and length checking on SID.
def check_sid(SID):
    return type(SID) == str and len(SID) < 15 and len(SID) > 6

# Returns a list of the names of the groups that belong to the person with this SID.
def group_names(SID):
    if not check_sid(SID):
        print("SID is not inputted properly. Check that it is correct and is a string in quotes.")
        return
    person = get_person(SID)
    if not person:
        print("Please specify a proper person. Please make sure the SID is correct and inputted as a string.")
        return []
    return list(person.groups)

# Returns a list of Group objects that belong to the given SID.
def groups(SID):
    if not check_sid(SID):
        print("SID is not inputted properly. Check that it is correct and is a string in quotes.")
        return
    return [get_group(group_name) for group_name in group_names(SID)]

# Return a list of Group objects that belong to the given SID. More optimal than above version.
def get_persons_groups(SID):
    if not check_sid(SID):
        print("SID is not inputted properly. Check that it is correct and is a string in quotes.")
        return
    group_dict = {}
    all_groups = group_names(SID)
    # Sort the person's groups by how deep it is in the ULAB tree.
    all_groups.sort(key=lambda x: x.count("-"))
    for group_name in all_groups:
        if group_name == ROOT_GROUP:
            group_dict[ROOT_GROUP] = get_group(ROOT_GROUP)
        else:
            # The parent name of this group name is everything in group name before the last hyphen.
            parent_name = group_name.rsplit("-", 1)[0]
            if parent_name in group_dict:
                # Pass in the parent Group object of the group we are trying to get if possible to optimize.
                group_dict[group_name] = get_group(group_name, group_dict[parent_name])
            else:
                group_dict[group_name] = get_group(group_name)
    return list(group_dict.values())

def get_values(sheetRange):
    try:
        sheet = service.spreadsheets().values().get(spreadsheetId=spreadsheet_Id, range=sheetRange).execute()
        values = sheet.get('values', [])
    except Exception as e:
        print("Sheet {} could not be accessed. Check your connection and rerun the tool. If the issue persists, please contact the Employee Management team in ATG.".format(sheetRange))
        print(e)
        sys.exit(1)
    if not values:
        print("Error accessing {}. Please report this to the Employee Management Team.".format(sheetRange))
        return []
    return values

# Returns a list of SIDs of people who are in the provided group.
def person_from_group(group) :
    persons = []
    values = get_values(ROSTER)
    SIDindex = 0
    group_index = values[0].index(group)
    num_rows = len(values)
    num_cols = len(values[0])
    for row_index in range(1, num_rows):
        if values[row_index][group_index] == 'y':
            persons.append(values[row_index][SIDindex])
    return persons

# Input n should not be zero-indexed.
def num_to_letter(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string

def missing_field(field): # Returns list of SIDs
    persons = []
    values = get_values(ROSTER)
    field_index = values[0].index(field)
    SID_index = values[0].index('SID')
    for row_index in range(1, len(values)):
        if values[row_index][field_index] == '':
            SID = values[row_index][SID_index]
            persons.append(SID)
    return persons

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
        print("Please specify a valid parent for this new group. To create astro under the physics group, call create_group('astro', 'ulab-physics').")
        return False
    if parent.hasPeople():
        print("{} is a leaf group and has people. Please remove these people and then create subgroups for {}".format(parent_name, parent_name))
        return False
    if not parent.drive_id:
        print("{} does not have a valid corresponding drive folder. Please check with an admin.")
        return False

    # Need to get the group's id after successfully creating the corresponding drive folder.
    drive.create_new_directory(name, parent.drive_id)
    group_id = drive.get_group_id(name, parent.drive_id)
    drive.create_new_directory('Content', group_id)

    new_group = Group(name=name, parent=parent, drive_id=group_id)
    new_group.save_group()
    # Parent got a new subgroup, so we need to save this as well.
    parent.save_group()
    return True

######### FOR DEMO PURPOSES ONLY

# def create_folder_structure(group):
#     print(group)
#     if group.parent == None:
#         drive.create_new_directory(group.name)
#     else:
#         drive.create_new_directory(group.name, drive.get_group_id(group.parent.name))
#     drive.create_new_directory('Content', drive.get_group_id(group.name))
#     print(group.subgroups)
#     if not group.isLeaf():
#         for subgroup in group.subgroups:
#             print(subgroup)
#             create_folder_structure(get_group(subgroup, group))
#     else:
#         folders = {'Test Folder 1', 'Test Folder 2', 'Test Folder 3'}
#         for folder in folders:
#             print(folder)
#             print('parent', group.parent.name)
#             try :
#                 drive.create_new_directory(folder, drive.get_group_id(group.name, drive.get_group_id(group.parent.name)))
#             except :
#                 print("A drive directory was unable to be created. There is an issue with the API. Please contact the Employee Management team in ATG.")


#         people = list(group.people.keys())
#         people_obj = batch_get_persons(people)
#         emails = [p.person_fields[Person.EMAIL] for p in people_obj if p.person_fields[Person.EMAIL]]

#         permission_group = group
#         while permission_group.parent != None:
#             if permission_group.isLeaf():
#                 for email in emails:
#                     drive.add_permissions(email, permission_group.name)
#                     print(permission_group, email)
#             else:
#                 for email in emails:
#                     try :
#                         drive.add_permissions(email, 'Content', drive.get_group_id(permission_group.parent.name))
#                     except :
#                         print("Drive permissions were unable to be added. There is an issue with the API. Please contact the Employee Management team in ATG.")
#             permission_group = permission_group.parent

#     return True

#########

# Returns a printed tree of the ULAB organization chart. 
def print_tree():
	ulab = get_group(ROOT_GROUP)
	if not ulab:
		print("Could not access the root group.")
		return False
	print(ulab)

# Returns the total number of groups in the organization.
def total_num_groups():
    return len(get_all_group_names())

# Returns a list of the names of all the current groups in the organization.
def get_all_group_names():
    groups = []
    values = get_values(ROSTER)
    try:
        groupIndexStart = group_start_index()
        for group_index in range(groupIndexStart, len(values[0])):
            groups.append(values[0][group_index])
    except IndexError as e:
        print("Error in accessing the values of the sheet. Values are not filled. Check your connection and rerun the tool. Please contact the Employee Management team in ATG if the issue persists.")
        print(e)
        sys.exit(1)
    return groups

# Returns the column index of the root group on the main roster. This group indicates the beginning of the groups.
def group_start_index() :
    values = get_values(ROSTER)
    try:
        for group_index in range(0, len(values[0])):
            if values[0][group_index].find(ROOT_GROUP) != -1:
                break
    except IndexError as e:
        print("Error in accessing the values of the sheet. Values are not filled. Check your connection and rerun the tool. Please contact the Employee Management team in ATG if the issue persists.")
        print(e)
        sys.exit(1)
    return group_index

# Removes the group with the given title. Does not allow removal of the root group.
def remove_group(group_name):
    if group_name == ROOT_GROUP:
        print("The root group cannot be removed.")
        return False
    else:
        group = get_group(group_name)
        if not group:
            print("Please specify a proper group. Please check the format of the group name. The group name for astro is ulab-physics-astro.")
            return False
        group.remove_group()
        parent = group.parent
        if not parent or not parent.drive_id:
            print("This group's parent is not valid. Please specify a correct group.")
        # Commit the changes made to the parent group.
        parent.save_group()

        drive.delete_directory(group_name, parent.drive_id)

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
    required_fields = set([Person.SID, Person.FIRST_NAME, Person.LAST_NAME, Person.EMAIL])
    for field_name in fields:
        if field_name in required_fields and not fields[field_name]:
            # Missing some required information.
            print("Missing a required field.")
            return False
    return True

# Call this function with fields being a dictionary with all of Person.FIELDS included as keys.
def add_person_to_mainroster(fields):
    # Checks to make sure we have the required fields.
    person_fields = fields.copy()
    if not check_fields(person_fields):
        print("Fields are not inputted properly. Make sure to input fields as a dictionary of values with keys like Person.SID. Required fields are SID, first name, last name, and email.")
        return False

    # get_person takes in an integer SID.
    new_person = get_person(int(person_fields[Person.SID]))
    if new_person:
        print("This person already exists in the main roster.")
        return False
    new_person = Person(person_fields)
    new_person.save_person()

    drive.add_permissions(new_person.get_email(), 'Content', drive.get_group_id(ROOT_GROUP))

    return True

def add_person_to_group(SID, role, group_name):
    if not check_sid(SID):
        print("SID is not inputted properly. Check that it is correct and is a string in quotes.")
        return
    if type(role) != str:
    	print("Please input the role as a string.")
    	return
    person = get_person(SID)
    if not person:
        print("Please specify a proper person. Please make sure the SID is correct and inputted as a string.")
        return False
    group = get_group(group_name)
    if not group:
        print("Please specify a proper group. Please check the format of the group name. The group name for astro is ulab-physics-astro.")
        return False
    if not group.add_person_to_group(person, role):
        return False

    person.save_person()
    return True

def add_people_to_group(SIDs, roles, group):
	if len(SIDs) != len(roles):
		print("Please check the number of SIDS and roles passed in.")
		return
	for SID, role in zip(SIDs, roles):
		add_person_to_group(SID, role, group_name)
	return True

def del_person_from_group(SID, group_name):
    if not check_sid(SID):
        print("SID is not inputted properly. Check that it is correct and is a string in quotes.")
        return
    person = get_person(SID)
    if not person:
        print("Please specify a proper person. Please make sure the SID is correct and inputted as a string.")
        return False
    group = get_group(group_name)
    if not group:
        print("Please specify a proper group. Please check the format of the group name. The group name for astro is ulab-physics-astro.")
        return False
    if not group.remove_person_from_group(person):
        return False
    person.save_person()
    return True

def del_person_from_ulab(SID):
    if not check_sid(SID):
        print("SID is not inputted properly. Check that it is correct and is a string in quotes.")
        return
    person = get_person(SID)
    if not person:
        print("This person does not exist. Please make sure the SID is correct and inputted as a string.")
        return False
    return person.remove_person()

# Returns the Group object corresponding to the given group name. If there
# is no group of this name, returns None. This function is used in the Group class as well.
# Needs to create a Group instance according to the constructor defined in the Group class and return that instance.
# The parent field of Group is another Group and the subgroups field is a list of subgroup names.
# So, to get the Group <group_name> you would also need to get the parent group.
# (Just call get_group on the parent group name for this.)
def get_group(group_name, parent=None, checked=False):
    if not checked and get_sheetid(group_name) == -1:
        return None
    values = get_values(group_name)
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
        group = Group(group_name, persons, subgroups, True, None, drive.get_group_id(group_name))
    else:
        parent_name = values[1][parent_index]
        if parent and isinstance(parent, Group):
            group = Group(group_name, persons, subgroups, True, parent, drive.get_group_id(group_name, parent.drive_id))
        else:
            par = get_group(parent_name)
            group = Group(group_name, persons, subgroups, True, par, drive.get_group_id(group_name, par.drive_id))
    return group


# Returns the Person object corresponding to the given SID. If there is no person with this SID, returns None.
# Needs access to the spreadsheets. Needs to create a Person instance according to the constructor defined in the Person class.
# For now just create the Person instance with basic field information (name, sid, email, etc.). We can add other fields later.
def get_person(SID):
    if not check_sid(SID):
        print("SID is not inputted properly. Check that it is correct and is a string in quotes.")
        return
    values = get_values(ROSTER)
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
    check = all([check_sid(SID) for SID in SIDs])
    if not check:
        return
    values = get_values(ROSTER)
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

    def __init__(self, name, people={}, subgroups=set(), exists=False, parent=None, drive_id=None):
        self.name = name
        self.subgroups = subgroups
        # Stored as a dictionary of SIDs to roles, as Person objects here
        # will take up a considerable amount of space. This dictionary
        # should only be populated at leaf node groups.
        self.people = people.copy()
        self.exists = exists
        self.parent = parent
        self.drive_id = drive_id
        self.googlegroup = 'example@googlegroups.com'

    # Returns whether this group is a leaf or not. The group is a leaf only if there are no subgroups.
    def isLeaf(self):
        return not bool(self.subgroups)

    # Returns whether this group has any people stored at it or not. Should only be true for leaf groups.
    def hasPeople(self):
        return self.isLeaf() and len(self.people) > 0

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
            subgroups.append(get_group(group_name, self, True))
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
            return False
        else:
            self.people[person.person_fields[Person.SID]] = role
            group = self
            # Only add the person to the leaf group's corresponding folder.
            try :
                drive.add_permissions(person.get_email(), group.name)
            except :
                print("The drive was unable to add permissions. There is an issue with the API. Please contact the Employee Management team in ATG.")
            while group.parent != None:
                if group.name not in person.groups:
                    person.groups.add(group.name)
                # For any folders higher up, only add the person to the upper folder's Content folder.
                # This maintains utmost secrecy between groups. Design should be updated later to give
                # members of ulab who are part of the front office to have full access to the parent folder as well.
                if not group.parent.drive_id:
                    print("Parent group {} does not have a corresponding drive folder id.".format(group.parent))
                try :
                    drive.add_permissions(email, 'Content', group.parent.drive_id)
                except :
                    print("The drive was unable to add permissions. There is an issue with the API. Please contact the Employee Management team in ATG.")
                group = group.parent
            self.save_group()
            return True

    # Removes a person from this group if they are in the group. If this group is a leaf, the person is only removed from this leaf group.
    # If this group is an internal node, we remove the person from all the subgroups as well. Returns true if the remove is successful and
    # returns false otherwise. Needs to call save person on the person removed from this group.
    def remove_person_from_group(self, person):
        if not person or not isinstance(person, Person):
            print("Please provide a proper person.")
            return False
        if self.isLeaf():
            if person.person_fields[Person.SID] not in self.people:
                print("This person does not exist in the group.")
                return True
            else:
                if self.name in person.groups:
                    person.groups.remove(self.name)
                    try :
                        drive.remove_permissions(person.get_email(), self.name, self.parent.drive_id)
                    except :
                        print("The drive was unable to remove permissions. There is an issue with the API. Please contact the Employee Management team in ATG.")

                self.people.pop(person.get_sid(), None)
                self.save_group()
                return True
        else:
            # Recursively remove the person from subgroups.
            for subgroup in self.get_subgroups():
                if subgroup and isinstance(subgroup, Group):
                    subgroup.remove_person_from_group(person)

            # Remove this group from the set of group names for the person.
            if self.name in person.groups:
                person.groups.remove(self.name)
                try :
                    drive.remove_permissions(person.get_email(), "Content", self.parent.drive_id)
                except :
                    print("The drive was unable to remove permissions. There is an issue with the API. Please contact the Employee Management team in ATG.")

            return True

    # Still need to call save group on the parent after this call in order to commit the removal of the subgroup.
    def remove_group(self):
        if self.name == ROOT_GROUP:
            print("Root group cannot be removed.")
            return
        if get_sheetid(self.name) == -1:
            print("Group has no corresponding sheet.")
            return
        values = get_values(self.name)
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
        values = get_values(ROSTER)
        title_row = values[0]
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
        try:
            delete_group_response = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_Id, body=delete_group_body).execute()
            delete_group_column_response = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_Id, body=delete_group_column).execute()
        except Exception as e:
            print("The column was unable to be removed due to the batch not updating in the spreadsheet. There is an issue with the API. Please contact the Employee Management team in ATG.")
            print(e)
            sys.exit(1)

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
                try:
                    add_sheet_response = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_Id, body=body).execute()
                    add_group_template_response = service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet_Id, body=new_group_body).execute()
                except Exception as e:
                    print("The sheet either could not be filled in or could not be added. There is an issue with the API. Please contact the Employee Management team in ATG.")
                    print(e)
                    sys.exit(1)
            if self.parent:
                self.parent.add_subgroup(self.name)
            # Add group to mainroster column.
            values = get_values(ROSTER)
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
            addemptycolumnreq = [{
                "appendDimension": {
                    "sheetId": get_sheetid(ROSTER),
                    "dimension": "COLUMNS",
                    "length": 1
                    }
                }
            ]
            append_empty_body = {"requests": addemptycolumnreq}
            try:
                new_group_column_response = service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet_Id, body=new_group_column_body).execute()
                append_empty_column_response = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_Id, body=append_empty_body).execute()
            except Exception as e:
                print("The default column body could not be created. Check your connection and rerun the tool. If the issue persists, contact the Employee Management team in ATG.")
                print(e)
                sys.exit(1)

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
        try:
            overWriteGroup = service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet_Id, body=body).execute()
        except:
            print("Either the changes of deleting a subgroup or deleting a person could not be accounted for. There is an issue with the API. Please contact the Employee Management team in ATG.s")

    def __repr__(self):
        return " ".join([word.capitalize() for word in self.name.split("-")])

    def __str__(self):
        return self.str_helper()

    def str_helper(self, level=0):
        tree = "\t\t"*level+self.name+"\n"
        for subgroup in self.get_subgroups():
            if subgroup:
                tree += subgroup.str_helper(level+1)
        return tree

"""
After making modifications to a person, use the save() function to commit the changes
to the spreadsheet.
"""

class Person:

    ROLES = {"supervisor": "Supervisor"}

    # Define fields for the Person type here.
    SID  = 'SID'
    FIRST_NAME = 'first_name'
    MIDDLE_NAME = 'middle_name'
    LAST_NAME = 'last_name'
    USERNAME = 'calnet'
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

    FIELDS = [SID, FIRST_NAME, MIDDLE_NAME, LAST_NAME, USERNAME, EMAIL, PHONE_NUMBER, MAJORS, BACK_ID, DIETARY_PREFERENCES,
                EXPECTED_GRADUATION, CLASSES, GOALS, ACCOUNTS, ACCESSES]

    HUMAN_FRIENDLY_FIELDS = {SID: 'SID', FIRST_NAME: 'First Name', MIDDLE_NAME: 'Middle Name', LAST_NAME: 'Last Name',
                            USERNAME: 'CalNet', EMAIL: 'Email', PHONE_NUMBER: 'Phone Number',
                            MAJORS: 'Majors', BACK_ID: 'Back ID', DIETARY_PREFERENCES: 'Dietary Preferences',
                            EXPECTED_GRADUATION: 'Expected Graduation', CLASSES: 'Classes', GOALS: 'Goals',
                            ACCOUNTS: 'Accounts', ACCESSES: 'Accesses'}

    def __init__(self, person_fields={}, exists=False, groups=set([ROOT_GROUP])):
        self.person_fields = {}
        self.person_fields[Person.SID] = person_fields.get(Person.SID, "")
        self.person_fields[Person.FIRST_NAME] = person_fields.get(Person.FIRST_NAME, '')
        self.person_fields[Person.MIDDLE_NAME] = person_fields.get(Person.MIDDLE_NAME, '')
        self.person_fields[Person.LAST_NAME] = person_fields.get(Person.LAST_NAME, '')
        self.person_fields[Person.USERNAME] = person_fields.get(Person.USERNAME, '')
        self.person_fields[Person.EMAIL] = person_fields.get(Person.EMAIL, '')
        self.person_fields[Person.PHONE_NUMBER] = person_fields.get(Person.PHONE_NUMBER, 'XXX-XXX-XXXX')
        self.person_fields[Person.MAJORS] = person_fields.get(Person.MAJORS, [])
        self.person_fields[Person.BACK_ID] = person_fields.get(Person.BACK_ID, "XXXXXX")
        self.person_fields[Person.DIETARY_PREFERENCES] = person_fields.get(Person.DIETARY_PREFERENCES, 'No Dietary Preference')
        self.person_fields[Person.EXPECTED_GRADUATION] = person_fields.get(Person.EXPECTED_GRADUATION, 'N/A')
        self.person_fields[Person.CLASSES] = person_fields.get(Person.CLASSES, [])
        self.person_fields[Person.GOALS] = person_fields.get(Person.GOALS, '')
        self.person_fields[Person.ACCOUNTS] = person_fields.get(Person.ACCOUNTS, [])
        self.person_fields[Person.ACCESSES] = person_fields.get(Person.ACCESSES, [])
        # Stores the string names of the groups that this person is a part of. Using the get_group() in employee-management.py
        # to retrieve the Group object for this group. Since order does not matter, we will use a set.
        self.groups = groups
        self.exists = exists

    # Getter methods.
    def get_sid(self):
        return self.person_fields[Person.SID]
    def get_first_name(self):
        return self.person_fields[Person.FIRST_NAME]
    def get_middle_name(self):
        return self.person_fields[Person.MIDDLE_NAME]
    def get_last_name(self):
        return self.person_fields[Person.LAST_NAME]
    def get_username(self):
        return self.person_fields[Person.USERNAME]
    # We are removing the supervisor column from the sheet. Need to go through this person's groups
    # and collect each of the group's supervisors.
    def get_supervisors(self):
        groups = get_persons_groups(self.get_sid())
        leaf_groups = filter(lambda x: x and x.isLeaf(), groups)
        supervisors = set()
        for leaf in leaf_groups:
            for sid, role in leaf.people:
                if role == Person.ROLES["supervisor"]:
                    supervisors.add(sid)
        return batch_get_persons(list(supervisors))

    def get_email(self):
        return self.person_fields[Person.EMAIL]
    def get_phone(self):
        return self.person_fields[Person.PHONE_NUMBER]
    def get_majors(self):
        return self.person_fields[Person.MAJORS]
    def get_hard_id(self):
        return self.person_fields[Person.BACK_ID]
    def get_dietary_preferences(self):
        return self.person_fields[Person.DIETARY_PREFERENCES]
    def get_expected_grad(self):
        return self.person_fields[Person.EXPECTED_GRADUATION]
    def get_classes(self):
        return self.person_fields[Person.CLASSES]
    def get_goals(self):
        return self.person_fields[Person.GOALS]
    def get_accounts(self):
        return self.person_fields[Person.ACCOUNTS]
    def get_accesses(self):
        return self.person_fields[Person.ACCESSES]

    # Setter methods.
    def set_sid(self, SID):
        # Need a check to make sure that this new SID doesnt already exist.
        self.person_fields[Person.SID] = SID

    def set_first_name(self, first_name):
        self.person_fields[Person.FIRST_NAME] = first_name

    def set_middle_name(self, middle_name):
        self.person_fields[Person.MIDDLE_NAME] = middle_name

    def set_last_name(self, last_name):
        self.person_fields[Person.LAST_NAME] = last_name

    def set_username(self, username):
        self.person_fields[Person.USERNAME] = username

    def set_email(self, email):
        self.person_fields[Person.EMAIL] = email

    def set_phone(self, phone_no):
        self.person_fields[Person.PHONE_NUMBER] = phone_no

    def add_major(self, major):
        majors = self.person_fields[Person.MAJORS]
        if major not in majors:
            majors.append(major)

    def remove_major(self, major):
        majors = self.person_fields[Person.MAJORS]
        if major in majors:
            majors.remove(major)

    def set_hard_id(self, back_id):
        self.person_fields[Person.BACK_ID] = back_id

    def set_dietary_preferences(self, diet):
        self.person_fields[Person.DIETARY_PREFERENCES] = diet

    def set_expected_grad(self, grad):
        self.person_fields[Person.EXPECTED_GRADUATION] = grad

    def add_class(self, course):
        classes = self.person_fields[Person.CLASSES]
        if course not in classes:
            classes.append(course)

    def remove_class(self, course):
        classes = self.person_fields[Person.CLASSES]
        if course in classes:
            classes.remove(course)

    def set_goals(self, goals):
        self.person_fields[Person.GOALS] = goals

    def add_account(self, account):
        accounts = self.person_fields[Person.ACCOUNTS]
        if account not in accounts:
            accounts.append(account)

    def remove_account(self, account):
        accounts = self.person_fields[Person.ACCOUNTS]
        if account in accounts:
            accounts.remove(account)

    def add_access(self, access):
        accesses = self.person_fields[Person.ACCESSES]
        if access not in accesses:
            accesses.append(access)

    def remove_access(self, access):
        accesses = self.person_fields[Person.ACCESSES]
        if access in accesses:
            accesses.remove(access)

    # Adds the string name of a group to this person.
    def add_group(self, group) :
        self.groups.add(group)

    def remove_person(self):
        ulab = get_group(ROOT_GROUP)
        ulab.remove_person_from_group(self)
        values = get_values(ROSTER)
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
            try :
                delete_person_response = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_Id, body=delete_person_body).execute()
            except :
                print("The person was not deleted due to an issue with the batch update. There is an issue with the API. Please contact the Employee Management team in ATG.")

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
        values = get_values(ROSTER)
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
        try :
            update_person_response = service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet_Id, body=update_person_body).execute()
        except :
            print("The person was not added to the sheet due to an issue with the batch update. There is an issue with the API. Please contact the Employee Management team in ATG.")
    def __repr__(self):
        return self.get_first_name().capitalize() + self.get_middle_name().capitalize() + self.get_last_name().capitalize()

if __name__ == '__main__':
    main()
    