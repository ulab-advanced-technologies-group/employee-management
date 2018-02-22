from __future__ import print_function
import httplib2
import os

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/drive-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/drive'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Drive API Python Quickstart'


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
                                   'drive-python-quickstart.json')

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

credentials = get_credentials()
http = credentials.authorize(httplib2.Http())
service = discovery.build('drive', 'v3', http=http)

# We want to only create a new directory when adding a new group?
# We want to test creating a directory for entirety of ULAB?
# Folder directory id for the EM_DEMO directory. Will change later when ported to
# the real ULAB folder.
parent_directory_id = '1eKHhSEiIAGJn3qEiItvyop6MulVsG-4L'

# If true, group already in directory and subdirectories, false otherwise
def check_group(name, parentId=parent_directory_id):
    # if parentId is parent_directory_id:
    #     query = """trashed = false"""
    # else:
    query = """trashed = false and '""" + parentId + "'" + """ in parents"""
    try :
        files = service.files().list(fields="files(name, id)", q=query).execute()
        groups = files.get('files')
    except :
        print("The files that have corresponding group names and IDs could not be accessed. There is an issue with the API. Please contact the Employee Management team in ATG.")
    for group in groups:
        print(group)
        if name == group.get('name'):
            return True
        elif check_group(name, group.get('id')):
            return True
    return False

def create_new_directory(name, parentId=parent_directory_id):
    # Checks if group is subgroup in parent by one level
    print(name, parentId)
    if not check_group(name, parentId):
        newGroup_metadata = {
            'name': name,
            'mimeType': 'application/vnd.google-apps.folder',
            'parents': [parentId]
        }
        try :
            newGroup = service.files().create(body=newGroup_metadata, fields='id').execute()
            newGroupId = newGroup['id']
        except :
            print("A group directory could not be created. There is an issue with the API. Please contact the Employee Management team in ATG.")

        ### adds edit privileges to file/folder
        #add_permissions('khang.nguyencampbell@gmail.com', newGroupId)
        return
    print('Group already in directory')


def delete_directory(group_name, parent_id=parent_directory_id):
    if group_name != 'ULAB' or group_name != 'ulab':
        group_id = get_group_id(group_name, parent_id)
        try :
            del_group = service.files().delete(fileId=group_id).execute()
        except :
            print("The group directory could not be deleted. There is an issue with the API. Please contact the Employee Management team in ATG.")
        print(group_name, 'deleted')
    else:
        print('Cannot delete ULAB')

def get_group_id(group_name, parentId=parent_directory_id):
    # if parentId is None:
    #     query = """trashed = false and name='""" + group_name + """'"""
    # else:
    query = """trashed = false and name='""" + group_name + """' and '""" + parentId + "'" + """ in parents"""
    # print(query)
    try :
        results = service.files().list(pageSize=10, fields="nextPageToken, files(id, name)", q=query).execute()
    except :
        print("The files that have corresponding group names and IDs could not be accessed. There is an issue with the API. Please contact the Employee Management team in ATG.")
    # print('results',results)
    items = results.get('files', [])
    if not items:
        print('No groups found')
    else:
        for item in items:
            # print('folder', item)
            if item['name'] == group_name:
                return item['id']
        print('Group not found in folder')

def get_permission_id(email_address, group_id):
    try :
        response = service.permissions().list(fileId=group_id, fields="permissions(id, emailAddress)").execute()
    except :
        print("The permissions could not be accessed for this group. There is an issue with the API. Please contact the Employee Management team in ATG.")
        perms = response.get('permissions')
        for perm in perms:
            if perm['emailAddress'] == email_address:
                return perm['id']
        return None

def add_permissions(email_address, group_name, parent_directory_id=parent_directory_id):
    group_id = get_group_id(group_name, parent_directory_id)
    permissions = {
        'role': 'writer',
        'type': 'user',
        'emailAddress': email_address
    }
    try :
        add_pm = service.permissions().create(fileId=group_id,
                    body=permissions, sendNotificationEmail=False).execute()
    except :
        print("The permissions were unable to be added. There is an issue with the API.Please contact the Employee Management team in ATG.")

def remove_permissions(email_address, group_name, parent_directory_id=parent_directory_id):
    group_id = get_group_id(group_name, parent_directory_id)
    perm_id = get_permission_id(email_address, group_id)
    if perm_id != None:
        try :
            del_pm = service.permissions().delete(fileId=group_id, permissionId=perm_id).execute()
        except :
            print("The permissions were unable to be removed. There is an issue with the API. Please contact the Employee Management team in ATG.")
        print(email_address, "removed from " + group_name + " folder")
        return True
    print(email_address, "not found in given group")
    return False
