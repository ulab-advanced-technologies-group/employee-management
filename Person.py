"""
After making modifications to a person, use the save() function to commit the changes
to the spreadsheet.
"""

class Person:
    def __init__(self, first_name='', middle_name='', last_name='', SID='', email='', phone='', majors=[], year='',
                        back_id='', dietary_preferences='', expected_graduation='', classes=[],
                        goals='', accounts=[], accesses=[]):
        self.first_name = first_name
        self.middle_name = middle_name
        self.last_name = last_name
        self.SID = SID
        self.email = email
        self.phone = phone
        self.majors = majors
        self.year = year
        self.back_id = back_id
        self.dietary_preferences = dietary_preferences
        self.expected_graduation = expected_graduation
        self.classes = classes
        self.goals = goals
        self.accounts = accounts
        self.accesses = accesses
        # Stores the string names of the groups that this person is a part of. Using the get_group() in employee-management.py
        # to retrieve the Group object for this group. Since order does not matter, we will use a set.
        self.groups = set()

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

    def adjustEmail(newEmail) :
        self.Email = newEmail

    def adjustPhoneNumber(newPhoneNumber) :
        self.PhoneNumber = newPhoneNumber

    def dietaryPreferences(DietaryPreferences):
        self.DietaryPreferences = DietaryPreferences

    def linkSchedule(schedule):
        self.schedule = schedule
