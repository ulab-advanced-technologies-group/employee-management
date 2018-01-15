class Person:
    def __init__(self, FirstName, MiddleName, LastName, SID, Email, PhoneNumber) :
        self.FirstName = FirstName
        self.LastName = LastName
        self.Roles = []
        self.SID = SID
        self.MiddleName = MiddleName
        self.Email = Email
        self.PhoneNumber = PhoneNumber
        self.DietaryPreferences = None
        self.schedule = None
        self.Groups = []

<<<<<<< HEAD
    def __init__(self, FirstName, LastName, SID) :
        self.FirstName = FirstName
        self.LastName = LastName
        self.SID = SID

    def addGroup(Group) :
        self.Groups.append(Group)

=======
>>>>>>> f341050710447d401d6ac9a3908f344a0efb3fc7
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
