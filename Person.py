class Person:
    def __init__(self, FirstName, LastName, SID) :
        self.FirstName = FirstName
        self.LastName = LastName
        self.Roles = []
        self.SID = SID
        self.Groups = []

    def addGroup(Group) :
        self.Groups.append(Group)

    def addRole(Role) :
        self.Roles.append(Role)
