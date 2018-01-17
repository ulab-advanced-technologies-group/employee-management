class Group :
    def __init__(self, name) :
        self.name = name
        self.subgroups = []
        self.people = []

    def addsubGroup(self, obj) :
        self.subgroups.append(obj)

    def addPerson(Person) :
        if Person in self.people :
            print("This person already exists in the group.")
        else :
            self.people.append(Person)

    def removePerson(Person) :
        if Person not in self.people :
            print("This person does not exist in the group.")
        else
            self.people.remove(Person)
