class Group :
    def __init__(self, name) :
        self.name = name
        self.subgroups = []

    def addsubGroup(self, obj) :
        self.subgroups.append(obj)
