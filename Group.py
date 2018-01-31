"""
During tree traversals of the group structure, Group objects will not be
created until needed. The subgroups field of Groups will store string names
of subgroups and when the actual Group object is needed, we call get_group on
that string name.

After making modifications to a group, use the save() function to commit the changes
to the spreadsheet.
"""

class Group:

    def __init__(self, name='', subgroups=[], parent=None):
        self.name = name
        self.subgroups = subgroups
        # Stored as a dictionary of SIDs to roles, as Person objects here
        # will take up a considerable amount of space. This dictionary
        # should only be populated at leaf nodes.
        self.people = {}
        self.parent = parent
        self.googlegroup = 'example@googlegroups.com'

    # Returns whether this group is a leaf or not. The group is a leaf only if there are no subgroups.
    def isLeaf(self):
        return not bool(subgroups)

    # Takes in a string name of the subgroup we are adding as a child of this group.
    def add_subgroup(self, subgroup):
        if isinstance(subgroup, Group)
            self.subgroups.append(subgroup)

    # Adds a person to the group if they are not already in it. People can only be added to groups
    # that are leaves in the ULAB organization. Traverse the path back to the root and add the person
    # to the group's respective column on mainroster.
    def add_person_to_group(self, person, role):
        if not person or not isinstance(person, Person):
            print("Please provide a proper person.")
            return
        if not self.isLeaf():
            print("Please specify a more specific subgroup to add this person to.")
            return

        if person.SID in self.people:
            print("This person already exists in the group.")
        else :
            self.people[person.SID] = role
            group = self
            while group.parent != None:
                if group.name not in person.groups:
                    person.groups.add(group.name)
                group = group.parent

    # Removes a person from this group if they are in the group. If this group is a leaf, the person is only removed from this leaf group.
    # If this group is an internal node, we remove the person from all the subgroups as well. Returns true if the remove is successful and
    # returns false otherwise.
    def remove_person_from_group(self, person):
        if not person or not isinstance(person, Person):
            print("Please provide a proper person.")
            return

        if self.isLeaf():
            if person.SID not in self.people:
                print("This person does not exist in the group.")
            else
                del self.people[person.SID]
        else:
            for subgroup_name in self.subgroups:
                subgroup = Group.get_group(subgroup_name)
                if isinstance(subgroup, Group):
                    if group.name in person.groups :
                        person.groups.remove(group.name)
                    
    def
