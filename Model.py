from typing import Type
class Variable :
    def __init__(self, code, Value=None, lbl=None):
        self.code = code
        self.value = Value
        self.label = lbl

class VariableValue :
    def __init__(self, code, Value=None, Type=None, isValueUsed=None, lbl=None):
        self.code = code
        self.value = Value
        self.type = Type
        self.isValueUsed = isValueUsed
        self.label = lbl

class TaskType :
    def __init__(self, code, name=None, lbl=None):
        self.code = code
        self.name = name
        self.label = lbl

class Technique :
    def __init__(self, code, value=None, lbl=None):
        self.code = code
        self.value = value
        self.label = lbl

class Level :
    def __init__(self, code, Value=None, lbl=None):
        self.code = code
        self.value = Value
        self.label = lbl

class Student :
    def __init__(self, code, Value=None, lbl=None):
        self.code = code
        self.value = Value
        self.label = lbl

class Group :
    def __init__(self, code, Value=None, lbl=None):
        self.code = code
        self.value = Value
        self.label = lbl

class LocalInsInfo :
    def __init__(self, code, Value=None, lbl=None):
        self.code = code
        self.label = lbl

class Relationship :
    def __init__(self,typeLeft, left, right, typeRight, nameRelationship, lbl=None):
        self.typeL = typeLeft
        self.left = left
        self.right = right
        self.typeR = typeRight
        self.nameRelationship = nameRelationship
        self.lable = lbl

