import sys
from visitor import *
import pyparsing

class lhs_var_visitor(visitor):

    def __init__(self):
        self.variables = set()
        self.visited = set()
    
    def visit(self, item):
        from statements import Let_Statement

        if (str(item) in self.visited):
            return False
        self.visited.add(str(item))
        if ("Let_Statement" in str(type(item))):
            if (isinstance(item.name, str)):
                self.variables.add(item.name)
            elif (isinstance(item.name, pyparsing.ParseResults) and
                  (item.name[0].lower().replace("$", "").replace("#", "").replace("%", "") == "mid")):
                self.variables.add(str(item.name[1]))

        return True
