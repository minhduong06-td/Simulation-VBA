


from visitor import *
from statements import *

class var_defn_visitor(visitor):
    """
    Collect the names of all declared variables.
    """

    def __init__(self):
        self.variables = set()
        self.visited = set()
    
    def visit(self, item):
        if (item in self.visited):
            return False
        self.visited.add(item)        
        if (isinstance(item, Dim_Statement)):
            for name, _, _, _ in item.variables:
                self.variables.add(str(name))
        if (isinstance(item, Let_Statement)):
            self.variables.add(str(item.name))
        return True
