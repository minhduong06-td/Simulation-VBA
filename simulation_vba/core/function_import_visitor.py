from visitor import *
from procedures import *

class function_import_visitor(visitor):

    def __init__(self):
        self.names = set()
        self.aliases = set()
        self.funcs = {}
        self.visited = set()
        
    def visit(self, item):
        if (item in self.visited):
            return False
        self.visited.add(item)
        if (isinstance(item, External_Function)):
            self.funcs[str(item.name)] = str(item.alias_name)
            self.names.add(str(item.alias_name))
            self.aliases.add(str(item.name))
        return True
