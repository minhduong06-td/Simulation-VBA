from visitor import visitor
import procedures


class function_defn_visitor(visitor):

    def __init__(self):
        self.funcs = set()
        self.func_objects = set()
        self.visited = set()
    
    def visit(self, item):
        if (item in self.visited):
            return False
        self.visited.add(item)
        if ((isinstance(item, procedures.Sub)) or
            (isinstance(item, procedures.Function)) or
            (isinstance(item, procedures.PropertyLet))):
            self.funcs.add(str(item.name))
            self.func_objects.add(item)
        return True
