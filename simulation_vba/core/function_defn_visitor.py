




import os, sys

_thismodule_dir = os.path.normpath(os.path.abspath(os.path.dirname(__file__)))
_parent_dir = os.path.normpath(os.path.join(_thismodule_dir, '../..'))
if _parent_dir not in sys.path:
    sys.path.insert(0, _parent_dir)

from simulation_vba.core import *


class function_defn_visitor(visitor):
    """
    Collect the names of all locally declared functions.
    """

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
