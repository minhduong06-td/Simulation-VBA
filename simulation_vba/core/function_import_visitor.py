"""
SimulationVBA: Visitor for collecting the names of locally defined functions

SimulationVBA is a specialized engine to parse, analyze and interpret Microsoft
VBA macros (Visual Basic for Applications), mainly for malware analysis.

Author: Philippe Lagadec - http://www.decalage.info
License: BSD, see source code or documentation

Project Repository:
https://github.com/decalage2/ViperMonkey
"""

# === LICENSE ==================================================================


from visitor import *
from procedures import *

class function_import_visitor(visitor):
    """
    Collect the names and aliases of all functions imported from DLLs.
    """

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
