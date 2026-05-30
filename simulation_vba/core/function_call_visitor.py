"""
SimulationVBA: Visitor for collecting the names of all called functions

SimulationVBA is a specialized engine to parse, analyze and interpret Microsoft
VBA macros (Visual Basic for Applications), mainly for malware analysis.

Author: Philippe Lagadec - http://www.decalage.info
License: BSD, see source code or documentation

Project Repository:
https://github.com/decalage2/ViperMonkey
"""

# === LICENSE ==================================================================


from visitor import *

class function_call_visitor(visitor):
    """
    Collect the names of all called functions.
    """

    def __init__(self):
        self.called_funcs = set()
        self.visited = set()
    
    def visit(self, item):

        import statements
        import expressions
        import lib_functions

        if (item in self.visited):
            return False
        self.visited.add(item)
        if (isinstance(item, statements.Call_Statement)):
            if (not isinstance(item.name, expressions.MemberAccessExpression)):
                self.called_funcs.add(str(item.name))
        if (isinstance(item, expressions.Function_Call)):
            self.called_funcs.add(str(item.name))
        if (isinstance(item, statements.File_Open)):
            self.called_funcs.add("Open")
        if (isinstance(item, statements.Print_Statement)):
            self.called_funcs.add("Print")
        if (isinstance(item, lib_functions.Chr)):
            self.called_funcs.add("Chr")
        if (isinstance(item, lib_functions.Asc)):
            self.called_funcs.add("Asc")
        if (isinstance(item, lib_functions.StrReverse)):
            self.called_funcs.add("StrReverse")
        if (isinstance(item, lib_functions.Environ)):
            self.called_funcs.add("Environ")
        return True        
