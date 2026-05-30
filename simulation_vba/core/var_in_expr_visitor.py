"""
SimulationVBA: Visitor for collecting the names declared variables.

SimulationVBA is a specialized engine to parse, analyze and interpret Microsoft
VBA macros (Visual Basic for Applications), mainly for malware analysis.

Author: Philippe Lagadec - http://www.decalage.info
License: BSD, see source code or documentation

Project Repository:
https://github.com/decalage2/ViperMonkey
"""

# === LICENSE ==================================================================


from visitor import *
from statements import *

class var_in_expr_visitor(visitor):
    """
    Get the names of all variables that appear in an expression.
    """

    def __init__(self, context=None):
        self.variables = set()
        self.visited = set()
        self.context = context
    
    def visit(self, item):
        from expressions import SimpleNameExpression
        from expressions import Function_Call
        from expressions import MemberAccessExpression
        from vba_object import VbaLibraryFunc
        from vba_object import VBA_Object

        if (item in self.visited):
            return False
        self.visited.add(item)

        if (isinstance(item, SimpleNameExpression)):
            self.variables.add(str(item.name))

        if (("Function_Call" in str(type(item))) and (self.context is not None)):

            if (hasattr(item, "name") and (self.context.contains(item.name))):
                ref = self.context.get(item.name)
                if (isinstance(ref, list) or isinstance(ref, str)):
                    self.variables.add(str(item.name))

        if (isinstance(item, MemberAccessExpression)):
            rhs = item.rhs
            if (isinstance(rhs, list)):
                rhs = rhs[-1]
            if (isinstance(rhs, SimpleNameExpression)):
                self.variables.add(str(item))
                    
        return True
