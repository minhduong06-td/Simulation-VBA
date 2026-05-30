"""
SimulationVBA: Visitor for collecting Let statements.

SimulationVBA is a specialized engine to parse, analyze and interpret Microsoft
VBA macros (Visual Basic for Applications), mainly for malware analysis.

Author: Philippe Lagadec - http://www.decalage.info
License: BSD, see source code or documentation

Project Repository:
https://github.com/decalage2/ViperMonkey
"""

# === LICENSE ==================================================================


from visitor import *

class let_statement_visitor(visitor):
    """
    Get all Let statements.
    """

    def __init__(self, var_name=None):
        self.let_statements = set()
        self.visited = set()
        self.var_name = var_name
    
    def visit(self, item):
        from statements import Let_Statement
        if (item in self.visited):
            return False
        self.visited.add(item)        
        if ((isinstance(item, Let_Statement)) and
            ((self.var_name is None) or (item.name == self.var_name))):
            self.let_statements.add(item)
        return True
