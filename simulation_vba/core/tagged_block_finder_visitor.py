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
from statements import *

class tagged_block_finder_visitor(visitor):
    """
    Collect all the tagged block (labeled block) elements.
    """

    def __init__(self):
        self.blocks = {}
        self.visited = set()
    
    def visit(self, item):
        if (item in self.visited):
            return False
        self.visited.add(item)
        if (isinstance(item, TaggedBlock)):
            self.blocks[item.label] = item
        return True
