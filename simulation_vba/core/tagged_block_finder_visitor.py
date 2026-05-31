from visitor import *
from statements import *

class tagged_block_finder_visitor(visitor):

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
