"""
SimulationVBA: Class template for visitor classes for the visitor design pattern.

SimulationVBA is a specialized engine to parse, analyze and interpret Microsoft
VBA macros (Visual Basic for Applications), mainly for malware analysis.

Author: Philippe Lagadec - http://www.decalage.info
License: BSD, see source code or documentation

Project Repository:
https://github.com/decalage2/ViperMonkey
"""

# === LICENSE ==================================================================


class visitor(object):
    """
    The class template for visitor objects for the visitor design pattern.
    Visitors can be accepted by the accept method of VBA_Object objects.
    """

    def visit(self):
        raise NotImplementedError("Not implemented.")
