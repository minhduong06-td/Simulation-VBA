"""
SimulationVBA: VBA Library

SimulationVBA is a specialized engine to parse, analyze and interpret Microsoft
VBA macros (Visual Basic for Applications), mainly for malware analysis.

Author: Philippe Lagadec - http://www.decalage.info
License: BSD, see source code or documentation

Project Repository:
https://github.com/decalage2/ViperMonkey
"""

# === LICENSE ==================================================================


__version__ = '0.02'

class from_unicode_str(str):
    """
    Marker class to mark strings created by StrConv() with the
    vbaFromUnicode option. VipeMonkey currently assumes that unless
    specifically noted, all strings are unicode. This class is used to
    mark strings that are pure ascii, not unicode.
    """
    pass
