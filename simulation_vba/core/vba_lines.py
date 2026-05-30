#!/usr/bin/env python





__version__ = '0.02'



from logger import log

from pyparsing import *




ParserElement.setDefaultWhitespaceChars(' \t\x19')



non_line_termination_character = CharsNotIn('\x0D\x0A', exact=1)
line_terminator = Literal('\x0D\x0A') | Literal('\x0D') | Literal('\x0A')
non_terminated_line = Optional(CharsNotIn('\x0D\x0A'))
source_line = Optional(CharsNotIn('\x0D\x0A')) + line_terminator
module_body_physical_structure = ZeroOrMore(source_line) + Optional(non_terminated_line)


whitespaces = Word(' \t\x19').leaveWhitespace()
line_continuation = (whitespaces + '_' + Optional(whitespaces) + line_terminator).leaveWhitespace()
line_continuation.setParseAction(replaceWith(' '))
extended_line = Combine(ZeroOrMore(line_continuation | non_line_termination_character) + line_terminator)
module_body_logical_structure = ZeroOrMore(extended_line)
logical_line = LineStart() + ZeroOrMore(extended_line.leaveWhitespace()) + line_terminator
module_body_lines = Combine(ZeroOrMore(logical_line))


def vba_collapse_long_lines(vba_code):
    """
    Parse a VBA module code to detect continuation line characters (underscore) and
    collapse split lines. Continuation line characters are replaced by spaces.

    :param vba_code: str, VBA module code
    :return: str, VBA module code with long lines collapsed
    """
    if (vba_code is None):
        return ""
    if vba_code[-1] != '\n':
        vba_code += '\n'
    vba_code = vba_code.replace(' _\r\n', ' ')
    vba_code = vba_code.replace(' _\r', ' ')
    vba_code = vba_code.replace(' _\n', ' ')
    return vba_code
