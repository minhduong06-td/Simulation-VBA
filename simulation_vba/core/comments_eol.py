#!/usr/bin/env python

"""@package comments_eol
Parsing of VB comments and end of line markers.
"""

import logging
from pyparsing import Literal, SkipTo, Combine, Suppress, Optional, CaselessKeyword, OneOrMore
from vba_lines import line_terminator
from logger import log
if (log.getEffectiveLevel() == logging.DEBUG ):
    log.debug('importing comments_eol')

"""
SimulationVBA is a specialized engine to parse, analyze and interpret Microsoft
VBA macros (Visual Basic for Applications), mainly for malware analysis.

Author: Philippe Lagadec - http://www.decalage.info
License: BSD, see source code or documentation

Project Repository:
https://github.com/decalage2/ViperMonkey
"""

# === LICENSE ==================================================================




__version__ = '0.02'



single_quote = Literal("'")
comment_body = SkipTo(line_terminator)

comment_single_quote = Combine(single_quote + comment_body)

rem_statement = Suppress(Combine(CaselessKeyword('Rem') + comment_body))



EOL = Optional(comment_single_quote) + line_terminator

EOS = Suppress(Optional(";")) + OneOrMore(EOL | Literal(':'))
