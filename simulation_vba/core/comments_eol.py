#!/usr/bin/env python
import logging
from pyparsing import Literal, SkipTo, Combine, Suppress, Optional, CaselessKeyword, OneOrMore
from vba_lines import line_terminator
from logger import log
if (log.getEffectiveLevel() == logging.DEBUG ):
    log.debug('importing comments_eol')
__version__ = '0.02'
single_quote = Literal("'")
comment_body = SkipTo(line_terminator)
comment_single_quote = Combine(single_quote + comment_body)
rem_statement = Suppress(Combine(CaselessKeyword('Rem') + comment_body))

EOL = Optional(comment_single_quote) + line_terminator
EOS = Suppress(Optional(";")) + OneOrMore(EOL | Literal(':'))
