#!/usr/bin/env python



__version__ = '0.02'


import logging
import re

from pyparsing import *

from logger import log
from vba_object import VBA_Object


boolean_literal = Regex(re.compile('(True|False)', re.IGNORECASE))
boolean_literal.setParseAction(lambda t: bool(t[0].lower() == 'true'))



decimal_literal = Regex(re.compile('(?P<value>[+\-]?\d+)[%&^]?[!#@]?'))
decimal_literal.setParseAction(lambda t: int(t.value))

octal_literal = Regex(re.compile('&o?(?P<value>[0-7]+)[%&^]?', re.IGNORECASE))
octal_literal.setParseAction(lambda t: int(t.value, base=8))

hex_literal = Regex(re.compile('&h(?P<value>[0-9a-f]+)[%&^]?', re.IGNORECASE))
hex_literal.setParseAction(lambda t: int(t.value, base=16))

integer = decimal_literal | octal_literal | hex_literal



float_literal = Regex(re.compile('(?P<value>[+\-]?\d+\.\d*([eE][+\-]?\d+)?)[!#@]?')) | \
                Regex(re.compile('(?P<value>[+\-]?\d+[eE][+\-]?\d+)[!#@]?'))
float_literal.setParseAction(lambda t: float(t.value))


class String(VBA_Object):

    def __init__(self, original_str, location, tokens):
        super(String, self).__init__(original_str, location, tokens)
        self.value = tokens[0]        
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed "%r" as String' % self)

    def __repr__(self):
        return '"' + str(self.value) + '"'

    def eval(self, context, params=None):
        r = self.value
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("String.eval: return " + r)
        return r

    def to_python(self, context, params=None, indent=0):
        r = str(self.value).\
            replace("\\", "\\\\").\
            replace('"', '\\"').\
            replace("\n", "\\n").\
            replace("\t", "\\t").\
            replace("\r", "\\r")
        for i in range(0, 9):
            repl = hex(i).replace("0x", "")
            if (len(repl) == 1):
                repl = "0" + repl
            repl = "\\x" + repl
            r = r.replace(chr(i), repl)
        for i in range(11, 13):
            repl = hex(i).replace("0x", "")
            if (len(repl) == 1):
                repl = "0" + repl
            repl = "\\x" + repl
            r = r.replace(chr(i), repl)
        for i in range(14, 32):
            repl = hex(i).replace("0x", "")
            if (len(repl) == 1):
                repl = "0" + repl
            repl = "\\x" + repl
            r = r.replace(chr(i), repl)
        for i in range(127, 255):
            repl = hex(i).replace("0x", "")
            if (len(repl) == 1):
                repl = "0" + repl
            repl = "\\x" + repl
            r = r.replace(chr(i), repl)
        return '"' + r + '"'

quoted_string = QuotedString('"', escQuote='""', convertWhitespaceEscapes=False)('value')
quoted_string.setParseAction(String)

quoted_string_keep_quotes = QuotedString('"', escQuote='""', unquoteResults=False, convertWhitespaceEscapes=False)
quoted_string_keep_quotes.setParseAction(lambda t: str(t[0]))



date_string = QuotedString('#')
date_string.setParseAction(lambda t: str(t[0]))



literal = boolean_literal | integer | quoted_string | date_string | float_literal
literal.setParseAction(lambda t: t[0])

