#!/usr/bin/env python



__version__ = '0.02'


from curses_ascii import isprint
import logging
from pyparsing import *

from vba_object import *
from literals import *
import vb_str

from logger import log



expression = Forward()


class Chr(VBA_Object):
    """
    6.1.2.11.1.4 VBA Chr function
    """

    def __init__(self, original_str, location, tokens):
        super(Chr, self).__init__(original_str, location, tokens)
        self.arg = tokens[0]
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed %r as %s' % (self, self.__class__.__name__))

    def to_python(self, context, params=None, indent=0):
        arg_str = to_python(self.arg, context)
        r = "core.vba_library.run_function(\"_Chr\", vm_context, [" + arg_str + "])"
        return r

    def return_type(self):
        return "STRING"
    
    def eval(self, context, params=None):

        import vba_library
        chr_handler = vba_library._Chr()
        param = eval_arg(self.arg, context)
        return chr_handler.eval(context, [param])

    def __repr__(self):
        return 'Chr(%s)' % repr(self.arg)

chr_ = (
    Suppress(Regex(re.compile('Chr[BW]?\$?', re.IGNORECASE)))
    + Suppress('(')
    + expression
    + Suppress(')')
)
chr_.setParseAction(Chr)


class Asc(VBA_Object):
    """
    VBA Asc function
    """

    def __init__(self, original_str, location, tokens):
        super(Asc, self).__init__(original_str, location, tokens)

        self.arg = None
        if (len(tokens) > 0):
            self.arg = tokens[0]

    def to_python(self, context, params=None, indent=0):
        return "ord(" + to_python(self.arg, context) + ")"

    def return_type(self):
        return "INTEGER"
    
    def eval(self, context, params=None):

        if (self.arg is None):
            try:
                return context.get("asc")
            except KeyError:
                return "NULL"
        
        c = eval_arg(self.arg, context)

        c_str = None
        try:
            c_str = str(c).strip()
        except UnicodeEncodeError:
            c_str = filter(isprint, c).strip()
        if (c_str == "**MATCH ANY**"):
            return c

        if (c == "NULL"):
            return 0
        
        if (isinstance(c, int)):
            r = c
        else:


            if (c_str == "**MATCH ANY**"):
                r = "**MATCH ANY**"

            else:
                r = vb_str.get_ms_ascii_value(c_str)

        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Asc: return %r" % r)
        return r

    def __repr__(self):
        return 'Asc(%s)' % repr(self.arg)


asc = Suppress((CaselessKeyword('Asc') | CaselessKeyword('AscW')))  + Optional(Suppress('(') + expression + Suppress(')'))
asc.setParseAction(Asc)


class StrReverse(VBA_Object):
    """
    VBA StrReverse function
    """

    def __init__(self, original_str, location, tokens):
        super(StrReverse, self).__init__(original_str, location, tokens)
        self.arg = tokens[0]

    def return_type(self):
        return "STRING"
        
    def eval(self, context, params=None):
        return eval_arg(self.arg, context)[::-1]

    def __repr__(self):
        return 'StrReverse(%s)' % repr(self.arg)

strReverse = Suppress(CaselessLiteral('StrReverse') + Literal('(')) + expression + Suppress(Literal(')'))
strReverse.setParseAction(StrReverse)


class Environ(VBA_Object):
    """
    VBA Environ function
    """

    def __init__(self, original_str, location, tokens):
        super(Environ, self).__init__(original_str, location, tokens)
        self.arg = tokens.arg

    def return_type(self):
        return "STRING"        

    def eval(self, context, params=None):
        arg = eval_arg(self.arg, context=context)
        value = '%%%s%%' % arg
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('evaluating Environ(%s) => %r' % (arg, value))
        return value

    def __repr__(self):
        return 'Environ(%s)' % repr(self.arg)

environ = Suppress(CaselessKeyword('Environ') + '(') + expression('arg') + Suppress(')')
environ.setParseAction(Environ)
