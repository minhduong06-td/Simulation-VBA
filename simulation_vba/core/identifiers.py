#!/usr/bin/env python
import re
__version__ = '0.02'
from pyparsing import *
from reserved import *
from logger import log

reserved_keywords = CaselessKeyword("ChrB") | \
                    CaselessKeyword("ChrB") | \
                    CaselessKeyword("ChrW") | \
                    CaselessKeyword("Asc") | \
                    CaselessKeyword("Case") | \
                    CaselessKeyword("On") | \
                    CaselessKeyword("Sub") | \
                    CaselessKeyword("If") | \
                    CaselessKeyword("Then") | \
                    CaselessKeyword("For") | \
                    CaselessKeyword("Next") | \
                    CaselessKeyword("Public") | \
                    CaselessKeyword("Private") | \
                    CaselessKeyword("Declare") | \
                    CaselessKeyword("Function") | \
                    CaselessKeyword("To")

strict_reserved_keywords = reserved_keywords | \
                           Regex(re.compile('Open', re.IGNORECASE)) | \
                           Regex(re.compile('While', re.IGNORECASE))



general_identifier = Word(initChars=alphas + alphas8bit + '_' + '?', bodyChars=alphanums + '_' + '?' + alphas8bit) + \
                     Suppress(Optional("^")) + Suppress(Optional("%")) + Suppress(Optional("!..."))

lex_identifier = general_identifier | Regex(r"%\w+%") | "..."


identifier = NotAny(reserved_identifier) + lex_identifier

identifier.setParseAction(lambda t: t[0])



foreign_name = Literal('[') + CharsNotIn('\x0D\x0A') + Literal(']')


builtin_type = reserved_type_identifier | (Suppress("[") + reserved_type_identifier + Suppress("]")) \
               | CaselessKeyword("object") | CaselessLiteral("[object]")

type_suffix = Word(r"%&^!#@$", exact=1) + \
              NotAny(Optional(Regex(r" +")) + ((NotAny(reserved_keywords) + Word(alphanums)) | '"'))
typed_name = Combine(identifier + type_suffix)

untyped_name = identifier
entity_name = typed_name | untyped_name
unrestricted_name = entity_name | reserved_identifier


base_attrib = Combine(
    NotAny(reserved_keywords)
    + (Combine(Literal('.') + lex_identifier) | Combine(entity_name + Optional(Literal('.') + lex_identifier)))
    + Optional(CaselessLiteral('$'))
    + Optional(CaselessLiteral('#'))
    + Optional(CaselessLiteral('%'))
)

TODO_identifier_or_object_attrib = base_attrib ^ Suppress(Literal("{")) + base_attrib + Suppress(Literal("}"))

base_attrib_loose = Combine(
    Combine(Literal('.') + lex_identifier)
    | Combine(entity_name + Optional(Literal('.') + lex_identifier))
    + Optional(CaselessLiteral('$'))
    + Optional(CaselessLiteral('#'))
    + Optional(CaselessLiteral('%'))
    | Combine(entity_name + Literal('.') + lex_identifier + Literal('.') + lex_identifier)
)

TODO_identifier_or_object_attrib_loose = base_attrib_loose ^ Suppress(Literal("{")) + base_attrib_loose + Suppress(Literal("}"))

enum_val_id = Regex(re.compile(r"\[[^\]]+\]"))
