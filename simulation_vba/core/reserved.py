#!/usr/bin/env python
__version__ = '0.02'
from pyparsing import *
from logger import log
from identifiers import *
def caselessKeywordsList(keywords):
    p = CaselessKeyword(keywords[0])
    for kw in keywords[1:]:
        p |= CaselessKeyword(kw)
    return p

statement_keyword = caselessKeywordsList(
    ("Call", "Const", "Declare", "DefBool", "DefByte",
     "DefCur", "DefDate", "DefDbl", "DefInt", "DefLng", "DefLngLng", "DefLngPtr", "DefObj",
     "DefSng", "DefStr", "DefVar", "Dim", "Do", "Else", "ElseIf", "End If",
     "Enum", "Event", "Exit", "Friend", "Function",
     "GoSub", "GoTo", "If", "Implements", "Let", "Loop", "LSet", "Next",
     "On", "Open", "Option", "Private", "Public", "RaiseEvent", "ReDim",
     "Resume", "RSet", "Select", "Set", "Static", "Stop", "Sub",
     "Unlock", "Wend", "While", "With"))

rem_keyword = CaselessKeyword("Rem")

marker_keyword = caselessKeywordsList(
    ("As", "ByRef", "ByVal ", "Case", "For", "Each", "Else", "In", "New",
     "Shared", "Until", "WithEvents", "Optional", "ParamArray", "Preserve",
     "Tab", "Then"))

operator_identifier = caselessKeywordsList(
    ("AddressOf", "And", "Eqv", "Imp", "Is", "Like", "New", "Mod",
     "Not", "Or", "TypeOf", "Xor"))

reserved_name = caselessKeywordsList((
    "CBool", "CByte", "CCur", "CDate",
    "CLng", "CLngLng", "CLngPtr", "CSng", "CStr", "CVar", "CVErr",
    "DoEvents", "Fix", "Int", "Len", "LenB", "PSet", "Sgn", "String"))

special_form = caselessKeywordsList((
    "Array", "Circle", "InputB", "LBound", "UBound"))


simple_type_identifier = Word(initChars=alphas, bodyChars=alphanums + '_')
reserved_complex_type_identifier = Group(simple_type_identifier + ZeroOrMore("." + simple_type_identifier))

reserved_atomic_type_identifier = caselessKeywordsList((
    "Boolean", "Byte", "Currency", "Date", "Double", "Integer",
    "Long", "LongLong", "LongPtr", "Single", "String", "Variant"))

reserved_type_identifier = reserved_atomic_type_identifier | reserved_complex_type_identifier

boolean_literal_identifier = CaselessKeyword("true") | CaselessKeyword("false")

object_literal_identifier = CaselessKeyword("nothing")

variant_literal_identifier = CaselessKeyword("empty") | CaselessKeyword("null")

literal_identifier = boolean_literal_identifier | object_literal_identifier

reserved_for_implementation_use = caselessKeywordsList((
    "LINEINPUT", "VB_Base", "VB_Control",
    "VB_Creatable", "VB_Customizable", "VB_Description", "VB_Exposed", "VB_Ext_KEY ",
    "VB_GlobalNameSpace", "VB_HelpID", "VB_Invoke_Func", "VB_Invoke_Property ",
    "VB_Invoke_PropertyPut", "VB_Invoke_PropertyPutRefVB_MemberFlags", "VB_Name",
    "VB_PredeclaredId", "VB_ProcData", "VB_TemplateDerived", "VB_UserMemId",
    "VB_VarDescription", "VB_VarHelpID", "VB_VarMemberFlags", "VB_VarProcData ",
    "VB_VarUserMemId"))

future_reserved = caselessKeywordsList(("CDecl", "Decimal", "DefDec"))

reserved_identifier = statement_keyword | marker_keyword | operator_identifier \
                      | special_form | reserved_name | literal_identifier | rem_keyword \
                      | reserved_for_implementation_use | future_reserved

