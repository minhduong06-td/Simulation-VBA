#!/usr/bin/env python
try:
    unicode
except NameError:
    unicode = str
try:
    basestring
except NameError:
    basestring = (str, bytes)
try:
    long
except NameError:
    long = int

__version__ = '0.08'
import logging
from comments_eol import *
from expressions import *
from vba_context import *
from reserved import *
from from_unicode_str import *
from vba_object import to_python
from vba_object import _eval_python
from vba_object import _boilerplate_to_python
from vba_object import _updated_vars_to_python
from vba_object import _loop_vars_to_python
from vba_object import _get_var_vals
import procedures
from let_statement_visitor import *
from var_in_expr_visitor import *
from lhs_var_visitor import *
from function_call_visitor import *
import vb_str
import loop_transform
from utils import safe_print
import utils
import traceback
import string
import pyparsing
from logger import log
import sys
import re
import base64
from curses_ascii import isprint
import hashlib

def _vba_truthy(val):
    orig_repr = repr(val)
    if (len(orig_repr) > 200):
        orig_repr = orig_repr[:200] + "..."
    val_type = type(val).__name__

    if val is None:
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("If guard truthy: repr=%s type=%s -> False (None)", orig_repr, val_type)
        return False

    if isinstance(val, bool):
        r = val
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("If guard truthy: repr=%s type=%s -> %s (bool)", orig_repr, val_type, r)
        return r

    if isinstance(val, (int, float)):
        r = (val != 0)
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("If guard truthy: repr=%s type=%s -> %s (numeric)", orig_repr, val_type, r)
        return r

    if isinstance(val, str):
        if (val == ""):
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("If guard truthy: repr=%s type=%s -> False (empty string)", orig_repr, val_type)
            return False
        try:
            num = float(val)
            r = (num != 0.0)
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("If guard truthy: repr=%s type=%s -> %s (numeric string)", orig_repr, val_type, r)
            return r
        except (ValueError, TypeError):
            pass
        if (val.lower() in ("false", "vbfalse")):
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("If guard truthy: repr=%s type=%s -> False (falsy token)", orig_repr, val_type)
            return False
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("If guard truthy: repr=%s type=%s -> True (non-empty string)", orig_repr, val_type)
        return True

    if isinstance(val, (list, tuple, set, dict)):
        r = len(val) > 0
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("If guard truthy: repr=%s type=%s -> %s (collection)", orig_repr, val_type, r)
        return r

    r = bool(val)
    if (log.getEffectiveLevel() == logging.DEBUG):
        log.debug("If guard truthy: repr=%s type=%s -> %s (fallback)", orig_repr, val_type, r)
    return r


def is_simple_statement(s):

    return (isinstance(s, Dim_Statement) or
            isinstance(s, Option_Statement) or
            isinstance(s, Prop_Assign_Statement) or
            isinstance(s, Let_Statement) or
            isinstance(s, LSet_Statement) or
            isinstance(s, Exit_For_Statement) or
            isinstance(s, Exit_While_Statement) or
            isinstance(s, Exit_Function_Statement) or
            isinstance(s, Redim_Statement) or
            isinstance(s, Goto_Statement) or
            isinstance(s, On_Error_Statement) or
            isinstance(s, File_Open) or
            isinstance(s, Print_Statement))
    


class UnknownStatement(VBA_Object):

    def __init__(self, original_str, location, tokens):
        super(UnknownStatement, self).__init__(original_str, location, tokens)
        self.text = tokens.text
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed %r' % self)

    def __repr__(self):
        return 'Unknown statement: %s' % repr(self.text)

    def eval(self, context, params=None):
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug(self)

known_keywords_statement_start = (Optional(CaselessKeyword('Public') | CaselessKeyword('Private') | CaselessKeyword('End')) + \
                                  (CaselessKeyword('Sub') | CaselessKeyword('Function'))) | \
                                  CaselessKeyword('Set') | CaselessKeyword('For') | CaselessKeyword('Next') | \
                                  CaselessKeyword('If') | CaselessKeyword('Then') | CaselessKeyword('Else') | \
                                  CaselessKeyword('ElseIf') | CaselessKeyword('End If') | CaselessKeyword('New') | \
                                  CaselessKeyword('#If') | CaselessKeyword('#Else') | CaselessKeyword('#ElseIf') | CaselessKeyword('#End If') | \
                                  CaselessKeyword('Exit') | CaselessKeyword('Type') | CaselessKeyword('As') | CaselessKeyword("ByVal") | \
                                  CaselessKeyword('While') | CaselessKeyword('Do') | CaselessKeyword('Until') | CaselessKeyword('Select') | \
                                  CaselessKeyword('Case') | CaselessKeyword('On') | CaselessKeyword('End') 

unknown_statement = NotAny(known_keywords_statement_start) + \
                    Combine(OneOrMore(CharsNotIn('":\'\x0A\x0D') | quoted_string_keep_quotes),
                            adjacent=False).setResultsName('text')
unknown_statement.setParseAction(UnknownStatement)




class Attribute_Statement(VBA_Object):

    def __init__(self, original_str, location, tokens):
        super(Attribute_Statement, self).__init__(original_str, location, tokens)
        self.name = tokens.name
        self.value = tokens.value
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed %r' % self)

    def __repr__(self):
        return 'Attribute %s = %r' % (self.name, self.value)

    
quoted_identifier = Combine(Suppress('"') + identifier + Suppress('"'))
quoted_identifier.setParseAction(lambda t: str(t[0]))

attribute_statement = CaselessKeyword('Attribute').suppress() + lex_identifier('name') + Suppress('=') + literal('value')
attribute_statement.setParseAction(Attribute_Statement)


class Option_Statement(VBA_Object):
    def __init__(self, original_str, location, tokens):
        super(Option_Statement, self).__init__(original_str, location, tokens)
        self.name = tokens.name
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed %r as Option_Statement' % self)

    def __repr__(self):
        return 'Option %s' % (self.name)

option_statement = CaselessKeyword('Option').suppress() + unrestricted_name + Optional(unrestricted_name)
option_statement.setParseAction(Option_Statement)



type_expression = lex_identifier + Optional('.' + lex_identifier)


type_declaration_composite = Optional(CaselessKeyword('Public') | CaselessKeyword('Private')) + CaselessKeyword('Type') + \
                             lex_identifier + Suppress(EOS) + \
                             OneOrMore(lex_identifier + \
                                       Optional(Suppress(Literal('(') + Optional(expr_list) + Literal(')'))) + \
                                       Optional(Suppress(Literal('(') + expression + CaselessKeyword("To") + expression + Literal(')'))) + \
                                       CaselessKeyword('As') + reserved_type_identifier + \
                                       Suppress(Optional("*" + (decimal_literal | lex_identifier))) + Suppress(EOS)) + \
                             CaselessKeyword('End') + CaselessKeyword('Type') + \
                             ZeroOrMore( Literal(':') + (CaselessKeyword('Public') | CaselessKeyword('Private')) + CaselessKeyword('Type') + \
                                         lex_identifier + Optional(Suppress(Literal('(') + Optional(expr_list) + Literal(')'))) + Suppress(EOS) + \
                                         OneOrMore(lex_identifier + CaselessKeyword('As') + reserved_type_identifier + Suppress(EOS)) + \
                                         CaselessKeyword('End') + CaselessKeyword('Type') )

type_declaration = type_declaration_composite



array_designator = Literal("(") + Literal(")")
function_type = CaselessKeyword("as") + type_expression + Optional(array_designator)


class Parameter(VBA_Object):

    def __init__(self, original_str, location, tokens):
        super(Parameter, self).__init__(original_str, location, tokens)
        self.name = tokens.name
        self.my_type = tokens.type
        self.init_val = tokens.init_val
        self.is_array = False
        self.mechanism = str(tokens.mechanism)
        self.is_optional = (len(tokens.is_optional) > 0)
        if (('(' in str(tokens)) and (')' in str(tokens))):
            self.mechanism = 'ByRef'
            self.is_array = True
        if (len(self.mechanism) == 0):
            self.mechanism = 'ByRef'
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed %r as Parameter' % self)

    def __repr__(self):
        r = ""
        if (self.mechanism):
            r += str(self.mechanism) + " "
        r += str(self.name)
        if self.my_type:
            r += ' as ' + str(self.my_type)
        if (self.init_val):
            r += ' = ' + str(self.init_val)
        return r

    def to_python(self, context, params=None, indent=0):
        name_str = str(self.name)
        init_str = ""
        if ((self.init_val is not None) and (len(str(self.init_val)) > 0)):
            init_str = "=" + to_python(self.init_val, context=context)
        r = name_str + init_str
        return r
    

default_value = Literal("=").suppress() + expr_const('default_value')

parameter_mechanism = CaselessKeyword('ByVal') | CaselessKeyword('ByRef')

optional_prefix = (CaselessKeyword("optional") + parameter_mechanism) \
                  | (parameter_mechanism + CaselessKeyword("optional"))

parameter_type = Optional(array_designator) + CaselessKeyword("as").suppress() \
                 + (type_expression | CaselessKeyword("Any"))

untyped_name_param_dcl = identifier + Optional(parameter_type)


parameter = Optional(CaselessKeyword("optional").suppress())('is_optional') + Optional(parameter_mechanism('mechanism')) + Optional(CaselessKeyword("ParamArray").suppress()) + \
            TODO_identifier_or_object_attrib('name') + \
            Optional(CaselessKeyword("(") + ZeroOrMore(" ").suppress() + CaselessKeyword(")")) + \
            Optional(CaselessKeyword('as').suppress() + (lex_identifier('type') ^ reserved_complex_type_identifier('type'))) + \
            Optional('=' + expression('init_val'))
parameter.setParseAction(Parameter)

parameters_list = delimitedList(parameter, delim=',')



statement_label_definition = LineStart() + ((identifier('label_name') + Suppress(":"))
                                            | (integer('label_int') + Optional(Suppress(":"))))
statement_label = identifier | integer
statement_label_list = delimitedList(statement_label, delim=',')




class TaggedBlock(VBA_Object):

    def __init__(self, original_str, location, tokens):
        super(TaggedBlock, self).__init__(original_str, location, tokens)
        if (tokens is None):
            return
        self.block = tokens.block
        self.label = str(tokens.label).replace(":", "")
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed %r' % self)

    def __repr__(self):
        return 'Tagged Block: %s: %s' % (repr(self.label), repr(self.block))

    def eval(self, context, params=None):

        if (context.exit_func):
            return

        do_const_assignments(self.block, context)

        for s in self.block:
            if (not hasattr(s, "eval")):
                continue
            s.eval(context, params=params)

            if (context.must_handle_error()):
                break
            context.clear_error()

            if (context.goto_executed or s.exited_with_goto):
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("GOTO executed. Go to next loop iteration.")
                break
            
        context.handle_error(params)

tagged_block = Forward()
label_statement = Forward()
        
statement = Forward()
statement_no_orphan = Forward()
statements_line = Forward()
statements_line_no_eos = Forward()
statement_restricted = Forward()
external_function = Forward()

block_statement = rem_statement | external_function | (statement_no_orphan ^ statements_line_no_eos)
simple_call_list = Forward()
statement_block = ZeroOrMore(simple_call_list ^ tagged_block ^ (block_statement + EOS.suppress()))
statement_block_not_empty = OneOrMore(tagged_block ^ (block_statement + EOS.suppress()))
tagged_block <<= label_statement('label') + Suppress(EOS) + statement_block('block')
tagged_block.setParseAction(TaggedBlock)

def do_const_assignments(code_block, context):

    if (not isinstance(code_block, list)):
        code_block = [code_block]

    for s in code_block:
        if (isinstance(s, Dim_Statement) and (s.decl_type.lower() == "const")):
            log.info("Pre-running const assignment '" + str(s) + "'")
            s.eval(context)


class Dim_Statement(VBA_Object):

    def __init__(self, original_str, location, tokens):
        super(Dim_Statement, self).__init__(original_str, location, tokens)
        
        self.decl_type = ""
        var_info = []
        for f in tokens:
            if (str(f).lower() == "const"):
                self.decl_type = str(f)
            if (isinstance(f, pyparsing.ParseResults)):
                var_info.append(f)
        tokens = var_info
        
        self.init_val = "NULL"
        last_var = tokens[0]
        if ((len(last_var) >= 3) and
            (last_var[len(last_var) - 2] == '=')):
            self.init_val = last_var[len(last_var) - 1]
            
        self.variables = []
        for var in tokens:

            is_array = False
            size = None
            if ((len(var) > 1) and (var[1] == '(')):
                is_array = True
                if (isinstance(var[2], int)):
                    size = var[2]
                if ((len(var) > 3) and (isinstance(var[3], int))):
                    size = var[3]

            curr_type = None
            if ((len(var) > 1) and
                (var[-1:][0] != ")") and
                (var[-1:][0] != self.init_val)):
                curr_type = var[-1:][0]

            self.variables.append((var[0], is_array, curr_type, size))

        tmp_vars = []
        final_type = self.variables[len(self.variables) - 1][2]
        for var in self.variables:
            curr_type = var[2]
            if (curr_type is None):
                curr_type = final_type
            tmp_vars.append((var[0], var[1], curr_type, var[3]))
        self.variables = tmp_vars
        
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed %r as Dim_Statement' % str(self))

    def __repr__(self):
        r = "Dim "
        first = True
        for var in self.variables:
            if (not first):
                r += ", "
            first = False
            r += str(var[0])
            if (var[1]):
                r += "("
                if (var[3] is not None):
                    r += str(var[3])
                r += ")"
            if (var[2]):
                r += " As " + str(var[2])
        if (self.init_val is not None):
            r += " = " + str(self.init_val)
        return r

    def to_python(self, context, params=None, indent=0):        

        init_val = ''
        if (self.init_val is not None):
            init_val = to_python(self.init_val, context=context)
            
        r = ""
        for var in self.variables:

            curr_init_val = init_val
            curr_type = var[2]
            if (curr_type is not None):

                if ((curr_type == "Long") or (curr_type == "Integer")):
                    curr_init_val = "0"
                if (curr_type == "String"):
                    curr_init_val = '""'
                if (curr_type == "Boolean"):
                    curr_init_val = "False"
                
                if (var[1]):
                    curr_type += " Array"
                    curr_init_val = []
                    if ((var[3] is not None) and
                        ((curr_type == "Byte Array") or (curr_type == "Integer Array"))):
                        curr_init_val = str([0] * (var[3] + 1))
                    if ((var[3] is not None) and (curr_type == "String Array")):
                        curr_init_val = str([''] * var[3])

            elif (var[1]):
                curr_init_val = str([])

            if ((context.global_scope) and (curr_init_val is None)):
                curr_init_val = "0"

            if (context.contains(var[0], local=True)):
                vm_val = context.get(var[0])
                if (vm_val == "__ALREADY_SET__"):
                    vm_val = context.get("__ORIG__" + var[0])
                curr_init_val = to_python(vm_val, context)

            if (curr_init_val == '"NULL"'):
                curr_init_val = "0"

            context.set(var[0], "__ALREADY_SET__", var_type=curr_type)
            context.set(var[0], "__ALREADY_SET__", var_type=curr_type, force_global=True)
            context.set("__ORIG__" + var[0], curr_init_val, force_local=True)
            context.set("__ORIG__" + var[0], curr_init_val, force_global=True)
                
            var_name = utils.fix_python_overlap(str(var[0]))
            r += " " * indent + var_name + " = " + str(curr_init_val) + "\n"

        return r
    
    def eval(self, context, params=None):

        if (context.exit_func):
            return

        init_val = ''
        if (self.init_val is not None):
            init_val = eval_arg(self.init_val, context=context)
            
        for var in self.variables:

            curr_init_val = init_val
            curr_type = var[2]
            if (curr_type is not None):

                curr_type = str(curr_type)
                if ((curr_type == "Long") or (curr_type == "Integer")):
                    curr_init_val = 0
                if (curr_type == "String"):
                    curr_init_val = ''
                if (curr_type == "Boolean"):
                    curr_init_val = False
                
                if (var[1]):
                    curr_type += " Array"
                    curr_init_val = []
                    if ((var[3] is not None) and
                        ((curr_type == "Byte Array") or (curr_type == "Integer Array"))):
                        curr_init_val = [0] * (var[3] + 1)
                    if ((var[3] is not None) and (curr_type == "String Array")):
                        curr_init_val = [''] * var[3]

            elif (var[1]):
                curr_init_val = []
                if (var[3] is not None):
                    curr_init_val = [0] * (var[3] + 1)

            if ((context.global_scope) and (curr_init_val is None)):
                curr_init_val = "NULL"

            if (context.contains(var[0], local=True)):
                curr_init_val = context.get(var[0])
                
            is_const = (self.decl_type.lower() == "const")
            is_local = context.in_procedure and (not is_const)
            context.set(var[0], curr_init_val, var_type=curr_type, force_global=is_const, force_local=is_local)
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("DIM " + str(var[0]) + " As " + str(curr_type) + " = " + str(curr_init_val))
    








constant_expression = expression
lower_bound = constant_expression + CaselessKeyword('to').suppress()
upper_bound = constant_expression
dim_spec = Optional(lower_bound) + upper_bound
bounds_list = delimitedList(dim_spec)
array_dim = '(' + Optional(bounds_list('bounds')) + ')'
constant_name = simple_name_expression
string_length = constant_name | integer
fixed_length_string_spec = CaselessKeyword("string").suppress() + Suppress("*") + string_length
type_spec = fixed_length_string_spec | type_expression
as_type = CaselessKeyword('As').suppress() + type_spec
defined_type_expression = simple_name_expression
class_type_name = defined_type_expression
as_auto_object = CaselessKeyword('as').suppress() + CaselessKeyword('new').suppress() + expression
as_clause = as_auto_object | as_type
array_clause = array_dim('bounds') + Optional(as_clause)
untyped_variable_dcl = Suppress(Optional(CaselessKeyword("WithEvents"))) + identifier + Optional(array_clause('bounds') | as_clause)
typed_variable_dcl = Suppress(Optional(CaselessKeyword("WithEvents"))) + typed_name + Optional(array_dim)
variable_dcl = (typed_variable_dcl | untyped_variable_dcl) + Optional('=' + expression('expression'))
variable_declaration_list = delimitedList(Group(variable_dcl))
local_variable_declaration = (CaselessKeyword("Dim") | CaselessKeyword("Static") | (Suppress(Optional(Literal("#"))) + CaselessKeyword("Const"))) + \
                             Optional(CaselessKeyword("Shared")).suppress() + variable_declaration_list

dim_statement = local_variable_declaration
dim_statement.setParseAction(Dim_Statement)


class Global_Var_Statement(Dim_Statement):
    pass

public_private = Forward()
global_variable_declaration = Optional(public_private) + \
                              Optional(CaselessKeyword("Shared")).suppress() + \
                              Optional(CaselessKeyword("Const")) + \
                              variable_declaration_list
global_variable_declaration.setParseAction(Global_Var_Statement)


class Let_Statement(VBA_Object):

    def __init__(self, original_str, location, tokens):
        super(Let_Statement, self).__init__(original_str, location, tokens)

        self.string_op = None
        self.index = None
        self.index1 = None
        if (original_str is None):
            return

        self.name = tokens.name
        string_ops = set(["mid", "mid$"])
        self.string_op = None
        if (hasattr(self.name, "__len__") and
            (len(self.name) > 0) and
            (self.name[0].lower() in string_ops)):
            self.string_op = {}
            self.string_op["op"] = self.name[0].lower()
            self.string_op["args"] = self.name[1:]
        self.expression = tokens.expression
        self.index = None
        if (tokens.index != ''):
            self.index = tokens.index
        self.index1 = None
        if (tokens.index1 != ''):
            self.index1 = tokens.index1
        self.op = tokens["op"]
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed %r as Let_Statement' % self)

    def __repr__(self):
        if (self.index is None):
            return 'Let %s %s %r' % (self.name, self.op, self.expression)
        else:
            return 'Let %s(%r) %s %r' % (self.name, self.index, self.op, self.expression)

    def to_python(self, context, params=None, indent=0):        

        r = ""
        try:

            var_val = context.get(self.name, global_only=True)
            if ((not (isinstance(var_val, procedures.Function) or
                     isinstance(var_val, procedures.Sub) or
                     isinstance(var_val, VbaLibraryFunc))) and
                (var_val != "__ALREADY_SET__")):

                spaces = " " * indent
                r += "global " + utils.fix_python_overlap(str(self.name)) + \
                     "\n" + spaces

        except KeyError:
            pass

        python_var_name = str(self.name)
        if (self.index is None):
            
            if ((self.string_op is not None) and
                ((self.string_op["op"] == "mid") or (self.string_op["op"] == "mid$"))):
                
                args = self.string_op["args"]
                if (len(args) < 3):
                    return "ERROR: Wrong # args to mid. " + str(self)
                the_str_var = to_python(args[0], context)
                start = to_python(args[1], context)
                size = to_python(args[2], context)
                rhs = to_python(self.expression, context)
                
                r += the_str_var + " = " + the_str_var + "[:" + start + "-1] + " + rhs + " + " + the_str_var + "[(" + start + "-1 + " + size + "):]"

            elif (context.get_type(self.name) == "Byte Array"):
                val = "coerce_to_int_list(" + to_python(self.expression, context, params=params) + ")"
                r += utils.fix_python_overlap(python_var_name) + " " + str(self.op) + " " + val

            elif (context.get_type(self.name) == "String"):
                val = "coerce_to_str(" + to_python(self.expression, context, params=params) + ")"
                r += utils.fix_python_overlap(python_var_name) + " " + str(self.op) + " " + val

            else:
                r += utils.fix_python_overlap(python_var_name) + " " + str(self.op) + " " + to_python(self.expression, context, params=params)
                
        else:
            py_var = utils.fix_python_overlap(python_var_name)
            if (py_var.startswith(".")):
                py_var = py_var[1:]
            index = to_python(self.index, context, params=params)
            indices = [index]
            if (self.index1 is not None):
                indices.append(to_python(self.index1, context, params=params))
            val = to_python(self.expression, context, params=params)
            op = str(self.op)
            index_str = ""
            first = True
            for i in indices:
                if (not first):
                    index_str += ", "
                first = False
                index_str += i
            index_str = "[" + index_str + "]"
            if (op == "="):
                r += py_var + " = update_array(" + py_var + ", " + index_str + ", " + val + ")"
            else:
                r += py_var + "[" + index + "] " + op + " " + val

        context.set(self.name, "__ALREADY_SET__")
                
        if (r.startswith(".")):
            r = r[1:]
        r = " " * indent + r
        return r
        
    def _handle_change_callback(self, var_name, context):

        var_name = str(var_name)
        if ("." in var_name):
            var_name = var_name[var_name.rindex(".") + 1:]

        callback_name = var_name + "_Change"

        try:

            callback = context.get(callback_name)
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("Found change callback " + callback_name)

            if ((is_procedure(callback)) and
                (callback_name not in context.skip_handlers)):

                context.skip_handlers.add(callback_name)
                
                log.info("Running change callback " + callback_name)
                callback.eval(context)
                
                context.skip_handlers.remove(callback_name)

        except KeyError:

            pass

    def _make_let_statement(self, the_str_var, mod_str):

        tmp_let = Let_Statement(None, None, None)

        tmp_let.name = the_str_var
        if (isinstance(the_str_var, Function_Call)):

            tmp_let.name = the_str_var.name

            if (len(the_str_var.params) > 0):
                tmp_let.index = the_str_var.params[0]
            if (len(the_str_var.params) > 1):
                tmp_let.index1 = the_str_var.params[1]

        tmp_let.expression = mod_str
        tmp_let.op = "="

        return tmp_let
        
    def _handle_string_mod(self, context, rhs):

        if (self.string_op is None):
            return False

        if ((self.string_op["op"] == "mid") or (self.string_op["op"] == "mid$")):

            args = self.string_op["args"]
            if (len(args) < 3):
                return False
            the_str = eval_arg(args[0], context)
            the_str_var = args[0]
            start = utils.int_convert(eval_arg(args[1], context), leave_alone=True)
            size = utils.int_convert(eval_arg(args[2], context), leave_alone=True)
            
            if ((not isinstance(the_str, str)) and (not isinstance(the_str, list))):
                context.report_general_error("Assigning " + str(self.name) + " failed. " + str(the_str_var) + " not str or list.")
                return False
            if (type(the_str) != type(rhs)):
                context.report_general_error("Assigning " + str(self.name) + " failed. " + str(type(the_str)) + " != " + str(type(rhs)))
                return False
            if (not isinstance(start, int)):
                context.report_general_error("Assigning " + str(self.name) + " failed. Start is not int (" + str(type(start)) + ").")
                return False
            if (not isinstance(size, int)):
                context.report_general_error("Assigning " + str(self.name) + " failed. Size is not int (" + str(type(size)) + ").")
                return False
            if (((start-1 + size) > len(the_str)) or (start < 1)):
                context.report_general_error("Assigning " + str(self.name) + " failed. " + str(start + size) + " out of range.")
                return False

            vb_rhs = vb_str.VbStr(rhs, context.is_vbscript)
            
            if (vb_rhs.len() > size):
                vb_rhs = vb_rhs.get_chunk(0, size)
            if (vb_rhs.len() < size):
                size = vb_rhs.len()
                
            vb_the_str = vb_str.VbStr(the_str, context.is_vbscript)
            mod_str = vb_the_str.update_chunk(start - 1, start - 1 + size, vb_rhs).to_python_str()

            tmp_let = self._make_let_statement(the_str_var, mod_str)
            tmp_let.eval(context)
            return True

        return False

    def _handle_autoincrement(self, lhs, rhs):

        r = "NULL"
        try:
            if (self.op == "+="):
                r = (lhs + rhs)
            elif (self.op == "-="):
                r = (lhs - rhs)
        except:
            pass
        return r

    def _handle_lhs_call(self, context):

        if (self.index is None):
            return None
        func_name = str(self.name)
        if ("." in func_name):
            func_name = func_name[func_name.rindex(".") + 1:]
        func_call_str = func_name + "(" + str(self.index).replace("'", "") + ")"
        try:
            func_call = function_call.parseString(func_call_str, parseAll=True)[0]
            return func_call.eval(context)
        except ParseException:
            pass
        return None
    
    def eval(self, context, params=None):

        if (context.exit_func):
            return
        
        if ((context.contains(self.name)) and
            (isinstance(context.get(self.name), procedures.Function))):
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("Adding uninitialized '" + str(self.name) + "' function return var to local context.")
            context.set(self.name, 'NULL', force_local=True)
        
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('try eval expression: %s' % self.expression)
        rhs_type = context.get_type(str(self.expression))
        value = eval_arg(self.expression, context=context)
        if (context.have_error()):
            log.warn('Short circuiting assignment %s due to thrown VB error.' % str(self))
            return
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('eval expression: %s = %s' % (self.expression, value))

        if (self.name == ".Text"):

            try:
                tmp_str = filter(isprint, str(value).strip())
                value = base64.b64decode(tmp_str)
            except Exception as e:
                log.warning("base64 conversion of '" + str(value) + "' failed. " + str(e))

        if ((str(self.name).endswith(".Arguments")) or
            (str(self.name).endswith(".Path"))):
            context.report_action(self.name, value, 'Possible Scheduled Task Setup', strip_null_bytes=True)
        if (str(self.name).endswith(".CommandLine")):
            context.report_action('Run Command', value, self.name, strip_null_bytes=True)
            
        if (self._handle_string_mod(context, value)):
            return

        if (str(self.name).endswith("OnSheetActivate")):

            func_name = str(self.expression).strip()
            try:
                func = context.get(func_name)
                log.info("Emulating OnSheetActivate handler function " + func_name + "...")
                func.eval(context)
                return
            except KeyError:
                context.report_general_error("WARNING: Cannot find OnSheetActivate handler function %s" % func_name)

        if ((self.op == "+=") or (self.op == "-=")):
            lhs = context.get(self.name)
            value = self._handle_autoincrement(lhs, value)

        if (self.index is None):

            if ((context.get_type(self.name) == "Byte Array") and
                (isinstance(value, str))):

                if (value != "NULL"):

                    bad_byte_count = 0
                    for c in value:
                        if (not isprint(c)):
                            bad_byte_count += 1
                    is_raw_data = ((len(value) > 0) and (((bad_byte_count + 0.0)/len(value)) > .2))
                    
                    tmp = []
                    pos = 0
                    for c in value:

                        tmp.append(ord(c))

                        if ((not isinstance(value, from_unicode_str)) and (not is_raw_data)):
                            tmp.append(0)

                    value = tmp

                else:
                    return
                    
            elif ((context.get_type(self.name) == "String") and
                  (isinstance(value, list))):

                all_ints = True
                for i in value:
                    if (not isinstance(i, int)):
                        all_ints = False
                        break
                if (all_ints):
                    try:
                        tmp = ""
                        pos = 0
                        step = 2
                        if ((rhs_type == "Byte Array") or
                            (rhs_type == "Byte")):
                            step = 1
                        while (pos < len(value)):

                            c = value[pos]
                            if (c == 0):
                                pos += step
                                continue

                            tmp += chr(c)
                            pos += step
                        value = tmp
                    except:
                        pass

                all_chars = True
                for i in value:
                    if ((i is not None) and
                        ((not isinstance(i, str)) or (len(i) > 1))):
                        all_chars = False
                        break
                if (all_chars):
                    tmp = ""
                    for i in value:
                        if (i is not None):
                            tmp += i
                    value = tmp

            elif (((context.get_type(self.name) == "Integer") or
                   (context.get_type(self.name) == "Long")) and
                  (isinstance(value, str))):
                try:
                    if (value == "NULL"):
                        value = 0
                    else:
                        value = int(value)
                except:
                    context.report_general_error("Cannot convert '" + str(value) + "' to int. Defaulting to 0.")
                    value = 0

            if (value != "ERROR"):

                val_repr = repr(value)
                if (len(val_repr) > 120):
                    val_repr = val_repr[:120] + "..."
                log.info("Assign %s <- value type=%s len=%d repr=%s",
                         self.name, type(value).__name__,
                         len(value) if isinstance(value, str) else 0,
                         val_repr)
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug('setting %s = %s' % (self.name, value))
                context.set(self.name, value)

                self._handle_change_callback(self.name, context)

            else:

                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug('Not setting ' + self.name + ", eval of RHS gave an error.")

        else:

            if (((context.get_type(self.name) == "Long Array") or
                 (context.get_type(self.name) == "Integer Array")) and
                (isinstance(value, str))):

                if (value != "NULL"):

                    num = "not an integer"
                    try:
                        expr = expression.parseString(value, parseAll=True)[0]
                        num = str(expr)
                        if (hasattr(expr, "eval")):
                            num = str(expr.eval(context))
                    except ParseException:
                        context.report_general_error("Cannot parse '" + value + "' to integer.")
                    if (not num.isdigit()):
                        context.report_general_error("Cannot convert '" + value + "' to integer. Setting to 0.")
                        num = 0
                    else:
                        num = int(num)
                    value = num
                        
            index = utils.int_convert(eval_arg(self.index, context=context))
            index1 = None
            if (self.index1 is not None):
                index1 = utils.int_convert(eval_arg(self.index1, context=context))
                
            arr_var = None
            try:
                arr_var = context.get(self.name)
            except KeyError:
                context.report_general_error("WARNING: Cannot find array variable %s" % self.name)

                call_r = self._handle_lhs_call(context)
                if (call_r is not None):
                    return

                arr_var = []
                
            if ((not isinstance(arr_var, list)) and (not isinstance(arr_var, str))):

                arr_var = []

            if ((isinstance(arr_var, list)) and (index >= 0)):

                if (index >= len(arr_var)):
                    arr_var.extend([0] * (index - len(arr_var) + 1))
                if (index1 is not None):
                    if (not isinstance(arr_var[index], list)):
                        arr_var[index] = []
                    if (index1 >= len(arr_var[index])):
                        arr_var[index].extend([0] * (index1 - len(arr_var[index])))
                
                if (index1 is None):
                    arr_var = arr_var[:index] + [value] + arr_var[(index + 1):]
                else:
                    new_arr = arr_var[index]
                    new_arr = new_arr[:index1] + [value] + new_arr[(index1 + 1):]
                    arr_var[index] = new_arr

            if ((isinstance(arr_var, str)) or (isinstance(arr_var, unicode))):

                if (index >= len(arr_var)):
                    arr_var += "\0"*(index - len(arr_var))
                
                if ((isinstance(value, str)) or (isinstance(value, unicode))):
                    arr_var = arr_var[:index] + value + arr_var[(index + 1):]
                elif (isinstance(value, int)):
                    try:
                        arr_var = arr_var[:index] + chr(value) + arr_var[(index + 1):]
                    except Exception as e:
                        log.error(str(e))
                        context.report_general_error(str(value) + " cannot be converted to ASCII.")
                else:
                    context.report_general_error("Unhandled value type " + str(type(value)) + " for array update.")
                        
            if (value != "ERROR"):

                context.set(self.name, arr_var)

                self._handle_change_callback(self.name, context)

            else:

                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug('Not setting ' + self.name + ", eval of RHS gave an error.")


class LSet_Statement(Let_Statement):
    pass
                


string_modification = (CaselessKeyword('Mid') | CaselessKeyword('Mid$')) + Optional(Suppress('(')) + expr_list('params') + Optional(Suppress(')'))

let_statement = (
    Optional(CaselessKeyword('Let') | CaselessKeyword('Set')).suppress()
    + Optional(Suppress(CaselessKeyword('Const')))
    + (
        (
            (
                Optional(Suppress('(')) + TODO_identifier_or_object_attrib('name') + Optional(Suppress(')'))
                + (Optional(Suppress('(') + Optional(expression('index')) + Optional(',' + expression('index1')) + Suppress(')')) ^ \
                   Optional(Suppress('(') + expression('index') + Suppress(')') + Suppress('(') + expression('index1') + Suppress(')'))) \
            )
            ^ member_access_expression('name')
            ^ string_modification('name')
        )
        |
        (
            Literal(".")
            + (
                TODO_identifier_or_object_attrib_loose('name')
                + Optional(
                    Suppress('(')
                    + Optional(expression('index'))
                    + Optional(',' + expression('index1'))
                    + Suppress(')')
                )
            )
            ^ member_access_expression('name')
            ^ string_modification('name')
        )
    )
    + (Literal('=') | Literal('+=') | Literal('-='))('op')
    + (expression('expression') ^ boolean_expression('expression'))
)
let_statement.setParseAction(Let_Statement)

lset_statement = (
    CaselessKeyword('LSet').suppress()
    + Optional(Suppress(CaselessKeyword('Const')))
    + Optional(".")
    + (
        (
            TODO_identifier_or_object_attrib('name')
            + Optional(
                Suppress('(')
                + Optional(expression('index'))
                + Optional(',' + expression('index1'))
                + Suppress(')')
            )
        )
        ^ member_access_expression('name')
        ^ string_modification('name')
    )
    + (Literal('=') | Literal('+=') | Literal('-='))('op')
    + (expression('expression') ^ boolean_expression('expression'))
)
lset_statement.setParseAction(LSet_Statement)


class Prop_Assign_Statement(VBA_Object):
    def __init__(self, original_str, location, tokens):
        super(Prop_Assign_Statement, self).__init__(original_str, location, tokens)
        self.prop = tokens.prop
        self.param = tokens.param
        self.value = tokens.value
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed %r as Prop_Assign_Statement' % self)

    def __repr__(self):
        return str(self.prop) + " " + str(self.param) + ":=" + str(self.value)

    def to_python(self, context, params=None, indent=0):
        return " " * indent + "pass"
    
    def eval(self, context, params=None):
        if (context.exit_func):
            return

prop_assign_statement = (
    Optional(Suppress("."))
    + (member_access_expression("prop") ^ lex_identifier("prop"))
    + lex_identifier('param')
    + Suppress(':=')
    + expression('value')
    + ZeroOrMore(',' + lex_identifier('param') + Suppress(':=') + expression('value'))
)
prop_assign_statement.setParseAction(Prop_Assign_Statement)


class For_Statement(VBA_Object):

    def __init__(self, original_str, location, tokens):
        super(For_Statement, self).__init__(original_str, location, tokens)
        self.is_loop = True
        self.name = tokens.name
        self.start_value = tokens.start_value
        self.end_value = tokens.end_value
        self.step_value = tokens.get('step_value', 1)
        if self.step_value != 1:
            self.step_value = self.step_value[0]
        self.statements = tokens.statements
        self.body = self.statements
        self.only_atomic = None
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed %r as %s' % (self, self.__class__.__name__))

    def __repr__(self):
        return 'For %s = %r to %r step %r' % (self.name,
                                              self.start_value, self.end_value, self.step_value)

    def _get_loop_indices(self, context):

        start = eval_arg(self.start_value, context=context)
        if (isinstance(start, basestring)):
            start = utils.int_convert(start)

        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('FOR loop - start: %r = %r' % (self.start_value, start))

        end = eval_arg(self.end_value, context=context)
        if (isinstance(end, basestring)):
            end = utils.int_convert(end)
        if (end is None):
            log.warning("Not emulating For loop. Loop end '" + str(self.end_value) + "' evaluated to None.")
            return (None, None, None)
            
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('FOR loop - end: %r = %r' % (self.end_value, end))

        if self.step_value != 1:
            step = eval_arg(self.step_value, context=context)
            if (isinstance(step, basestring)):
                step = utils.int_convert(step)
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug('FOR loop - step: %r = %r' % (self.step_value, step))
        else:
            step = 1

        if ((start > end) and (step > 0)):
            step = step * -1

        return (start, end, step)
        
    def _get_loop_indices_python(self, context):

        start = to_python(self.start_value, context=context)
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('FOR loop - start: %r = %r' % (self.start_value, start))

        end = to_python(self.end_value, context=context)
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('FOR loop - end: %r = %r' % (self.end_value, end))

        if self.step_value != 1:
            step = to_python(self.step_value, context=context)
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug('FOR loop - step: %r = %r' % (self.step_value, step))
        else:
            step = "1"

        return (start, end, step)
    
    def to_python(self, context, params=None, indent=0):

        loop_var = str(self.name)

        tmp_context = context
        tmp_context.set(loop_var, "__LOOP_VAR__", force_local=True)
        tmp_context.set(loop_var, "__LOOP_VAR__", force_global=True)        
        
        indent_str = " " * indent
        
        start, end, step = self._get_loop_indices_python(context)

        if (len(self.statements) == 0):
            r = indent_str + loop_var + " = " + str(end) + "\n"
            return r
        
        rev_code = ""
        if (step < 0):
            rev_code = "[::-1]"
            step = abs(step)
        loop_start = indent_str + "exit_all_loops = False\n"
        loop_start += indent_str + loop_var + " = " + str(start) + "\n"
        loop_start += indent_str + "while (((" + loop_var + " <= coerce_to_int(" + str(end) + ")) and (" + str(step) + " > 0)) or " + \
                      "((" + loop_var + " >= coerce_to_int(" + str(end) + ")) and (" + str(step) + " < 0))):\n"
        loop_start += indent_str + " " * 4 + "if exit_all_loops:\n"
        loop_start += indent_str + " " * 8 + "break\n"
        loop_start = indent_str + "# Start emulated loop.\n" + loop_start

        loop_init, prog_var = _loop_vars_to_python(self, tmp_context, indent)
            
        save_vals = _updated_vars_to_python(self, context, indent)
        
        loop_body = ""
        end_var = str(end)
        loop_body += indent_str + " " * 4 + "if (int(float(" + loop_var + ")/(coerce_to_int(" + end_var + ") if coerce_to_int(" + end_var + ") != 0 else 1)*100) == " + prog_var + "):\n"
        body_escaped = str(self).replace('"', '\\"').replace("\\n", " :: ")
        loop_body += indent_str + " " * 8 + "safe_print(str(int(float(" + loop_var + ")/(coerce_to_int(" + end_var + ") if coerce_to_int(" + end_var + ") != 0 else 1)*100)) + \"% done with loop " + body_escaped + "\")\n"
        loop_body += indent_str + " " * 8 + prog_var + " += 1\n"
        body_str = to_python(self.statements, tmp_context, params=params, indent=indent+4, statements=True)
        if (body_str.strip() == '""'):
            body_str = "\n"
        loop_body += body_str
        loop_body += indent_str + " " * 4 + loop_var + " += " + str(step) + "\n"
        
        python_code = loop_init + "\n" + \
                      loop_start + "\n" + \
                      loop_body + "\n" + \
                      save_vals + "\n"

        return python_code

    def _handle_medium_loop(self, context, params, end, step):


        all_static_assigns = True
        for s in self.statements:

            if (not isinstance(s, Let_Statement)):
                all_static_assigns = False
                break

            is_constant = (str(s.expression).isdigit())
            if ((not isinstance(s.expression, SimpleNameExpression)) and (not is_constant)):
                all_static_assigns = False
                break

            if (str(s.expression).strip().lower() == str(self.name).strip().lower()):
                all_static_assigns = False
                break

        if (not all_static_assigns):
            return False
        
        log.info("Short circuited loop. " + str(self))
        for s in self.statements:

            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug('FOR loop eval statement: %r' % s)
            if (not isinstance(s, VBA_Object)):
                continue
            s.eval(context=context)
                
            if (context.must_handle_error()):
                break
            context.clear_error()

        context.handle_error(params)

        context.set(self.name, end + step)
                     
        return True
                
    def _handle_simple_loop(self, context, start, end, step):

        
        if (len(self.statements) != 1):
            return (None, None)

        body_raw = str(self.statements[0]).replace("Let ", "")
        body = body_raw.replace("(", "").replace(")", "").strip()
        if (body.startswith("Dim ")):

            self.statements[0].eval(context)

            context.set(self.name, end + step)

            log.info("Short circuited Dim only loop " + str(self))
            return ("N/A", "N/A")

        if (body.endswith(" = " + str(self.name))):

            fields = body_raw.split(" ")
            var = fields[0].strip()
            if (("(" not in var) and (")" not in var)):
                context.set(self.name, end + step)
                return (var, end + step)

        body1 = body.replace("Call_Statement:", "").strip()
        if (body1.startswith("Debug.Print")):

            self.statements[0].eval(context)

            context.set(self.name, end + step)

            log.info("Short circuited Debug.Print only loop " + str(self))
            return ("N/A", "N/A")

        if (re.search(r"'On', 'Error', 'Goto',", body) is not None):

            self.statements[0].eval(context)

            context.set(self.name, end + step)

            log.info("Short circuited 'On Error' only loop " + str(self))
            return ("N/A", "N/A")
            
        fields = body.split(" ")
        if (len(fields) < 5):
            return (None, None)
        op = fields[3].strip()
        var = fields[0].strip()
        if (var != fields[2].strip()):
            return (None, None)
        if (op not in ['+', '-', '*']):
            return (None, None)

        for f in fields[2:]:
            if (f.strip() == str(self.name).strip()):
                return (None, None)
                
        expr_str = ""
        for e in fields[4:]:
            expr_str += " " + e
        num = None
        try:
            expr = expression.parseString(expr_str, parseAll=True)[0]
            num = str(expr)
            if (hasattr(expr, "eval")):
                num = str(expr.eval(context))
        except ParseException:
            return (None, None)
        if (not num.isdigit()):
            return (None, None)
        
        init_val = None
        try:

            init_val = context.get(var)
            if (init_val == "NULL"):
                init_val = 0

            if (not str(init_val).isdigit()):
                return (None, None)
            init_val = int(str(init_val))

        except KeyError:

            init_val = 0

        num_iters = (end - start + 1)/step
            
        try:
            num = int(num)
        except:
            return (None, None)
        r = None
        if (op == "+"):
            r = (var, init_val + num_iters*num)
        elif (op == "-"):
            r = (var, init_val - num_iters*num)
        elif (op == "*"):
            r = (var, init_val * pow(num, num_iters))
        else:
            return (None, None)

        context.set(self.name, end + step)

        return r

    def _no_state_change(self, prev_context, context):

        if (isinstance(self.name, pyparsing.ParseResults)):
            self.name = self.name[0] if self.name else ""

        if ((prev_context is None) or (context is None)):
            return False
        
        if (self.only_atomic is None):
            self.only_atomic = True
            for s in self.statements:
                if (not isinstance(s, VBA_Object)):
                    continue
                if ((not is_simple_statement(s)) and (not s.is_useless)):
                    self.only_atomic = False
        if (not self.only_atomic):
            return False

        prev_context = Context(context=prev_context, _locals=prev_context.locals, copy_globals=True)\
                                        .delete(self.name)\
                                        .delete(self.name.lower())\
                                        .delete("now")\
                                        .delete("application.username")\
                                        .delete("recentfiles.count")\
                                        .delete("activedocument.revisions.count")\
                                        .delete("thisdocument.revisions.count")\
                                        .delete("revisions.count")
        context = Context(context=context, _locals=context.locals, copy_globals=True)\
                                   .delete(self.name)\
                                   .delete(self.name.lower())\
                                   .delete("now")\
                                   .delete("application.username")\
                                   .delete("recentfiles.count")\
                                   .delete("activedocument.revisions.count")\
                                   .delete("thisdocument.revisions.count")\
                                   .delete("revisions.count")

        r = (prev_context == context)
        return r
    
    def eval(self, context, params=None):

        if (context.exit_func):
            return
        
        self.exited_with_goto = False
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('FOR loop: evaluating start, end, step')

        if (len(self.statements) == 0):
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("FOR loop: empty body. Skipping.")
            return

        do_const_assignments(self.statements, context)        
        
        start, end, step = self._get_loop_indices(context)
        if (start is None):
            log.warn("Cannot resolve loop index information, not doing JIT loop emulation.")
            return
            
        context.set(self.name, start)
            
        var, val = self._handle_simple_loop(context, start, end, step)
        if ((var is not None) and (val is not None)):
            log.info("Short circuited loop. Set " + str(var) + " = " + str(val))
            context.set(var, val)
            self.is_useless = True
            return

        do_body_once = self._handle_medium_loop(context, params, end, step)
        if (do_body_once):
            self.is_useless = True
            return

        if (_eval_python(self, context, params=params, add_boilerplate=True)):
            return
        
        if ((VBA_Object.loop_upper_bound > 0) and (end > VBA_Object.loop_upper_bound)):

            if (end > 100000000):
                end = 10
            else:

                end = VBA_Object.loop_upper_bound
            log.warn("FOR loop: upper loop iteration bound exceeded, setting to %r" % end)
        
        context.loop_stack.append(None)
        context.loop_object_stack.append(self)
        my_loop_stack_pos = len(context.loop_stack) - 1

        num_no_change = 0
        prev_context = None

        num_iters_run = 0
        throttle_io_limit = 100

        if (not context.contains(self.name)):
            log.warn("Cannot find loop variable " + str(self.name) + ". Skipping loop.")
            return

        context.clear_general_errors()
        while (((step > 0) and (context.get(self.name) <= end)) or
               ((step < 0) and (context.get(self.name) >= end))):

            context.goto_executed = False
            
            last_index = context.get(self.name)

            if ((num_iters_run < 10) or (num_no_change > 0)):

                if (self._no_state_change(prev_context, context)):
                    num_no_change += 1
                    if (num_no_change >= context.max_static_iters * 5):
                        log.warn("Possible useless For loop detected. Exiting loop.")
                        self.is_useless = True
                        break
                else:
                    num_no_change = 0                
                prev_context = Context(context=context, _locals=context.locals, copy_globals=True)

            num_iters_run += 1
            if ((num_iters_run > throttle_io_limit) and ((num_iters_run % 500) == 0)):
                log.warning("Long running loop. I/O has been throttled.")
            if ((num_iters_run > throttle_io_limit) and (not context.throttle_logging)):
                log.warning("Throttling output logging...")
                context.throttle_logging = True
            if (((num_iters_run < throttle_io_limit) or ((num_iters_run % 5000) == 0)) and
                context.throttle_logging):
                log.warning("Output is throttled...")
                context.throttle_logging = False

            if (context.get_general_errors() > (VBA_Object.loop_upper_bound/10000)):
                log.error("Loop is generating too many errors. Breaking loop.")
                break
                
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug('FOR loop: %s = %r' % (self.name, context.get(self.name)))
            done = False
            for s in self.statements:
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug('FOR loop eval statement: %r' % s)
                if (not isinstance(s, VBA_Object)):
                    continue
                s.eval(context=context)
                
                if ((my_loop_stack_pos >= len(context.loop_stack)) or context.loop_stack[my_loop_stack_pos]):
                    
                    self.exited_with_goto = (context.loop_stack[my_loop_stack_pos] == "GOTO")
                    if (log.getEffectiveLevel() == logging.DEBUG):
                        log.debug("FOR loop: exited loop with 'Exit For'")
                    done = True
                    break

                if (context.must_handle_error()):

                    if (not context.do_next_iter_on_error()):
                        done = True
                    break
                context.clear_error()

                if (context.goto_executed or s.exited_with_goto):
                    if (log.getEffectiveLevel() == logging.DEBUG):
                        log.debug("GOTO executed. Go to next loop iteration.")
                    break
                
            if (done):
                break

            context.clear_error()
            
            val = context.get(self.name)
            try:
                val = int(val)
                step = int(step)
            except Exception as e:
                context.report_general_error("Cannot update loop counter. Breaking loop. " + str(e))
                break
            new_index = val + step
            context.set(self.name, new_index)

            if (((new_index < start) and (step > 0)) or
                ((new_index > start) and (step < 0))):

                log.warn("Possible infinite For loop detected. Exiting loop.")
                break
        
        if (len(context.loop_stack) > 0):
            context.loop_stack.pop()
        if (len(context.loop_object_stack) > 0):
            context.loop_object_stack.pop()
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('FOR loop: end.')
            
        context.handle_error(params)

        context.goto_executed = False
        

bound_variable_expression = TODO_identifier_or_object_attrib


step_clause = CaselessKeyword('Step').suppress() + expression

for_clause = CaselessKeyword("For").suppress() \
             + lex_identifier('name') \
             + Suppress(Optional(CaselessKeyword("As") + type_expression)) \
             + Suppress("=") + expression('start_value') \
             + CaselessKeyword("To").suppress() + expression('end_value') \
             + Optional(step_clause('step_value'))

simple_for_statement = for_clause + Suppress(EOS) + statement_block('statements') \
                       + (CaselessKeyword("Next").suppress() | CaselessKeyword("End").suppress()) \
                       + Optional(lex_identifier) \
                       + FollowedBy(EOS)

simple_for_statement.setParseAction(For_Statement)

bad_next_statement = CaselessKeyword("Next") + Suppress(EOS)

for_start = for_clause + Suppress(EOS)
for_start.setParseAction(For_Statement)

for_end = CaselessKeyword("Next").suppress() + Optional(lex_identifier) + Suppress(EOS)


class For_Each_Statement(VBA_Object):

    def __init__(self, original_str, location, tokens):
        super(For_Each_Statement, self).__init__(original_str, location, tokens)
        self.is_loop = True
        self.statements = tokens.statements
        self.body = self.statements
        self.item = tokens.clause.item
        self.container = tokens.clause.container
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed %r as %s' % (self, self.__class__.__name__))

    def __repr__(self):
        return 'For Each %r In %r ...' % (self.item, self.container)

    def to_python(self, context, params=None, indent=0):

        loop_var = str(self.item)

        tmp_context = context
        tmp_context.set(loop_var, "__LOOP_VAR__")
        tmp_context.set(loop_var, "__LOOP_VAR__", force_global=True)
        
        indent_str = " " * indent
        
        loop_vals = to_python(self.container, tmp_context)

        loop_start = indent_str + "exit_all_loops = False\n"
        loop_start += indent_str + "for " + loop_var + " in " + loop_vals + ":\n"
        loop_start += indent_str + " " * 4 + "if exit_all_loops:\n"
        loop_start += indent_str + " " * 8 + "break\n"
        loop_start = indent_str + "# Start emulated loop.\n" + loop_start

        loop_init, prog_var = _loop_vars_to_python(self, tmp_context, indent)
        hash_object = hashlib.md5(str(self).encode())
        len_var = "len_" + hash_object.hexdigest()
        pos_var = "pos_" + hash_object.hexdigest()
        loop_init += indent_str + len_var + " = len(" + loop_vals + ")\n"
        loop_init += indent_str + pos_var + " = 0\n"
            
        save_vals = _updated_vars_to_python(self, context, indent)
        
        loop_body = ""
        loop_body += indent_str + " " * 4 + pos_var + " += 1\n"
        loop_body += indent_str + " " * 4 + "if (int(float(" + pos_var + ")/(" + len_var + " if " + len_var + " != 0 else 1)*100) == " + prog_var + "):\n"
        body_escaped = str(self).replace('"', '\\"').replace("\\n", " :: ")
        loop_body += indent_str + " " * 8 + "safe_print(str(int(float(" + pos_var + ")/(" + len_var + " if " + len_var + " != 0 else 1)*100)) + \"% done with loop " + body_escaped + "\")\n"
        loop_body += indent_str + " " * 8 + prog_var + " += 1\n"
        loop_body += to_python(self.statements, tmp_context, params=params, indent=indent+4, statements=True)
            
        python_code = loop_init + "\n" + \
                      loop_start + "\n" + \
                      loop_body + "\n" + \
                      save_vals + "\n"

        return python_code

    def eval(self, context, params=None):

        if (context.exit_func):
            return
        
        self.exited_with_goto = False
        context.loop_stack.append(None)
        context.loop_object_stack.append(self)
        my_loop_stack_pos = len(context.loop_stack) - 1
        

        container = eval_arg(self.container, context=context)
        try:
            container = context.get(self.container)
        except KeyError:
            pass
        except AssertionError:
            pass

        do_const_assignments(self.statements, context)

        if (_eval_python(self, context, params=params, add_boilerplate=True)):
            return
        
        if (not isinstance(container, list)):
            container = [container]
        try:
            for item_val in container:

                context.set(self.item, item_val)
                
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug('FOR EACH loop: %r = %r' % (self.item, context.get(self.item)))
                done = False
                context.goto_executed = False
                for s in self.statements:
                    if (log.getEffectiveLevel() == logging.DEBUG):
                        log.debug('FOR EACH loop eval statement: %r' % s)
                    if (not isinstance(s, VBA_Object)):
                        continue
                    s.eval(context=context)

                    if ((my_loop_stack_pos >= len(context.loop_stack)) or context.loop_stack[my_loop_stack_pos]):

                        self.exited_with_goto = (context.loop_stack[my_loop_stack_pos] == "GOTO")
                        if (log.getEffectiveLevel() == logging.DEBUG):
                            log.debug("FOR EACH loop: exited loop with 'Exit For'")
                        done = True
                        break

                    if (context.must_handle_error()):
                        done = True
                        break
                    context.clear_error()

                    if (context.goto_executed or s.exited_with_goto):
                        if (log.getEffectiveLevel() == logging.DEBUG):
                            log.debug("GOTO executed. Go to next loop iteration.")
                        break
                    
                if (done):
                    break

        except:

            pass
        
        if (len(context.loop_stack) > 0):
            context.loop_stack.pop()
        if (len(context.loop_object_stack) > 0):
            context.loop_object_stack.pop()
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('FOR EACH loop: end.')

        context.handle_error(params)
        
for_each_clause = CaselessKeyword("For").suppress() \
                  + CaselessKeyword("Each").suppress() \
                  + lex_identifier("item") \
                  + CaselessKeyword("In").suppress() \
                  + expression("container") \

real_simple_for_each_statement = for_each_clause('clause') + Suppress(EOS) + statement_block('statements') \
                                 + CaselessKeyword("Next").suppress() \
                                 + Optional(lex_identifier) \
                                 + FollowedBy(EOS)
real_simple_for_each_statement.setParseAction(For_Each_Statement)

bogus_simple_for_each_statement = for_each_clause('clause') + Suppress(EOS) + statement_block('statements') + ~CaselessKeyword("Next") + \
                                  CaselessKeyword("End") + (CaselessKeyword("Sub") | CaselessKeyword("Function"))
bogus_simple_for_each_statement.setParseAction(For_Each_Statement)



def _get_guard_variables(loop_obj, context):

    if (not hasattr(loop_obj.guard, "accept")):
        return {}
    
    var_visitor = var_in_expr_visitor()
    loop_obj.guard.accept(var_visitor)
    guard_var_names = var_visitor.variables

    r = {}
    for var in guard_var_names:
        try:
            r[var] = context.get(var)
        except:
            pass

    return r

class While_Statement(VBA_Object):

    def __init__(self, original_str, location, tokens):
        super(While_Statement, self).__init__(original_str, location, tokens)
        self.is_loop = True
        self.original_str = original_str[location:]
        self.loop_type = tokens.clause.type
        self.guard = tokens.clause.guard
        self.body = tokens[2]
        self._local_calls = None
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed %r as %s' % (self, self.__class__.__name__))

    def __repr__(self):
        r = "Do " + str(self.loop_type) + " " + str(self.guard) + "\\n"
        r += str(self.body) + "\\nLoop"
        return r

    def to_python(self, context, params=None, indent=0):

        indent_str = " " * indent

        until_pre = ""
        until_post = ""
        if (self.loop_type.lower() == "until"):
            until_pre = "not ("
            until_post = ")"
        
        loop_start = indent_str + "exit_all_loops = False\n"
        loop_start += indent_str + "max_errors = " + str(VBA_Object.loop_upper_bound/10000) + "\n"
        loop_start += indent_str + "while " + until_pre + to_python(self.guard, context) + until_post + ":\n"
        loop_start += indent_str + " " * 4 + "if exit_all_loops:\n"
        loop_start += indent_str + " " * 8 + "break\n"
        loop_start = indent_str + "# Start emulated loop.\n" + loop_start

        loop_init, prog_var = _loop_vars_to_python(self, context, indent)
            
        save_vals = _updated_vars_to_python(self, context, indent)
        
        loop_str = str(self).replace('"', '\\"').replace("\\n", " :: ")
        if (len(loop_str) > 100):
            loop_str = loop_str[:100] + " ..."
        loop_body = ""
        loop_body += indent_str + " " * 4 + "if (" + prog_var + " % 100) == 0:\n"
        loop_body += indent_str + " " * 8 + "safe_print(\"Done \" + str(" + prog_var + ") + \" iterations of While loop '" + loop_str + "'\")\n"
        loop_body += indent_str + " " * 4 + prog_var + " += 1\n"
        loop_body += indent_str + " " * 4 + "if (" + prog_var + " > " + str(VBA_Object.loop_upper_bound) + ") or " + \
                     "(vm_context.get_general_errors() > max_errors):\n"
        loop_body += indent_str + " " * 8 + "raise ValueError('Infinite Loop')\n"
        loop_body += to_python(self.body, context, params=params, indent=indent+4, statements=True)
            
        python_code = loop_init + "\n" + \
                      loop_start + "\n" + \
                      loop_body + "\n" + \
                      save_vals + "\n"

        return python_code
        
    def _eval_guard(self, curr_counter, final_val, comp_op):
        if (comp_op == "<="):
            return (curr_counter <= final_val)
        if (comp_op == "<"):
            return (curr_counter < final_val)
        if (comp_op == ">="):
            return (curr_counter >= final_val)
        if (comp_op == ">"):
            return (curr_counter > final_val)
        if ((comp_op == "==") or (comp_op == "=")):
            return (curr_counter == final_val)
        log.error("Loop guard operator '" + str(comp_op) + " cannot be emulated.")
        return False
        
    def _handle_simple_loop(self, context):


        if ((len(self.body) != 1) and (len(self.body) != 2)):
            return False

        if ("sleep(" in str(self.body[0]).lower()):
            return True

        if (("execute(" in str(self.body[0]).lower()) or
            ("executeglobal(" in str(self.body[0]).lower())):

            for s in self.body:
                if (not isinstance(s, VBA_Object)):
                    continue
                s.eval(context=context)

            return True
        
        loop_counter = str(self.guard).strip()
        m = re.match(r"(\w+)\s*([<>=]{1,2})\s*(\w+)", loop_counter)
        if (m is None):
            return False

        loop_counter = m.group(1)
        comp_op = m.group(2)
        upper_bound = m.group(3)
        
        var_inc = loop_counter + " = " + loop_counter
        body = str(self.body[0]).replace("Let ", "").replace("(", "").replace(")", "").strip()
        if_block = None
        if_val = None
        if (not body.startswith(var_inc)):

            if (len(self.body) != 2):
                return False
            
            body = None
            for s in self.body:

                tmp = str(s).replace("Let ", "").replace("(", "").replace(")", "").strip()
                if (tmp.startswith(var_inc)):
                    body = tmp
                    continue

                if (isinstance(s, If_Statement)):

                    if_guard = s.pieces[0]["guard"]
                    if_guard_str = str(if_guard).strip()
                    if (if_guard_str.startswith(loop_counter + " = ")):

                        if (("Else " in str(s)) or ("ElseIf " in str(s))):
                            return False
                        
                        if_block = s.pieces[0]["body"]
                        
                        try:
                            start = if_guard_str.rindex("=") + 1
                            tmp = if_guard_str[start:].strip()
                            if_val = int(str(tmp))
                        except ValueError:
                            return False

            if (if_block is None):
                body = None
                        
        if (body is None):
            return False
                        
        if (" " not in body):
            return False
        body = body.replace(var_inc, "").strip()
        op = body[:body.index(" ")]
        if (op not in ["+", "-", "*"]):
            return False
        num = body[body.index(" ") + 1:]
        try:
            num = int(num)
        except:
            return False

        curr_counter = coerce_to_int(eval_arg(loop_counter, context=context, treat_as_var_name=True))
        final_val = eval_arg(upper_bound, context=context, treat_as_var_name=True)
        try:
            final_val = int(final_val)
        except:
            return False
        
        if ((num == 1) and (op == "+")):
            if (comp_op == "<="):
                curr_counter = final_val + 1
            if (comp_op == "<"):
                curr_counter = final_val

        running = self._eval_guard(curr_counter, final_val, comp_op)
        if (self.loop_type.lower() == "until"):
            running = (not running)
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Short circuiting loop evaluation: Guard: " + str(self.guard))
            log.debug("Short circuiting loop evaluation: Body: " + str(self.body))
        while (running):
            
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("Short circuiting loop evaluation: Guard: " + str(self.guard))
                log.debug("Short circuiting loop evaluation: Test: " + str(curr_counter) + " " + comp_op + " " + str(final_val))
            if (op == "+"):
                curr_counter += num
            if (op == "-"):
                curr_counter -= num
            if (op == "*"):
                curr_counter *= num

            running = self._eval_guard(curr_counter, final_val, comp_op)
            if (self.loop_type.lower() == "until"):
                running = (not running)

        context.set(loop_counter, curr_counter)

        if (if_block is not None):
            for stmt in if_block:
                if (not isinstance(stmt, VBA_Object)):
                    continue
                if (hasattr(stmt, "eval")):
                    stmt.eval(context)
        
        return True

    def _has_local_calls(self, context):

        if (self._local_calls is not None):
            return self._local_calls

        if (not hasattr(self.body, "accept")):
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("Loop body has no accept() method.")
            return True

        call_visitor = function_call_visitor()
        for s in self.body:
            if (not isinstance(s, VBA_Object)):
                continue
            s.accept(call_visitor)

        for func in call_visitor.called_funcs:
            if (func not in context.external_funcs):
                self._local_calls = True
                return True

        self._local_calls = False
        return False
            
    def _no_state_change(self, prev_context, context):

        if ((prev_context is None) or (context is None)):
            return False
        
        if (self._has_local_calls(context)):
            return False

        guard_vars = _get_guard_variables(self, context)
        prev_context = Context(context=prev_context, _locals=prev_context.locals, copy_globals=True).delete("now").delete("application.username")
        for gvar in list(guard_vars.keys()):
            prev_context = prev_context.delete(gvar)
        curr_context = Context(context=context, _locals=context.locals, copy_globals=True).delete("now").delete("application.username")
        for gvar in list(guard_vars.keys()):
            curr_context = curr_context.delete(gvar)
        
        r = (prev_context == curr_context)
        return r

    def _has_constant_loop_guard(self, context):

        var_visitor = var_in_expr_visitor()
        self.guard.accept(var_visitor)
        if (len(var_visitor.variables) > 0):
            return None

        
        empty_context = Context()
        eval_guard_empty = str(eval_arg(self.guard, empty_context)).strip()
        if (eval_guard_empty == "True"):
            return True
        if (eval_guard_empty == "False"):
            return False
        return None
        
    def eval(self, context, params=None):

        if (context.exit_func):
            return
        
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('WHILE loop: start: ' + str(self))

        self.exited_with_goto = False
        if (len(self.body) == 0):

            if (hasattr(self.guard, "eval")):
                self.guard.eval(context)
            
            log.info("WHILE loop: empty body. Skipping.")
            return

        do_const_assignments(self.body, context)

        new_loop = loop_transform.transform_loop(self)
        if (new_loop != self):

            log.warning("Emulating transformed loop...")
            return new_loop.eval(context, params=params)
        
        if (self._handle_simple_loop(context)):

            return

        init_guard_val = self._has_constant_loop_guard(context)
        max_loop_iters = VBA_Object.loop_upper_bound
        is_infinite_loop = False
        if (init_guard_val is not None):

            if (init_guard_val):
                log.warn("Found infinite loop w. constant loop guard. Limiting iterations.")
                max_loop_iters = 2
                is_infinite_loop = True

            else:
                log.warn("Found loop that never runs w. constant loop guard. Skipping.")
                return

        if ((not is_infinite_loop) and
            (_eval_python(self, context, add_boilerplate=True))):
            return
        
        context.loop_stack.append(None)
        context.loop_object_stack.append(self)
        my_loop_stack_pos = len(context.loop_stack) - 1
        
        if (".readyState" in str(self.guard)):
            log.info("Limiting # of iterations of a .readyState loop.")
            max_loop_iters = 5
            
        old_guard_vals = _get_guard_variables(self, context)

        num_no_change_body = 0
        prev_context = None
        
        num_iters = 0
        num_no_change = 0
        context.clear_general_errors()
        while (True):
            
            if ((num_iters < 10) or (num_no_change > 0)):

                if (self._no_state_change(prev_context, context)):
                    num_no_change_body += 1
                    if (num_no_change_body >= context.max_static_iters * 500):
                        log.warn("Possible useless While loop detected. Exiting loop.")
                        self.is_useless = True
                        break
                else:
                    num_no_change = 0
                prev_context = Context(context=context, _locals=context.locals, copy_globals=True)
            
            if (num_iters > max_loop_iters):
                log.error("Maximum loop iterations exceeded. Breaking loop.")
                break
            num_iters += 1

            if (context.get_general_errors() > (max_loop_iters/10000)):
                log.error("Loop is generating too many errors. Breaking loop.")
                break
                
            guard_val = eval_arg(self.guard, context)
            if (self.loop_type.lower() == "until"):
                guard_val = (not guard_val)
            if (not guard_val):
                break
            
            done = False
            context.goto_executed = False
            for s in self.body:
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug('WHILE loop eval statement: %r' % s)
                if (not isinstance(s, VBA_Object)):
                    continue
                s.eval(context=context)

                if ((my_loop_stack_pos >= len(context.loop_stack)) or context.loop_stack[my_loop_stack_pos]):

                    self.exited_with_goto = (context.loop_stack[my_loop_stack_pos] == "GOTO")
                    if (log.getEffectiveLevel() == logging.DEBUG):
                        log.debug("WHILE loop: exited loop with 'Exit For'")
                    done = True
                    break

                if (context.must_handle_error()):
                    done = True
                    break
                context.clear_error()

                if (context.goto_executed or s.exited_with_goto):
                    if (log.getEffectiveLevel() == logging.DEBUG):
                        log.debug("GOTO executed. Go to next loop iteration.")
                    break
                
            if (done):
                break

            curr_guard_vals = _get_guard_variables(self, context)
            if (curr_guard_vals == old_guard_vals):
                num_no_change += 1
                if (num_no_change >= context.max_static_iters):
                    log.warn("Possible infinite While loop detected. Exiting loop.")
                    break
            else:
                num_no_change = 0

        if (len(context.loop_stack) > 0):
            context.loop_stack.pop()
        if (len(context.loop_object_stack) > 0):
            context.loop_object_stack.pop()
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('WHILE loop: end.')

        context.handle_error(params)
        
while_type = CaselessKeyword("While") | CaselessKeyword("Until")
        
while_clause = Optional(CaselessKeyword("Do").suppress()) + while_type("type") + boolean_expression("guard")

simple_while_statement = while_clause("clause") + Suppress(EOS) + Group(statement_block('body')) \
                       + (CaselessKeyword("Loop").suppress() |
                          CaselessKeyword("Wend").suppress() |
                          (CaselessKeyword("End").suppress() + CaselessKeyword("While").suppress()))

simple_while_statement.setParseAction(While_Statement)


class Do_Statement(VBA_Object):
    def __init__(self, original_str, location, tokens):
        super(Do_Statement, self).__init__(original_str, location, tokens)
        self.is_loop = True
        self.loop_type = tokens.type
        self.guard = tokens.guard
        if (self.guard is None):
            self.guard = True
        self.body = tokens[0]
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed %r as %s' % (self, self.__class__.__name__))

    def __repr__(self):
        r = "Do\\n" + str(self.body) + "\\n"
        r += "Loop " + str(self.loop_type) + " " + str(self.guard)
        return r

    def to_python(self, context, params=None, indent=0):

        indent_str = " " * indent

        loop_start = indent_str + "exit_all_loops = False\n"
        loop_start += indent_str + "max_errors = " + str(VBA_Object.loop_upper_bound/10000) + "\n"
        loop_start += indent_str + "while (True):\n"
        loop_start += indent_str + " " * 4 + "if exit_all_loops:\n"
        loop_start += indent_str + " " * 8 + "break\n"
        loop_start = indent_str + "# Start emulated loop.\n" + loop_start

        loop_init, prog_var = _loop_vars_to_python(self, context, indent)
            
        save_vals = _updated_vars_to_python(self, context, indent)
        
        loop_str = str(self).replace('"', '\\"').replace("\\n", " :: ")
        if (len(loop_str) > 100):
            loop_str = loop_str[:100] + " ..."
        loop_body = ""
        loop_body += indent_str + " " * 4 + "if (" + prog_var + " % 100) == 0:\n"
        loop_body += indent_str + " " * 8 + "safe_print(\"Done \" + str(" + prog_var + ") + \" iterations of Do While loop '" + loop_str + "'\")\n"
        loop_body += indent_str + " " * 4 + prog_var + " += 1\n"
        loop_body += indent_str + " " * 4 + "if (" + prog_var + " > " + str(VBA_Object.loop_upper_bound) + ") or " + \
                     "(vm_context.get_general_errors() > max_errors):\n"
        loop_body += indent_str + " " * 8 + "raise ValueError('Infinite Loop')\n"
        loop_body += to_python(self.body, context, params=params, indent=indent+4, statements=True)

        if (self.loop_type.lower() == "until"):
            loop_body += indent_str + " " * 4 + "if (" + to_python(self.guard, context) + "):\n"
        else:
            loop_body += indent_str + " " * 4 + "if (not (" + to_python(self.guard, context) + ")):\n"
        loop_body += indent_str + " " * 8 + "break\n"
        
        python_code = loop_init + "\n" + \
                      loop_start + "\n" + \
                      loop_body + "\n" + \
                      save_vals + "\n"

        return python_code
    
    def eval(self, context, params=None):

        if (context.exit_func):
            return
        
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('DO loop: start: ' + str(self))

        self.exited_with_goto = True
        if (len(self.body) == 0):
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("DO loop: empty body. Skipping.")
            return
        
        do_const_assignments(self.body, context)
        
        max_loop_iters = VBA_Object.loop_upper_bound
        if (".readyState" in str(self.guard)):
            log.info("Limiting # of iterations of a .readyState loop.")
            max_loop_iters = 5

        if (_eval_python(self, context, params=params, add_boilerplate=True)):
            return

        context.loop_stack.append(None)
        context.loop_object_stack.append(self)
        my_loop_stack_pos = len(context.loop_stack) - 1
        
        old_guard_vals = _get_guard_variables(self, context)
            
        num_iters = 0
        num_no_change = 0
        while (True):

            if (num_iters > max_loop_iters):
                log.error("Maximum loop iterations exceeded. Breaking loop.")
                break
            num_iters += 1
            
            done = False
            context.goto_executed = False
            for s in self.body:
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug('DO loop eval statement: %r' % s)
                if (not isinstance(s, VBA_Object)):
                    continue
                s.eval(context=context)

                if ((my_loop_stack_pos >= len(context.loop_stack)) or context.loop_stack[my_loop_stack_pos]):

                    self.exited_with_goto = (context.loop_stack[my_loop_stack_pos] == "GOTO")
                    if (log.getEffectiveLevel() == logging.DEBUG):
                        log.debug("Do loop: exited loop with 'Exit For'")
                    done = True
                    break

                if (context.must_handle_error()):
                    done = True
                    break
                context.clear_error()

                if (context.goto_executed or s.exited_with_goto):
                    if (log.getEffectiveLevel() == logging.DEBUG):
                        log.debug("GOTO executed. Go to next loop iteration.")
                    break
                
            if (done):
                break

            guard_val = eval_arg(self.guard, context)
            if (self.loop_type.lower() == "until"):
                guard_val = (not guard_val)
            if (not guard_val):
                break

            curr_guard_vals = _get_guard_variables(self, context)
            if (curr_guard_vals == old_guard_vals):
                num_no_change += 1
                if (num_no_change >= context.max_static_iters):
                    log.warn("Possible infinite While loop detected. Exiting loop.")
                    break
            else:
                num_no_change = 0
            
        if (len(context.loop_stack) > 0):
            context.loop_stack.pop()
        if (len(context.loop_object_stack) > 0):
            context.loop_object_stack.pop()
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('DO loop: end.')

        context.handle_error(params)
        
simple_do_statement = Suppress(CaselessKeyword("Do")) + Suppress(EOS) + \
                      Group(statement_block('body')) + \
                      Suppress(CaselessKeyword("Loop")) + Optional(while_type("type") + boolean_expression("guard"))

simple_do_statement.setParseAction(Do_Statement)



class Select_Statement(VBA_Object):
    def __init__(self, original_str, location, tokens):
        super(Select_Statement, self).__init__(original_str, location, tokens)
        self.select_val = tokens.select_val
        self.cases = tokens.cases
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed %r as %s' % (self, self.__class__.__name__))

    def __repr__(self):
        r = ""
        r += str(self.select_val)
        for case in self.cases:
            r += str(case)
        r += "End Select"
        return r

    def _to_python_if(self, context, indent, case, first):

        select_val_str = to_python(self.select_val, context)

        case.case_val.var_to_check = select_val_str
        case_guard_str = to_python(case.case_val, context)
        
        indent_str = " " * indent
        flow_str = "if "
        if (not first):
            flow_str = "elif "

        r = indent_str + flow_str + "(" + case_guard_str + "):\n"
            
        if ('in ["Else"]' in case_guard_str):
            r = indent_str + "else:\n"

        r += to_python(case.body, context, indent=indent+4, statements=True)
            
        return r
            
    def to_python(self, context, params=None, indent=0):
        r = ""
        first = True
        for case in self.cases:
            r += self._to_python_if(context, indent, case, first)
            first = False
        return r
    
    def eval(self, context, params=None):

        if (context.exit_func):
            return
        
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("eval select: " + str(self))
        if (not isinstance(self.select_val, VBA_Object)):
            return
        select_guard_val = self.select_val.eval(context, params)

        for case in self.cases:

            case_guard = case.case_val

            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("eval select: checking '" + str(select_guard_val) + " == " + str(case_guard) + "'")
            if (case_guard.eval(context, [select_guard_val])):

                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("eval select: take case " + str(case))
                for statement in case.body:

                    if (not isinstance(statement, VBA_Object)):
                        continue
                    statement.eval(context, params)

                    if (context.must_handle_error()):
                        break
                    context.clear_error()

                    if (context.goto_executed or statement.exited_with_goto):
                        if (log.getEffectiveLevel() == logging.DEBUG):
                            log.debug("GOTO executed. Break out of Select.")
                        break

                context.handle_error(params)
                    
                break

class Select_Clause(VBA_Object):
    def __init__(self, original_str, location, tokens):
        super(Select_Clause, self).__init__(original_str, location, tokens)
        self.select_val = tokens.select_val
        try:
            self.select_val = tokens.select_val[0]
        except TypeError:
            pass
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed %r as %s' % (self, self.__class__.__name__))

    def __repr__(self):
        r = ""
        r += "Select Case " + str(self.select_val) + "\\n " 
        return r

    def to_python(self, context, params=None, indent=0):
        return to_python(self.select_val, context)
    
    def eval(self, context, params=None):
        if (context.exit_func):
            return
        if (hasattr(self.select_val, "eval")):
            return self.select_val.eval(context, params)
        else:
            return self.select_val

class Case_Clause_Atomic(VBA_Object):
    def __init__(self, original_str, location, tokens):
        super(Case_Clause_Atomic, self).__init__(original_str, location, tokens)

        if (tokens[0] == "Else"):
            self.case_val = ["Else"]

        self.test_range = ((tokens.lbound != "") and (tokens.ubound != ""))
        if (self.test_range):
            self.case_val = [tokens.lbound, tokens.ubound]
        else:
            self.case_val = tokens
            
        if ((not self.test_range) and
            (tokens.case_val is not None) and
            (tokens.case_val != "")):
            self.case_val = tokens
        self.test_set = (not self.test_range) and (len(self.case_val) > 1)

        self.is_else = False
        for v in self.case_val:
            if (str(v).lower() == "else"):
                self.is_else = True
                break
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed %r as %s' % (self, self.__class__.__name__))

    def __repr__(self):
        r = ""
        if (self.test_range):
            r += str(self.case_val[0]) + " To " + str(self.case_val[1])
        elif (self.test_set):
            first = True
            for val in self.case_val:
                if (not first):
                    r += ", "
                first = False
                r += str(val)
        else:
            r += str(self.case_val[0])
        return r

    def to_python(self, context, params=None, indent=0):

        r = ""
        if (self.test_range):
            r += "range(" + to_python(self.case_val[0], context) + ", " + to_python(self.case_val[1], context) + " + 1)"
        elif (self.test_set):
            r += "["
            first = True
            for val in self.case_val:
                if (not first):
                    r += ", "
                first = False
                r += to_python(val, context)
            r += "]"
        else:
            r += "[" + to_python(self.case_val[0], context) + "]"
        return r
        
    def eval(self, context, params=None):

        if (context.exit_func):
            return
        
        test_val = params[0]

        if (self.is_else):
            return True
        
        if (self.test_range):

            start = None
            end = None
            try:
                start = utils.int_convert(eval_arg(self.case_val[0], context))
                end = utils.int_convert(eval_arg(self.case_val[1], context)) + 1
            except:
                return False                

            return (test_val in range(start, end))

        if (self.test_set):

            expected_vals = set()
            for val in self.case_val:
                try:
                    expected_vals.add(eval_arg(val, context))
                except:
                    return False

            return (test_val in expected_vals)

        expected_val = eval_arg(self.case_val[0], context)
        if (isinstance(test_val, int) and isinstance(expected_val, float)):
            test_val = 0.0 + test_val
        if (isinstance(test_val, float) and isinstance(expected_val, int)):
            expected_val = 0.0 + expected_val
        test_str = str(test_val)
        expected_str = str(expected_val)
        if (((test_str == "NULL") and (expected_str == "")) or
            ((expected_str == "NULL") and (test_str == ""))):
            return True
        return (test_str == expected_str)

class Case_Clause(VBA_Object):

    def __init__(self, original_str, location, tokens):
        super(Case_Clause, self).__init__(original_str, location, tokens)
        self.clauses = []
        for clause in tokens:
            self.clauses.append(clause)
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed %r as %s' % (self, self.__class__.__name__))

    def __repr__(self):
        r = "Case "
        first = True
        for clause in self.clauses:
            if (not first):
                r += ", "
            first = False
            r += str(clause)
        return r

    def to_python(self, context, params=None, indent=0):
        r = ""
        first = True
        for clause in self.clauses:
            if (not first):
                r += " or "
            first = False
            curr_str = to_python(clause, context)
            if ((not curr_str.startswith("range(")) and
                (not curr_str.startswith("["))):
                curr_str = "[" + curr_str + "]"
            r += self.var_to_check + " in " + curr_str
        return r
        
    def eval(self, context, params=None):

        if (context.exit_func):
            return
        
        for clause in self.clauses:
            guard_val = clause.eval(context, params=params)
            if (guard_val):
                return True
        return False
    
class Select_Case(VBA_Object):
    def __init__(self, original_str, location, tokens):
        super(Select_Case, self).__init__(original_str, location, tokens)
        self.case_val = tokens.case_val
        self.body = tokens.body
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed %r as %s' % (self, self.__class__.__name__))

    def __repr__(self):
        r = ""
        r += str(self.case_val) + " " + str(self.body)
        return r

    def eval(self, context, params=None):
        if (context.exit_func):
            return
    
select_clause = CaselessKeyword("Select").suppress() + CaselessKeyword("Case").suppress() \
                + (expression("select_val") ^ boolean_expression("select_val"))
select_clause.setParseAction(Select_Clause)

case_clause_atomic = ((expression("lbound") + CaselessKeyword("To").suppress() + expression("ubound")) | \
                      (CaselessKeyword("Else")) | \
                      (any_expression("case_val")))
case_clause_atomic.setParseAction(Case_Clause_Atomic)

case_clause = CaselessKeyword("Case").suppress() + \
              Suppress(Optional(CaselessKeyword("Is") + (Literal('=') ^ Literal('<') ^ Literal('>') ^ Literal('<=') ^ Literal('>=') ^ Literal('<>')))) + \
              case_clause_atomic + ZeroOrMore(Suppress(",") + case_clause_atomic)
case_clause.setParseAction(Case_Clause)

simple_statements_line = Forward()
select_case = case_clause("case_val") + \
              Optional((NotAny(EOS) + Group(simple_statements_line)("body")) ^ \
                       (Suppress(EOS) + Group(statement_block_not_empty('statements'))("body")))
select_case.setParseAction(Select_Case)

simple_select_statement = select_clause("select_val") + Suppress(EOS) + Group(ZeroOrMore(select_case + Suppress(Optional(EOS))))("cases") \
                          + Suppress(Optional(EOS)) + CaselessKeyword("End").suppress() + CaselessKeyword("Select").suppress()
simple_select_statement.setParseAction(Select_Statement)



class If_Statement(VBA_Object):

    def __init__(self, original_str, location, tokens):
        super(If_Statement, self).__init__(original_str, location, tokens)

        self.is_bogus = False
        if (isinstance(tokens, If_Statement)):
            self.pieces = tokens.pieces
            return
        if ((len(tokens) == 1) and (isinstance(tokens[0], If_Statement))):
            self.pieces = tokens[0].pieces
            return

        if ((len(tokens) == 1) and (isinstance(tokens[0], BoolExpr))):
            self.is_bogus = True
            return
        
        self.pieces = []
        for tok in tokens:
            if (len(tok) == 2):
                self.pieces.append({ 'guard' : tok[0], 'body' : tok[1]})
            elif (len(tok) == 1):
                self.pieces.append({ 'guard' : None, 'body' : tok[0]})
            else:
                log.error('If part %r has wrong # elements.' % str(tok))

        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed %r as %s' % (self, self.__class__.__name__))

    def get_children(self):

        if (self._children is not None):
            return self._children
        self._children = []
        if (self.is_bogus):
            return self._children
        for piece in self.pieces:

            if (isinstance(piece["body"], VBA_Object)):
                self._children.append(piece["body"])
            if ((isinstance(piece["body"], list)) or
                (isinstance(piece["body"], pyparsing.ParseResults))):
                for i in piece["body"]:
                    if (isinstance(i, VBA_Object)):
                        self._children.append(i)
            if (isinstance(piece["body"], dict)):
                for i in list(piece["body"].values()):
                    if (isinstance(i, VBA_Object)):
                        self._children.append(i)

            if (isinstance(piece["guard"], VBA_Object)):
                self._children.append(piece["guard"])
            if ((isinstance(piece["guard"], list)) or
                (isinstance(piece["guard"], pyparsing.ParseResults))):
                for i in piece["guard"]:
                    if (isinstance(i, VBA_Object)):
                        self._children.append(i)
            if (isinstance(piece["guard"], dict)):
                for i in list(piece["guard"].values()):
                    if (isinstance(i, VBA_Object)):
                        self._children.append(i)

        return self._children

    def __repr__(self):
        return self._to_str(False)

    def full_str(self):
        return self._to_str(True)
    
    def _to_str(self, full_str):
        if (self.is_bogus):
            return "BOGUS IF STATEMENT"
        r = ""
        first = True
        for piece in self.pieces:

            keyword = "If"
            if (not first):
                keyword = "ElseIf"
            if (piece["guard"] is None):
                keyword = "Else"
            first = False
            r += keyword + " "

            guard = ""
            keyword = ""
            if (piece["guard"] is not None):
                guard = None
                if (full_str):
                    guard = piece["guard"].full_str()
                else:
                    guard = piece["guard"].__repr__()
                    if (len(guard) > 5):
                        guard = guard[:6] + "..."
            r += guard + " "
            keyword = "Then "

            r += keyword
            body = None
            if (full_str):
                body = piece["body"].full_str()
            else:
                body = piece["body"].__repr__().replace("\n", "; ")
                if (len(body) > 25):
                    body = body[:26] + "..."
            r += body + " "

        if (full_str):
            log.debug("Guard: %s  Body: %s", guard, body)
            sys.exit(0)
        return r

    def to_python(self, context, params=None, indent=0):

        if (self.is_bogus):
            return ""

        r = ""
        first = True
        indent_str = " " * indent
        for piece in self.pieces:

            r += indent_str
            keyword = "if"
            if (not first):
                keyword = "elif"
            if (piece["guard"] is None):
                keyword = "else"
            first = False
            r += keyword + " "

            guard = ""
            keyword = ""
            if (piece["guard"] is not None):
                guard = to_python(piece["guard"], context)
            r += guard

            r += ":\n"
            body_str = to_python(piece["body"], context, indent=indent+4, statements=True)
            if (len(body_str.strip()) == 0):
                body_str = " " * (indent + 4) + "pass\n"
            r += body_str

        return r
    
    def eval(self, context, params=None):

        if (self.is_bogus):
            return
        
        if (context.exit_func):
            return
        
        for piece in self.pieces:

            guard = True
            if (piece["guard"] is not None):
                raw_guard = piece["guard"]
                log.info("If guard raw object: repr=%s type=%s", repr(raw_guard)[:200], type(raw_guard).__name__)

                guard = raw_guard.eval(context)
                log.info("If guard after .eval(): repr=%s type=%s", repr(guard)[:200], type(guard).__name__)

                while (isinstance(guard, VBA_Object) and hasattr(guard, "eval")):
                    guard = guard.eval(context)
                    log.info("If guard after re-eval VBA_Object: repr=%s type=%s", repr(guard)[:200], type(guard).__name__)

                if (isinstance(guard, str)):
                    if (context.contains(guard)):
                        log.info("If guard string matches variable '%s', resolving from context", guard)
                        guard = context.get(guard)
                        log.info("If guard resolved from context: repr=%s type=%s", repr(guard)[:200], type(guard).__name__)

                guard = _vba_truthy(guard)
                log.info("If guard final truth value: %s", guard)

            if (guard):

                for stmt in piece["body"]:
                    if (not isinstance(stmt, VBA_Object)):
                        continue
                    if (hasattr(stmt, "eval")):
                        stmt.eval(context)

                break

multi_line_if_statement = Group( CaselessKeyword("If").suppress() + boolean_expression + CaselessKeyword("Then").suppress() + Suppress(EOS) + \
                                 Group(statement_block)) + \
                                 ZeroOrMore(
                                     Group( CaselessKeyword("ElseIf").suppress() + boolean_expression + CaselessKeyword("Then").suppress() + Suppress(EOS) + \
                                            Group(statement_block))
                                 ) + \
                                 Optional(
                                     Group(CaselessKeyword("Else").suppress() + Group(simple_statements_line)) + Suppress(EOS) | \
                                     Group(CaselessKeyword("Else").suppress() + Suppress(EOS) + Group(statement_block))
                                 ) + \
                                 CaselessKeyword("End").suppress() + CaselessKeyword("If").suppress()
bad_if_statement = Group( CaselessKeyword("If").suppress() + boolean_expression + CaselessKeyword("Then").suppress() + Suppress(EOS) + \
                          Group(statement_block('statements'))) + \
                          ZeroOrMore(
                              Group( CaselessKeyword("ElseIf").suppress() + boolean_expression + CaselessKeyword("Then").suppress() + Suppress(EOS) + \
                                     Group(statement_block('statements')))
                          ) + \
                          Optional(
                              Group(CaselessKeyword("Else").suppress() + Suppress(EOS) + \
                                    Group(statement_block('statements')))
                          )

_single_line_if_statement = Group( CaselessKeyword("If").suppress() + boolean_expression + CaselessKeyword("Then").suppress() + \
                                   Group(simple_statements_line('statements')) )  + \
                                   ZeroOrMore(
                                       Group( CaselessKeyword("ElseIf").suppress() + boolean_expression + CaselessKeyword("Then").suppress() + \
                                              Group(simple_statements_line('statements')))
                                   ) + \
                                   Optional(
                                       (Group(CaselessKeyword("Else").suppress() + Group(simple_statements_line('statements'))) ^
                                        Group(CaselessKeyword("Else").suppress()))
                                   ) + Suppress(Optional(Optional(Literal(":")) + CaselessKeyword("End") + CaselessKeyword("If")))
single_line_if_statement = _single_line_if_statement
single_line_if_statement.setParseAction(If_Statement)

bogus_if_statement = CaselessKeyword("If").suppress() + boolean_expression + Optional(CaselessKeyword("Then")).suppress()

simple_if_statement = multi_line_if_statement ^ _single_line_if_statement ^ bogus_if_statement
simple_if_statement.setParseAction(If_Statement)


class If_Statement_Macro(If_Statement):

    def __init__(self, original_str, location, tokens):
        super(If_Statement_Macro, self).__init__(original_str, location, tokens)
        self.external_functions = {}
        for piece in self.pieces:
            for token in piece["body"]:
                if isinstance(token, External_Function):
                    if (log.getEffectiveLevel() == logging.DEBUG):
                        log.debug("saving VBA macro external func decl: %r" % token.name)
                    self.external_functions[token.name] = token

    def eval(self, context, params=None):

        if (context.exit_func):
            return
        
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("eval: " + str(self))
        then_part = self.pieces[0]
        for stmt in then_part["body"]:
            if (isinstance(stmt, VBA_Object)):
                stmt.eval(context)

simple_if_statement_macro = Group( CaselessKeyword("#If").suppress() + boolean_expression + CaselessKeyword("Then").suppress() + Suppress(EOS) + \
                                   Group(statement_block('statements'))) + \
                                   ZeroOrMore(
                                       Group( Suppress(CaselessKeyword("#ElseIf") | CaselessKeyword("ElseIf")) + \
                                              boolean_expression + CaselessKeyword("Then").suppress() + Suppress(EOS) + \
                                              Group(statement_block('statements')))
                                   ) + \
                                   Optional(
                                       Group(Suppress(CaselessKeyword("#Else") | CaselessKeyword("Else")) + Suppress(EOS) + \
                                             Group(statement_block('statements')))
                                   ) + \
                                   CaselessKeyword("#End If").suppress() + FollowedBy(EOS)

simple_if_statement_macro.setParseAction(If_Statement_Macro)


class Call_Statement(VBA_Object):

    log_funcs = ["CreateProcessA", "CreateProcessW", "CreateProcess", ".run", "CreateObject",
                 "Open", ".Open", "GetObject", "Create", ".Create", "Environ",
                 "CreateTextFile", ".CreateTextFile", ".Eval", "Run",
                 "SetExpandedStringValue", "WinExec", "FileCopy", "Load",
                 "FolderExists", "FileExists"]
    
    def __init__(self, original_str, location, tokens, name=None, params=None):
        super(Call_Statement, self).__init__(original_str, location, tokens)

        if ((name is not None) and (params is not None)):
            self.name = name
            self.params = params
            return

        self.name = tokens.name
        if (str(self.name).endswith("@")):
            self.name = str(self.name).replace("@", "")
        if (str(self.name).endswith("!")):
            self.name = str(self.name).replace("!", "")
        if (str(self.name).endswith("#")):
            self.name = str(self.name).replace("#", "")
        if (str(self.name).endswith("%")):
            self.name = str(self.name).replace("%", "")
        self.params = tokens.params

        if (str(self.name).lower() == "debug.print"):
            self.is_useless = True

        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed %r' % self)

    def __repr__(self):
        return 'Call_Statement: %s(%r)' % (self.name, self.params)

    def _to_python_handle_with_calls(self, context, indent):

        func_name = str(self.name).strip()
        if (not func_name.startswith(".")):
            return None

        if (len(context.with_prefix) == 0):
            return None

        tmp_var = SimpleNameExpression(None, None, None, name=str(context.with_prefix_raw))
        call_obj = Function_Call(None, None, None, old_call=self)
        call_obj.name = func_name[1:]
        full_expr = MemberAccessExpression(None, None, None, raw_fields=(tmp_var, [call_obj], []))
        
        r = to_python(full_expr, context, indent=indent)
        return r
    
    def to_python(self, context, params=None, indent=0):

        log.info(
            "CALL_STATEMENT start: name=%s raw_params_count=%d",
            self.name, len(self.params) if hasattr(self.params, '__len__') else '?',
        )
        if hasattr(self.params, '__iter__'):
            for _pi, _p in enumerate(self.params):
                log.info(
                    "  raw param[%d]: type=%s repr=%s",
                    _pi, type(_p).__name__, repr(_p)[:200],
                )

        dll_func_name = context.get_true_name(self.name)
        is_external = False
        if (dll_func_name is not None):
            is_external = True
            self.name = dll_func_name
        
        with_call_str = self._to_python_handle_with_calls(context, indent)
        if (with_call_str is not None):
            return with_call_str
        
        py_params = []
        old_bitwise = context.in_bitwise_expression
        context.in_bitwise_expression = True
        for p in self.params:
            py_params.append(to_python(p, context, params))
        context.in_bitwise_expression = old_bitwise

        indent_str = " " * indent
        if ((isinstance(self.name, VBA_Object)) and (len(self.params) == 0)):
            r = to_python(self.name, context, params)
            if (r.startswith(".")):
                r = r[1:]
            r = indent_str + r
            return r
            
        func_name = str(self.name)
        if ("." in func_name):
            func_name = func_name[func_name.index(".") + 1:]
        import vba_library
        is_internal = (func_name.lower() in vba_library.VBA_LIBRARY)
        if (is_internal or is_external):

            first = True
            args = "["
            for p in py_params:
                if (not first):
                    args += ", "
                first = False
                args += p

            if ((str(func_name) == "Execute") or
                (str(func_name) == "ExecuteGlobal") or
                (str(func_name) == "AddCode") or
                (str(func_name) == "AddFromString")):
                args += ", locals(), \"__JIT_EXEC__\""
            args += "]"

            r = None
            if is_internal:
                r = indent_str + "core.vba_library.run_function(\"" + str(func_name) + "\", vm_context, " + args + ")"
            else:
                r = indent_str + "core.vba_library.run_external_function(\"" + str(func_name) + "\", vm_context, " + args + ",\"\")"
            return r
                
        r = func_name + "("
        first = True
        for p in py_params:
            if (not first):
                r += ", "
            first = False
            r += p
        r += ")"
        if (r.startswith(".")):
            r = r[1:]
        r = indent_str + r
        
        return r

    def _handle_as_member_access(self, context):

        func_name = str(self.name).strip()
        if (("." not in func_name) or (func_name.startswith("."))):
            return None
        short_func_name = func_name[func_name.rindex(".") + 1:]
        
        memb_funcs = set(["AddItem", "Append_3"])
        if (short_func_name not in memb_funcs):
            return None

        func_call_str = func_name + "("
        first = True
        for p in self.params:
            if (not first):
                func_call_str += ", "
            first = False
            p_eval = eval_arg(p, context=context)
            if isinstance(p_eval, str):
                p_eval = p_eval.replace('"', '""').replace("\n", "\\n").replace("\r", "\\r")
                p_eval = '"' + p_eval + '"'
            func_call_str += str(p_eval)
        func_call_str += ")"
        try:
            memb_exp = member_access_expression.parseString(func_call_str, parseAll=True)[0]

            return memb_exp.eval(context)
        except ParseException as e:

            return None
        
    def _handle_with_calls(self, context):

        as_member_access = self._handle_as_member_access(context)
        if (as_member_access is not None):
            return as_member_access
        
        func_name = str(self.name).strip()
        if (not func_name.startswith(".")):
            return None

        if (len(context.with_prefix) == 0):
            return None

        call_obj = Function_Call(None, None, None, old_call=self)
        call_obj.name = func_name[1:]
        full_expr = MemberAccessExpression(None, None, None, raw_fields=(context.with_prefix, [call_obj], []))

        r = eval_arg(full_expr, context)
        return r
        
    def eval(self, context, params=None):

        if (context.exit_func):
            log.info("Exit function previously called. Not evaluating '" + str(self) + "'")
            return

        import vba_library
        vba_library.var_names = self.params
        
        dll_func_name = context.get_true_name(self.name)
        is_external = False
        if (dll_func_name is not None):
            is_external = True
            self.name = dll_func_name

        if isinstance(self.name, MemberAccessExpression):

            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("Call of member access expression " + str(self.name))
            r = self.name.eval(context, self.params)
            return r


        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Call: eval params: " + str(self.params))
        call_params = eval_args(self.params, context=context)
        str_params = repr(call_params)
        if (len(str_params) > 80):
            str_params = str_params[:80] + "..."

        if (context.have_error()):
            log.warn('Short circuiting function call %s(%s) due to thrown VB error.' % (self.name, str_params))
            return None

        if (str(self.name).strip().lower() in ("application.goto", "goto") and
            (len(call_params) > 0) and hasattr(context, "set_excel_selection")):
            context.set_excel_selection(call_params[0])
            return "NULL"

        if (not context.throttle_logging):
            log.info('Calling Procedure: %s(%r)' % (self.name, str_params))
        if (is_external):
            context.report_action("External Call", self.name + "(" + str(call_params) + ")", self.name, strip_null_bytes=True)
        if ((str(self.name).lower() in context._log_funcs) or
            (any(str(self.name).lower().endswith(func.lower()) for func in Function_Call.log_funcs))):
            context.report_action(self.name, call_params, 'Interesting Function Call', strip_null_bytes=True)

        r = self._handle_with_calls(context)
        if (r is not None):
            return r
        
        func_name = str(self.name).strip()
        if func_name.lower() == 'msgbox':
            context.report_action('Display Message', call_params, 'MsgBox', strip_null_bytes=True)
            return 1
        elif '.' in func_name:
            tmp_call_params = call_params
            if (func_name.endswith(".Write")):
                tmp_call_params = []
                for p in call_params:
                    if (isinstance(p, str)):
                        tmp_call_params.append(p.replace("\x00", ""))
                    else:
                        tmp_call_params.append(p)
            if ((func_name != "Debug.Print") and
                (not func_name.endswith("Add")) and
                (not func_name.endswith("Write")) and
                (len(tmp_call_params) > 0)):
                context.report_action('Object.Method Call', tmp_call_params, func_name, strip_null_bytes=True)

        try:

            if ("." in func_name):
                func_name = func_name[func_name.index(".") + 1:]
            
            s = context.get(func_name)
            if (s is None):
                raise KeyError("func not found")
            if (hasattr(s, "eval")):
                ret = s.eval(context=context, params=call_params)
                
                if (hasattr(s, "byref_params") and s.byref_params):
                    for byref_param_info in list(s.byref_params.keys()):
                        if (byref_param_info[1] < len(self.params)):
                            arg_var_name = str(self.params[byref_param_info[1]])
                            context.set(arg_var_name, s.byref_params[byref_param_info])

                context.exit_func = False

                return ret
            
        except KeyError:
            try:
                tmp_name = func_name.replace("$", "").replace("VBA.", "").replace("Math.", "").\
                           replace("[", "").replace("]", "").replace("'", "").replace('"', '')
                if ("." in tmp_name):
                    tmp_name = tmp_name[tmp_name.rindex(".") + 1:]
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("Looking for procedure %r" % tmp_name)
                s = context.get(tmp_name)
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("Found procedure " + tmp_name + " = " + str(s))
                if (s):
                    if (log.getEffectiveLevel() == logging.DEBUG):
                        log.debug("Found procedure. Running procedure " + tmp_name)
                    s.eval(context=context, params=call_params)
            except KeyError:

                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("Did not find procedure.")
                if ((func_name == "Application.Run") or (func_name == "Run")):

                    new_func = call_params[0]

                    new_params = call_params[1:]

                    if (log.getEffectiveLevel() == logging.DEBUG):
                        log.debug("Try indirect run of function '" + new_func + "'")
                    r = "NULL"
                    try:

                        s = context.get(new_func)
                        while (isinstance(s, str)):
                            s = context.get(s)
                            if (isinstance(s, procedures.Function) or
                                isinstance(s, procedures.Sub) or
                                isinstance(s, VbaLibraryFunc)):
                                s = s.eval(context=context, params=new_params)
                                r = s

                                context.exit_func = False
                            
                        return r

                    except KeyError:

                        context.increase_general_errors()
                        log.warning('Function %r not found' % func_name)
                        return r

                context.increase_general_errors()
                log.warning('Function %r not found' % func_name)
                    
            except Exception as e:
                traceback.print_exc(file=sys.stdout)
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("General error: " + str(e))
                return

call_params = (
    (Suppress('(') + Optional(expr_list('params')) + Suppress(')'))
    ^ (White(" \t") + Optional(Literal(",")) + expr_list('params'))
)
call_params_strict = (
    (Suppress('(') + Optional(expr_list_strict('params')) + Suppress(')'))
    ^ (White(" \t") + Optional(Literal(",")) + expr_list_strict('params'))
)
call_statement0 = NotAny(known_keywords_statement_start) + \
                  Optional(CaselessKeyword('Call').suppress()) + \
                  (member_access_expression('name'))  + \
                  Suppress(Optional(NotAny(White()) + '$') + \
                           Optional(NotAny(White()) + '#') + \
                           Optional(NotAny(White()) + '@') + \
                           Optional(NotAny(White()) + '%') + \
                           Optional(NotAny(White()) + '!')) + \
                           Optional(call_params_strict) + \
                           Suppress(Optional("," + CaselessKeyword("0")) + \
                                    Optional("," + (CaselessKeyword("true") | CaselessKeyword("false"))))
call_statement1 = NotAny(known_keywords_statement_start) + \
                  Optional(CaselessKeyword('Call').suppress()) + \
                  (TODO_identifier_or_object_attrib_loose('name')) + \
                  Suppress(Optional(NotAny(White()) + '$') + \
                           Optional(NotAny(White()) + '#') + \
                           Optional(NotAny(White()) + '@') + \
                           Optional(NotAny(White()) + '%') + \
                           Optional(NotAny(White()) + '!')) + \
                           Optional(call_params) + \
                           Suppress(Optional("," + CaselessKeyword("0")) + \
                                    Optional("," + (CaselessKeyword("true") | CaselessKeyword("false"))))
call_statement2 = NotAny(known_keywords_statement_start) + \
                  Optional(CaselessKeyword('Call').suppress()) + \
                  Combine(lex_identifier + OneOrMore(Literal(".") + lex_identifier)).setResultsName('name') + \
                  Suppress(Optional(NotAny(White()) + '$') + \
                           Optional(NotAny(White()) + '#') + \
                           Optional(NotAny(White()) + '@') + \
                           Optional(NotAny(White()) + '%') + \
                           Optional(NotAny(White()) + '!')) + \
                           Optional(call_params) + \
                           Suppress(Optional("," + CaselessKeyword("0")) + \
                                    Optional("," + (CaselessKeyword("true") | CaselessKeyword("false"))))
call_statement0.setParseAction(Call_Statement)
call_statement1.setParseAction(Call_Statement)
call_statement2.setParseAction(Call_Statement)

call_statement = (call_statement1 ^ call_statement0 ^ call_statement2)


class Exit_For_Statement(VBA_Object):
    def __init__(self, original_str, location, tokens):
        super(Exit_For_Statement, self).__init__(original_str, location, tokens)
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed %r' % self)

    def __repr__(self):
        return 'Exit For'

    def to_python(self, context, params=None, indent=0):
        return " " * indent + "break"

    def eval(self, context, params=None):
        if (context.exit_func):
            return
        if (len(context.loop_stack) > 0):
            context.loop_stack.pop()
        context.loop_stack.append("EXIT_FOR")

class Exit_While_Statement(Exit_For_Statement):
    def __repr__(self):
        return 'Exit Do'

exit_for_statement = CaselessKeyword('Exit').suppress() + CaselessKeyword('For').suppress()
exit_for_statement.setParseAction(Exit_For_Statement)

exit_while_statement = CaselessKeyword('Exit').suppress() + CaselessKeyword('Do').suppress()
exit_while_statement.setParseAction(Exit_While_Statement)

exit_loop_statement = exit_for_statement | exit_while_statement


class Exit_Function_Statement(VBA_Object):
    def __init__(self, original_str, location, tokens):
        super(Exit_Function_Statement, self).__init__(original_str, location, tokens)
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed %r' % self)

    def __repr__(self):
        return 'Exit Function'

    def to_python(self, context, params=None, indent=0):
        return " " * indent + "exit_all_loops = True"
    
    def eval(self, context, params=None):
        log.info("Explicit exit function invoked")
        context.exit_func = True

exit_func_statement = (CaselessKeyword('Exit').suppress() + CaselessKeyword('Function').suppress()) | \
                      (CaselessKeyword('Exit').suppress() + CaselessKeyword('Sub').suppress()) | \
                      (CaselessKeyword('Return').suppress())
exit_func_statement.setParseAction(Exit_Function_Statement)


class Redim_Statement(VBA_Object):
    def __init__(self, original_str, location, tokens):
        super(Redim_Statement, self).__init__(original_str, location, tokens)
        self.item = str(tokens.item)
        self.raw_item = tokens.item
        self.start = None
        if (hasattr(tokens, "start")):
            self.start = tokens.start
        self.end = None
        if (hasattr(tokens, "end")):
            self.end = tokens.end
        self.data_type = None
        if (hasattr(tokens, "data_type")):
            self.data_type = tokens.data_type
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed %r' % self)

    def __repr__(self):
        r = 'ReDim ' + str(self.item)
        if ((self.start is not None) and (self.end is not None)):
            r += "(" + str(self.start) + " To " + str(self.end) + ")"
        if (self.data_type is not None):
            r += " As " + self.data_type
        return r

    def to_python(self, context, params=None, indent=0):

        return "ERROR: ReDim JIT generation needs work."
        
        indent_str = " " * indent
        var_name = utils.fix_python_overlap(str(self.item))
        if (str(context.get_type(self.item)) == "Variant"):

            return indent_str + var_name + " = []"

        elif ((str(context.get_type(self.item)) == "Byte Array") or
              (self.data_type == "Byte") or
              (not context.contains(self.item))):

            if ((self.start is not None) and (self.end is not None)):

                start = None
                end = None
                try:

                    start = "int(" + to_python(self.start, context=context) + ")"
                    end = "int(" + to_python(self.end, context=context) + ")"

                    new_list = "[0] * (" + end + " - " + start + ")"
                    return indent_str + var_name + " = " + new_list

                except:
                    pass

        elif (isinstance(self.raw_item, Function_Call)):

            if (len(self.raw_item.params) > 0):

                new_size = "int(" + to_python(self.raw_item.params[0], context=context) + ")"

                new_list = "[0] * (" + new_size + ")"
                var_name = utils.fix_python_overlap(str(self.raw_item.name))
                return indent_str + var_name + " = " + new_list
                    
        return "ERROR: Cannot generate python code for '" + str(self) + "'"
        
    def eval(self, context, params=None):

        if (str(context.get_type(self.item)) == "Variant"):

            context.set(self.item, [])

        elif ((str(context.get_type(self.item)) == "Byte Array") or
              (not context.contains(self.item))):

            if ((self.start is not None) and (self.end is not None)):

                start = None
                end = None
                try:

                    start = int(eval_arg(self.start, context=context))
                    end = int(eval_arg(self.end, context=context))

                    new_list = [0] * (end - start)
                    context.set(self.item, new_list)

                except:
                    pass

        elif (isinstance(self.raw_item, Function_Call)):

            if (len(self.raw_item.params) > 0):

                new_size = eval_arg(self.raw_item.params[0], context=context)

                if (isinstance(new_size, int)):
                    new_list = [0] * (new_size)
                    var_name = self.raw_item.name
                    context.set(var_name, new_list)
                    
        return

redim_item = Optional(CaselessKeyword('Preserve')) + \
             expression('item') + \
             Optional('(' + expression('start') + CaselessKeyword('To') + expression('end') + ZeroOrMore("," + expression + CaselessKeyword('To') + expression) + ')') + \
             Optional(CaselessKeyword('As') + lex_identifier('data_type'))
redim_statement = CaselessKeyword('ReDim').suppress() + redim_item + ZeroOrMore("," + redim_item)

redim_statement.setParseAction(Redim_Statement)


class With_Statement(VBA_Object):

    def __init__(self, original_str, location, tokens):
        super(With_Statement, self).__init__(original_str, location, tokens)
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("tokens = " + str(tokens))
        self.body = tokens[-1]
        self.env = tokens.env
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed %r' % self)

    def __repr__(self):
        return 'With ' + str(self.env) + "\\n" + str(self.body) + " End With"

    def to_python(self, context, params=None, indent=0):

        with_dict = None
        if ((context.with_prefix_raw is not None) and
            (context.contains(str(context.with_prefix_raw)))):
            with_dict = context.get(str(context.with_prefix_raw))
            if (not isinstance(with_dict, dict)):
                with_dict = None
        if (with_dict is None):
            return "ERROR: Only doing JIT on Scripting.Dictionary With blocks."

        r = ""
        indent_str = " " * indent
        r += indent_str + "# With block: " + str(self).replace("\\n", "\\\\n")[:50] + "...\n"
        r += indent_str + "with_dict = " + str(with_dict) + "\n"
        
        r += to_python(self.body, context, indent=indent, statements=True)

        return r
                
    def eval(self, context, params=None):

        if (context.exit_func):
            return

        prefix_val = eval_arg(self.env, context)

        context.with_prefix_raw = self.env
        if (len(context.with_prefix) > 0):
            context.with_prefix += "." + str(self.env)
        else:
            context.with_prefix = str(prefix_val)
        if (context.with_prefix.startswith(".")):
            context.with_prefix = context.with_prefix[1:]

        do_const_assignments(self.body, context)
            
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("START WITH")
        try:
            tmp1 = iter(self.body)
        except TypeError:
            self.body = [self.body]
        for s in self.body:

            if (not isinstance(s, VBA_Object)):
                continue
            s.eval(context)

            if (context.must_handle_error()):
                break
            context.clear_error()

            if (context.goto_executed or s.exited_with_goto):
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("GOTO executed. Go to next loop iteration.")
                break
            
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("END WITH")
            
        context.with_prefix_raw = None
        if ("." not in context.with_prefix):
            context.with_prefix = ""
        else:
            end = context.with_prefix.rindex(".")
            context.with_prefix = context.with_prefix[:end]

        context.handle_error(params)
        
        return

with_statement = CaselessKeyword('With').suppress() + Optional(".") + (member_access_expression('env') ^ \
                                                                       ((lex_identifier('env') ^ function_call_limited('env')))) + Suppress(EOS) + \
                 Group(statement_block('body')) + \
                 CaselessKeyword('End').suppress() + CaselessKeyword('With').suppress()
with_statement.setParseAction(With_Statement)


class Goto_Statement(VBA_Object):
    def __init__(self, original_str, location, tokens):
        super(Goto_Statement, self).__init__(original_str, location, tokens)
        self.label = tokens.label
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed %r as Goto_Statement' % self)

    def __repr__(self):
        return 'Goto ' + str(self.label)

    def eval(self, context, params=None):

        if (context.exit_func):
            return
        
        if (self.label not in context.tagged_blocks):

            context.report_general_error("GOTO target " + str(self.label) + " is unknown.")
            return

        block = context.tagged_blocks[self.label]

        if (len(context.loop_object_stack) > 0):

            curr_loop = context.loop_object_stack[-1]
            jump_loop = None
            tag_block_txt = str(block).replace(" ", "").replace("\n", "")
            pos = len(context.loop_stack)
            for tmp_loop in context.loop_object_stack[::-1]:

                pos -= 1
                tmp_loop_txt = str(tmp_loop.body).replace(" ", "").replace("\n", "")
                if (tag_block_txt in tmp_loop_txt):
                    jump_loop = tmp_loop
                    break

            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("GOTO in loop.")
                log.debug("Jump to: " + tag_block_txt)
            if (jump_loop is None):

                context.loop_stack = ["GOTO"] * len(context.loop_stack)
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("Jumped out of all loops.")
                    log.debug(context.loop_stack)

            elif (jump_loop != curr_loop):

                tmp_stack = context.loop_stack
                context.loop_stack = context.loop_stack[:pos+1]
                context.loop_stack.extend(["GOTO"] * (len(tmp_stack) - (pos + 1)))
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("Jumped out of some loops.")
                    log.debug(context.loop_stack)
                
        if (not context.throttle_logging):
            log.info("GOTO " + str(self.label))
        block.eval(context, params)

        context.goto_executed = True

goto_statement = (CaselessKeyword('Goto').suppress() | CaselessKeyword('Gosub').suppress()) + (lex_identifier('label') | decimal_literal('label'))
goto_statement.setParseAction(Goto_Statement)


class Label_Statement(VBA_Object):
    def __init__(self, original_str, location, tokens):
        super(Label_Statement, self).__init__(original_str, location, tokens)
        self.label = tokens.label
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed %r as Label_Statement' % self)

    def __repr__(self):
        return str(self.label) + ':'

    def eval(self, context, params=None):
        return

label_statement <<= identifier('label') + Suppress(':')
label_statement.setParseAction(Label_Statement)


class On_Error_Statement(VBA_Object):
    def __init__(self, original_str, location, tokens):
        super(On_Error_Statement, self).__init__(original_str, location, tokens)
        self.tokens = tokens
        self.label = None
        if ((len(tokens) == 4) and (tokens[2].lower() == "goto")):
            self.label = str(tokens[3])
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed %r as On_Error_Statement' % self)

    def __repr__(self):
        return str(self.tokens)

    def to_python(self, context, params=None, indent=0):
        indent_str = " " * indent
        return indent_str + "# '" + str(self) + "' not emulated.\n" + \
            indent_str + "pass"
    
    def eval(self, context, params=None):

        if (self.label is not None):

            if (self.label in context.tagged_blocks):

                context.error_handler = context.tagged_blocks[self.label]
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("Setting On Error handler block to '" + self.label + "'.")

            else:
                log.warning("Cannot find error handler block '" + self.label + "'.")

        return

on_error_statement = CaselessKeyword('On') + Optional(Suppress(CaselessKeyword('Local'))) + (CaselessKeyword('Error') | lex_identifier) + \
                     ((CaselessKeyword('Goto') + (decimal_literal | lex_identifier)) |
                      (CaselessKeyword('Resume') + CaselessKeyword('Next')))

on_error_statement.setParseAction(On_Error_Statement)


resume_statement = CaselessKeyword('Resume') + Optional(lex_identifier)


class File_Open(VBA_Object):
    def __init__(self, original_str, location, tokens):
        super(File_Open, self).__init__(original_str, location, tokens)
        self.file_name = tokens.file_name
        self.file_id = tokens.file_id
        self.file_mode = None
        if (hasattr(tokens.type, "mode")):
            self.file_mode = tokens.type.mode
        self.file_access = None
        if (hasattr(tokens.type, "access")):
            self.file_access = tokens.type.access
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed %r as File_Open' % self)

    def __repr__(self):
        r = "Open " + str(self.file_name) + " For " + str(self.file_mode)
        if (self.file_access is not None):
            r += " Access " + str(self.file_access)
        r += " As " + str(self.file_id)
        return r

    def eval(self, context, params=None):

        if (context.exit_func):
            return
        
        name = self.file_name
        
        if (not str(name).lower().startswith("c:")):
            name = eval_arg(self.file_name, context=context)
            try:
                name = context.get(self.file_name)
            except KeyError:
                pass
            except AssertionError:
                pass

        file_id = ""
        if self.file_id:
            file_id = str(self.file_id)
            if not file_id.startswith('#'):

                try:
                    file_id = "#" + str(context.get(file_id))
                except KeyError:
                    file_id = "#" + str(context.get_num_open_files() + 1)
            context.set(file_id, name, force_global=True)

        context.report_action("OPEN", str(name), 'Open File', strip_null_bytes=True)
        context.open_file(name, file_id)


file_type = (
    Suppress(CaselessKeyword("For"))
    + (
        CaselessKeyword("Append")
        | CaselessKeyword("Binary")
        | CaselessKeyword("Input")
        | CaselessKeyword("Output")
        | CaselessKeyword("Input")
        | CaselessKeyword("Random")
    )("mode")
    + Suppress(Optional(CaselessKeyword("Lock")))
    + Optional(
        Optional(Suppress(CaselessKeyword("Access")))
        + (
            CaselessKeyword("Read Write")
            ^ CaselessKeyword("Read Shared")
            ^ CaselessKeyword("Read")
            ^ CaselessKeyword("Shared")
            ^ CaselessKeyword("Write")
        )("access")
    )
)

file_open_statement = (
    Suppress(CaselessKeyword("Open"))
    + expression("file_name")
    + Optional(
        file_type("type")
        + Suppress(CaselessKeyword("As"))
        + (file_pointer("file_id") | TODO_identifier_or_object_attrib("file_id") | file_pointer_loose("file_id"))
        + Suppress(Optional(CaselessKeyword("Len") + Literal("=") + expression))
    )
)
file_open_statement.setParseAction(File_Open)


class Print_Statement(VBA_Object):
    def __init__(self, original_str, location, tokens):
        super(Print_Statement, self).__init__(original_str, location, tokens)
        self.file_id = tokens.file_id
        self.value = tokens.value
        self.more_values = tokens.more_values
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed %r as Print_Statement' % self)

    def __repr__(self):
        r = "Print " + str(self.file_id) + ", " + str(self.value)
        return r

    def eval(self, context, params=None):

        if (context.exit_func):
            return
        
        file_id = eval_arg(self.file_id, context=context)
        try:
            file_id = context.get(self.file_id)
        except KeyError:
            pass
        except AssertionError:
            pass

        data = eval_arg(self.value, context=context)

        context.write_file(file_id, data)
        if (isinstance(data, str)):
            context.write_file(file_id, '\r\n')


print_statement = Suppress(CaselessKeyword("Print")) + file_pointer("file_id") + Suppress(Optional(",")) + expression("value") + \
                  ZeroOrMore(Suppress(Literal(';')) + expression)("more_values") + Suppress(Optional("," + lex_identifier))
print_statement.setParseAction(Print_Statement)


doevents_statement = Suppress(CaselessKeyword("DoEvents"))


simple_statement = (
    NotAny(Regex(r"End\s+Sub"))
    + (
        print_statement
        | dim_statement
        | option_statement
        | (
            prop_assign_statement
            ^ (let_statement | lset_statement | call_statement)
            ^ label_statement
            ^ expression
        )
        | exit_loop_statement
        | exit_func_statement
        | redim_statement
        | goto_statement
        | on_error_statement
        | file_open_statement
        | doevents_statement
        | rem_statement
        | resume_statement
        | global_variable_declaration
    )
)

simple_statement_restricted = (
    NotAny(Regex(r"End\s+Sub"))
    + (
        print_statement
        | dim_statement
        | option_statement
        | (
            prop_assign_statement
            ^ (let_statement | lset_statement | call_statement)
            ^ expression
        )
        | exit_loop_statement
        | exit_func_statement
        | redim_statement
        | goto_statement
        | on_error_statement
        | file_open_statement
        | doevents_statement
        | rem_statement
        | resume_statement
        | single_line_if_statement
    )
)

simple_statements_line <<= (
   (simple_statement_restricted + OneOrMore(Suppress(':') + simple_statement_restricted))
   ^ simple_statement_restricted
)

statements_line <<= (
    tagged_block
    ^ (Optional(statement_restricted + ZeroOrMore(Suppress(':') + statement_restricted)) + EOS.suppress())
)

statements_line_no_eos <<= (
    tagged_block
    ^ (Optional(statement_restricted + ZeroOrMore(Suppress(':') + statement_restricted)))
)


class External_Function(VBA_Object):

    file_count = 0
    def _createfile(self, params, context):

        fname = None
        if ((params[0] is not None) and (len(params[0]) > 0)):
            fname = "#" + str(params[0])
        else:
            External_Function.file_count += 1
            fname = "#SOME_FILE_" + str(External_Function.file_count)

        context.open_file(fname)

        return fname

    def _writefile(self, params, context):

        file_id = params[0]

        if (file_id not in context.open_files):
            context.report_general_error("File " + str(file_id) + " not open. Cannot write.")
            return 1
        
        data = params[1]
        if (not isinstance(data, int)):
            context.report_general_error("Cannot WriteFile() data that is not int.")
            return 0
        context.write_file(file_id, chr(data))
        return 0

    def _closehandle(self, params, context):

        file_id = params[0]
        context.close_file(file_id)
        return 0
    
    def __init__(self, original_str, location, tokens):
        super(External_Function, self).__init__(original_str, location, tokens)
        self.name = str(tokens.function_name)
        self.params = tokens.params
        self.lib_name = str(tokens.lib_info.lib_name)
        if isinstance(self.lib_name, basestring):
            self.lib_name = str(tokens.lib_name).strip('"').lower()
            if '.' not in self.lib_name:
                self.lib_name += '.dll'
        self.lib_name = str(self.lib_name)
        self.alias_name = str(tokens.lib_info.alias_name)
        if isinstance(self.alias_name, basestring):
            self.alias_name = self.alias_name.strip('"')
        if (len(self.alias_name.strip()) == 0):
            self.alias_name = self.name
        self.return_type = tokens.return_type
        self.vars = {}
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed %r' % self)

    def __repr__(self):
        return 'External Function %s (%s) from %s alias %s' % (self.name, self.params, self.lib_name, self.alias_name)

    def to_python(self, context, params=None, indent=0):

        lib_info = str(self.lib_name) + " / " + str(self.alias_name)

        indent_str = " " * indent
        r = indent_str + "# DLL Import: " + str(self.name) + " -> " + lib_info + "\n"
        return r
        
    def eval(self, context, params=None):

        if (context.exit_func):
            return


        if self.alias_name:
            function_name = self.alias_name
        else:
            function_name = self.name
        if (not context.throttle_logging):
            log.info('Evaluating external function %s(%r)' % (function_name, params))

        function_name = function_name.lower()
        if function_name.startswith('urldownloadtofile'):
            context.report_action('Download URL', params[1], 'External Function: urlmon.dll / URLDownloadToFile', strip_null_bytes=True)
            context.report_action('Write File', params[2], 'External Function: urlmon.dll / URLDownloadToFile', strip_null_bytes=True)
            return 0
        elif function_name.startswith('shellexecute'):
            cmd = None
            if (len(params) >= 4):
                cmd = str(params[2]) + " " + str(params[3])
            else:
                cmd = str(params[1]) + " " + str(params[2])
            context.report_action('Run Command', cmd, function_name, strip_null_bytes=True)
            return 0
        else:
            call_str = str(self.alias_name) + "(" + str(params) + ")"
            call_str = call_str.replace('\x00', "")
            context.report_action('External Call', call_str, str(self.lib_name) + " / " + str(self.alias_name))
        
        if (function_name.startswith('createfile')):
            return self._createfile(params, context)

        if (function_name.startswith('writefile')):
            return self._writefile(params, context)

        if (function_name.startswith('closehandle')):
            return self._closehandle(params, context)

        try:
            s = context.get_lib_func(function_name)
            if (s is None):
                raise KeyError("func not found")
            r = s.eval(context=context, params=params)
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("External function " + str(function_name) + " returns " + str(r))
            return r
        except KeyError:
            pass
        
        log.warning('Unknown external function %s from DLL %s' % (function_name, self.lib_name))

        return 1

function_type2 = CaselessKeyword('As').suppress() + lex_identifier('return_type') \
                 + Optional(Literal(".") + lex_identifier) \
                 + Optional(Literal('(') + Literal(')')).suppress()

public_private <<= Optional(CaselessKeyword('Public') | CaselessKeyword('Private') | CaselessKeyword('Global') | CaselessKeyword('Friend')) + \
                   Optional(CaselessKeyword('WithEvents'))

params_list_paren = Suppress('(') + Optional(parameters_list('params')) + Suppress(')')

lib_info = CaselessKeyword('Lib').suppress() + quoted_string('lib_name') \
           + Optional(CaselessKeyword('Alias') + quoted_string('alias_name'))

external_function <<= public_private + Suppress(CaselessKeyword('Declare') + Optional(CaselessKeyword('PtrSafe')) + \
                                                (CaselessKeyword('Function') | CaselessKeyword('Sub'))) + \
                                                lex_identifier('function_name') + lib_info('lib_info') + Optional(params_list_paren) + Optional(function_type2)
external_function.setParseAction(External_Function)


class TryCatch(VBA_Object):

    def __init__(self, original_str, location, tokens):
        super(TryCatch, self).__init__(original_str, location, tokens)
        tmp = TaggedBlock(original_str, location, None)
        tmp.block = tokens["try_block"]
        tmp.label = "try_block"
        self.try_block = tmp
        tmp = TaggedBlock(original_str, location, None)
        tmp.block = tokens["catch_block"]
        tmp.label = "catch_block"
        self.catch_block = tmp
        self.except_var = tokens["exception_var"]
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed %r' % self)

    def __repr__(self):
        return "Try::" + str(self.try_block) + "::Catch " + str(self.except_var) + " As Exception::" + str(self.catch_block) + "::End Try"

    def eval(self, context, params=None):

        old_handler = context.get_error_handler()
        context.error_handler = self.catch_block

        self.try_block.eval(context)

        context.error_handler = old_handler

try_catch = Suppress(CaselessKeyword('Try')) + Suppress(EOS) + statement_block('try_block') + \
            Suppress(CaselessKeyword('Catch')) + lex_identifier('exception_var') + Suppress(CaselessKeyword('As')) + Suppress(CaselessKeyword('Exception')) + \
            Suppress(EOS) + statement_block('catch_block') + Suppress(CaselessKeyword('##End')) + Suppress(CaselessKeyword('##Try'))
try_catch.setParseAction(TryCatch)


class NameStatement(VBA_Object):

    def __init__(self, original_str, location, tokens):
        super(NameStatement, self).__init__(original_str, location, tokens)
        self.old_name = tokens.old_name
        self.new_name = tokens.new_name
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed %r as NameStatement' % self)

    def __repr__(self):
        return "Name " + str(self.old_name) + " As " + str(self.new_name)

    def eval(self, context, params=None):

        old_name = eval_arg(self.old_name, context=context)
        new_name = eval_arg(self.new_name, context=context)

        context.report_action("File Rename", "Rename '" + old_name + "' to '" + new_name + "'", "File Rename", strip_null_bytes=True)
        
name_statement = CaselessKeyword('Name') + expression("old_name") + CaselessKeyword('As') + expression("new_name")
name_statement.setParseAction(NameStatement)


class Stop_Statement(VBA_Object):
    def __init__(self, original_str, location, tokens):
        super(Stop_Statement, self).__init__(original_str, location, tokens)
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed %r' % self)

    def __repr__(self):
        return 'Stop'

    def eval(self, context, params=None):
        pass

stop_statement = CaselessKeyword('Stop').suppress()
stop_statement.setParseAction(Stop_Statement)


class Line_Input_Statement(VBA_Object):
    def __init__(self, original_str, location, tokens):
        super(Line_Input_Statement, self).__init__(original_str, location, tokens)
        self.file_id = tokens.file_id
        self.var = tokens.var
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed %r as Line_Input_Statement' % self)

    def __repr__(self):
        return 'Line Input #' + str(self.file_id) + ", " + str(self.var)

    def eval(self, context, params=None):

        log.warn("'Line Input' statements not emulated. Treating '" + str(self) + "' as a NOOP.")

line_input_statement = CaselessKeyword('Line').suppress() + CaselessKeyword('Input').suppress() + \
                       Literal("#").suppress() + expression("file_id") + Literal(",") + \
                       expression("var")
line_input_statement.setParseAction(Line_Input_Statement)


def quick_parse_simple_call(tokens):
    text = str(tokens[0]).strip()
    r = []
    for i in text.split("\n"):
        i = i.strip()
        if (len(i) == 0):
            continue

        name = None
        params = None

        if ("(" in i):
            name = i[:i.index("(")].strip()
            params_str = i[i.index("(") + 1:].strip()
            if (params_str.endswith(")")):
                params_str = params_str[:-1]
                params = params_str.split(",")

        elif (" " in i):
            name = i[:i.index(" ")].strip()
            params_str = i[i.index(" ") + 1:].strip()
            params = params_str.split(",")

        if ((name is not None) and (params is not None)):

            tmp_params = []
            for p in params:

                param = None
                if (p.isdigit()):
                    tmp_params.append(int(p))

                elif (re.match(r"[_a-zA-Z][_a-zA-Z\d]*", p) is not None):
                    tmp_params.append(SimpleNameExpression(None, None, None, p))

                else:
                    tmp_params = None
                    break
            params = tmp_params

        if ((name is not None) and (params is not None)):
            r.append(Call_Statement(None, None, None, name=name, params=params))

        else:
            r.append(call_statement0.parseString(i, parseAll=True)[0])

    return r

simple_call_list = Regex(re.compile("(?:\w+\s*\(?(?:\w+\s*,\s*)*\s*\w+\)?\n){100,}"))
simple_call_list.setParseAction(quick_parse_simple_call)


class Orphaned_Marker(VBA_Object):
    def __init__(self, original_str, location, tokens):
        super(Orphaned_Marker, self).__init__(original_str, location, tokens)
        log.warning("Orphaned statement marker found.")

    def __repr__(self):
        return "' ORPHANED MARKER"

    def eval(self, context, params=None):
        pass
        
orphaned_marker = Suppress((CaselessKeyword("End") + CaselessKeyword("Function")) ^ \
                           (CaselessKeyword("End") + CaselessKeyword("Sub")))
orphaned_marker.setParseAction(Orphaned_Marker)


class EnumStatement(VBA_Object):

    def __init__(self, original_str, location, tokens):
        super(EnumStatement, self).__init__(original_str, location, tokens)
        self.name = str(tokens[0])
        self.values = []
        pos = 0
        enum_vals = tokens[1]
        last_val = -1
        for enum_val in enum_vals:
            last_val += 1
            if (len(enum_val) == 2):
                last_val = enum_val[1]
            self.values.append((str(enum_val[0]), last_val))
                
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed %r as Enum_Statement' % self)

    def __repr__(self):
        r = "Enum " + self.name + "\\n"
        for enum_val in self.values:
            r += "  " + enum_val[0] + " = " + str(enum_val[1]) + " \\n"
        r += "End Enum"
        return r

    def eval(self, context, params=None):

        for enum_val in self.values:
            context.set(enum_val[0], enum_val[1], force_global=True)

enum_value = Group((lex_identifier | enum_val_id)("name") + Optional(Suppress(Literal("=")) + integer("value")))
enum_statement = Suppress(Optional(CaselessKeyword('Public') | CaselessKeyword('Private'))) + \
                 Suppress(CaselessKeyword("Enum")) + lex_identifier("enum_name") + Suppress(EOS) + \
                 Group(OneOrMore(enum_value + Suppress(EOS))("enum_values")) + \
                 Suppress(CaselessKeyword("End")) + Suppress(CaselessKeyword("Enum"))
enum_statement.setParseAction(EnumStatement)
    
def extend_statement_grammar():

    global statement
    global statement_no_orphan
    global statement_restricted

    statement <<= try_catch | type_declaration | simple_for_statement | real_simple_for_each_statement | simple_if_statement | \
                  line_input_statement | simple_if_statement_macro | simple_while_statement | simple_do_statement | simple_select_statement | \
                  with_statement| simple_statement | rem_statement | \
                  (procedures.simple_function ^ orphaned_marker) | \
                  (procedures.simple_sub ^ orphaned_marker) | \
                  (procedures.property_let ^ orphaned_marker) | \
                  name_statement | stop_statement | enum_statement

    statement_no_orphan <<= try_catch | type_declaration | simple_for_statement | real_simple_for_each_statement | simple_if_statement | \
                            line_input_statement | simple_if_statement_macro | simple_while_statement | simple_do_statement | simple_select_statement | \
                            with_statement| simple_statement | rem_statement | procedures.simple_function | procedures.simple_sub | name_statement | stop_statement | \
                            enum_statement

    statement_restricted <<= try_catch | type_declaration | simple_for_statement | real_simple_for_each_statement | simple_if_statement | \
                             line_input_statement | simple_if_statement_macro | simple_while_statement | simple_do_statement | simple_select_statement | name_statement | \
                             with_statement| simple_statement_restricted | rem_statement | \
                             procedures.simple_function | procedures.simple_sub | stop_statement | enum_statement

