#!/usr/bin/env python



__version__ = '0.02'


import logging
import sys

from vba_context import *
from statements import *
from identifiers import *
import utils

from logger import log
from tagged_block_finder_visitor import *
from vba_object import to_python
from vba_object import _get_var_vals
from vba_object import _check_for_iocs


class Sub(VBA_Object):

    def __init__(self, original_str, location, tokens):
        super(Sub, self).__init__(original_str, location, tokens)
        self.name = tokens.sub_name
        self.params = tokens.params
        self.min_param_length = len(self.params)
        for param in self.params:
            if (param.is_optional):
                self.min_param_length -= 1
        self.statements = tokens.statements
        self.bogus_if = None
        if (len(tokens.bogus_if) > 0):
            self.bogus_if = tokens.bogus_if
        visitor = tagged_block_finder_visitor()
        self.accept(visitor)
        self.tagged_blocks = visitor.blocks
        log.info('parsed %r' % self)

    def __repr__(self):
        return 'Sub %s (%s): %d statement(s)' % (self.name, self.params, len(self.statements))

    def to_python(self, context, params=None, indent=0):

        tmp_context = Context(context=context, _locals=context.locals, copy_globals=True)
        global_var_info, _ = _get_var_vals(self, tmp_context, global_only=True)
        
        global_var_init_str = ""
        indent_str = " " * indent
        for global_var in global_var_info.keys():
            val = to_python(global_var_info[global_var], context)
            global_var_init_str += indent_str + str(global_var) + " = " + str(val) + "\n"

        tmp_context = Context(context=context)
        for param in self.params:
            tmp_context.set(param.name, "__FUNC_ARG__")

        tmp_context.curr_func_name = str(self.name)

        r = global_var_init_str
        
        indent_str = " " * indent
        func_args = "("
        first = True
        for param in self.params:
            if (not first):
                func_args += ", "
            first = False
            func_args += utils.fix_python_overlap(to_python(param, tmp_context))
        func_args += ")"
        r += indent_str + "def " + str(self.name) + func_args + ":\n"

        r += indent_str + " " * 4 + "import core.vba_library\n"
        r += indent_str + " " * 4 + "global vm_context\n\n"
        r += indent_str + " " * 4 + "# Function return value.\n"
        r += indent_str + " " * 4 + str(self.name) + " = 0\n\n"

        r += indent_str + " " * 4 + "# Referenced global variables.\n"
        for global_var in global_var_info.keys():
            r += indent_str + " " * 4 + "global " + str(global_var) + "\n"
        r += "\n"
        
        r += to_python(self.statements, tmp_context, indent=indent+4, statements=True)

        r += "\n" + _check_for_iocs(self, tmp_context, indent=indent+4)
        
        return r
    
    def eval(self, context, params=None):

        caller_context = context
        context = Context(context=caller_context)
        context.in_procedure = True

        context.curr_func_name = str(self.name)
        
        context.goto_executed = False
        
        context.tagged_blocks = self.tagged_blocks

        call_info = {}
        call_info["FUNCTION_NAME -->"] = self.name

        for param in self.params:
            init_val = None
            if (param.init_val is not None):
                init_val = eval_arg(param.init_val, context=context)
            call_info[param.name] = init_val

        self.byref_params = {}
        if ((params is not None) and (len(params) == len(self.params))):

            for i in range(len(params)):

                param_name = self.params[i].name
                param_value = params[i]

                if ((param_value == 0) and (self.params[i].my_type == "String")):
                    param_value = ""

                if ((self.params[i].my_type == "String") and (not self.params[i].is_array)):
                    param_value = str(param_value)
                    
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug('Function %s: setting param %s = %r' % (self.name, param_name, param_value))
                call_info[param_name] = param_value

                if (self.params[i].mechanism == "ByRef"):

                    self.byref_params[(param_name, i)] = None
            
        if (context.call_stack.count(call_info) > 0):
            log.warn("Recursive infinite loop detected. Aborting call " + str(call_info))
            return "NULL"

        context.call_stack.append(call_info)

        do_const_assignments(self.statements, context)
        
        for param_name in call_info.keys():
            context.set(param_name, call_info[param_name], force_local=True)

        old_global_scope = context.global_scope
        context.global_scope = False
                    
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('evaluating Sub %s(%s)' % (self.name, params))
        log.info('evaluating Sub %s' % self.name)
        context.clear_error()
        for s in self.statements:

            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug('Sub %s eval statement: %s' % (self.name, s))
            if (isinstance(s, VBA_Object)):
                s.eval(context=context)

            if (context.must_handle_error()):
                break
            context.clear_error()

            if (context.goto_executed or s.exited_with_goto):
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("GOTO executed. Control flow handled by GOTO, so skip rest of procedure statements.")
                break
            
        context.global_scope = old_global_scope
            
        context.handle_error(params)
        
        if (self.bogus_if is not None):
            if (isinstance(self.bogus_if, VBA_Object)):
                self.bogus_if.eval(context=context)
            elif (isinstance(self.bogus_if, list)):
                for cmd in self.bogus_if:
                    cmd.eval(context=context)

        for byref_param in self.byref_params.keys():
            self.byref_params[byref_param] = context.get(byref_param[0].lower())

        del context.call_stack[-1]

        context.goto_executed = False

        caller_context.got_error = context.got_error
        
        try:            
            context.get(self.name)
        except KeyError:

            context.set(self.name, '')

        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Returning from sub " + str(self))


procedure_scope = Optional(CaselessKeyword('Public') | CaselessKeyword('Private')
                           | CaselessKeyword('Global') | CaselessKeyword('Friend')).suppress()


static_keyword = Optional(CaselessKeyword('static'))






lifecycle_handler_name = CaselessKeyword("Class_Initialize") | CaselessKeyword("Class_Terminate")
implemented_name = identifier
event_handler_name = identifier
prefixed_name = identifier | lifecycle_handler_name

subroutine_name = identifier | lifecycle_handler_name


function_name = Combine(identifier + Suppress(Optional(type_suffix))) | lifecycle_handler_name




procedure_tail = FollowedBy(line_terminator) | comment_single_quote | Literal(":") + rem_statement



sub_start = Optional(CaselessKeyword('Static')) + public_private + Optional(CaselessKeyword('Static')) + CaselessKeyword('Sub').suppress() + lex_identifier('sub_name') \
            + Optional(params_list_paren) + EOS.suppress()
sub_start_single = Optional(CaselessKeyword('Static')) + public_private + CaselessKeyword('Sub').suppress() + lex_identifier('sub_name') \
                   + Optional(params_list_paren) + Suppress(':')
sub_end = (CaselessKeyword('End') + (CaselessKeyword('Sub') | CaselessKeyword('Function')) + EOS).suppress() | \
          bogus_simple_for_each_statement
simple_sub_end = (CaselessKeyword('End') + (CaselessKeyword('Sub') | CaselessKeyword('Function'))).suppress()
sub_end_single = Optional(Suppress(':')) + (CaselessKeyword('End') + (CaselessKeyword('Sub') | CaselessKeyword('Function')) + EOS).suppress()
multiline_sub = (sub_start + \
                 Group(ZeroOrMore(statements_line)).setResultsName('statements') + \
                 Optional(bad_if_statement('bogus_if')) + \
                 Suppress(Optional(bad_next_statement)) + \
                 sub_end)
simple_multiline_sub = (sub_start + \
                        Group(ZeroOrMore(statements_line)).setResultsName('statements') + \
                        Optional(bad_if_statement('bogus_if')) + \
                        Suppress(Optional(bad_next_statement)) + \
                        simple_sub_end)
singleline_sub = sub_start_single + simple_statements_line('statements') + sub_end_single
sub = singleline_sub | multiline_sub
simple_sub = simple_multiline_sub
sub.setParseAction(Sub)
simple_sub.setParseAction(Sub)

sub_start_line = public_private + CaselessKeyword('Sub').suppress() + lex_identifier('sub_name') \
                 + Optional(params_list_paren) + EOS.suppress()
sub_start_line.setParseAction(Sub)



def is_loop_statement(s):
    return (isinstance(s, For_Statement) or
            isinstance(s, For_Each_Statement) or
            isinstance(s, While_Statement) or
            isinstance(s, Do_Statement))

class Function(VBA_Object):

    def __init__(self, original_str, location, tokens):
        super(Function, self).__init__(original_str, location, tokens)
        self.return_type = None
        if (hasattr(tokens, "return_type")):
            self.return_type = tokens.return_type
        self.name = tokens.function_name
        self.params = tokens.params
        self.min_param_length = len(self.params)
        for param in self.params:
            if (param.is_optional):
                self.min_param_length -= 1
        self.statements = tokens.statements
        try:
            len(self.statements)
        except:
            self.statements = [self.statements]
        self.return_type = tokens.return_type
        self.vars = {}
        self.bogus_if = None
        if (len(tokens.bogus_if) > 0):
            self.bogus_if = tokens.bogus_if
        visitor = tagged_block_finder_visitor()
        self.accept(visitor)
        self.tagged_blocks = visitor.blocks
        log.info('parsed %r' % self)

    def __repr__(self):
        return 'Function %s (%s): %d statement(s)' % (self.name, self.params, len(self.statements))

    def to_python(self, context, params=None, indent=0):
        
        tmp_context = Context(context=context, _locals=context.locals, copy_globals=True)
        global_var_info, _ = _get_var_vals(self, tmp_context, global_only=True)
        
        global_var_init_str = ""
        indent_str = " " * indent
        for global_var in global_var_info.keys():
            val = to_python(global_var_info[global_var], context)
            global_var_init_str += indent_str + str(global_var) + " = " + str(val) + "\n"
        
        tmp_context = Context(context=context)
        for param in self.params:
            tmp_context.set(param.name, "__FUNC_ARG__")

        tmp_context.curr_func_name = str(self.name)

        r = global_var_init_str
        
        func_args = "("
        first = True
        for param in self.params:
            if (not first):
                func_args += ", "
            first = False
            func_args += utils.fix_python_overlap(to_python(param, tmp_context))
        func_args += ")"
        r += indent_str + "def " + str(self.name) + func_args + ":\n"

        r += indent_str + " " * 4 + "import core.vba_library\n"
        r += indent_str + " " * 4 + "global vm_context\n\n"
        r += indent_str + " " * 4 + "# Function return value.\n"
        r += indent_str + " " * 4 + str(self.name) + " = 0\n\n"

        r += indent_str + " " * 4 + "# Referenced global variables.\n"
        for global_var in global_var_info.keys():
            r += indent_str + " " * 4 + "global " + str(global_var) + "\n"
        r += "\n"
            
        r += to_python(self.statements, tmp_context, indent=indent+4, statements=True)

        r += "\n" + _check_for_iocs(self, tmp_context, indent=indent+4)
        
        r += "\n" + indent_str + " " * 4 + "return " + str(self.name) + "\n"

        return r

    def eval(self, context, params=None):

        caller_context = context
        context = Context(context=caller_context)        
        context.in_procedure = True

        context.goto_executed = False

        context.curr_func_name = str(self.name)
        
        context.tagged_blocks = self.tagged_blocks

        call_info = {}
        call_info["FUNCTION_NAME -->"] = (self.name, None)

        if (len(self.params) == 0):
            call_info[self.name] = ('NULL', None)

        for param in self.params:
            init_val = None
            if (param.init_val is not None):
                init_val = eval_arg(param.init_val, context=context)
            call_info[param.name] = (init_val, None)
            
        array_indices = None
        if ((self.params is not None) and
            (params is not None) and
            (len(self.params) == 0) and
            (len(params) > 0)):
            array_indices = params
            
        self.byref_params = {}
        defined_param_pos = -1
        for defined_param in self.params:

            defined_param_pos += 1
            param_value = "NULL"
            param_name = defined_param.name
            if ((params is not None) and (defined_param_pos < len(params))):
                param_value = params[defined_param_pos]

            if (((param_value == 0) or (param_value == "NULL")) and (defined_param.my_type == "String")):
                param_value = ""

            if (defined_param.my_type == "String"):
                param_value = utils.safe_str_convert(param_value)
                    
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug('Function %s: setting param %s = %r' % (self.name, param_name, param_value))

            if ((param_name not in call_info) or
                (call_info[param_name] == ('', None)) or
                (param_value != "")):
                call_info[param_name] = (param_value, defined_param.my_type)

            if (defined_param.mechanism == "ByRef"):

                self.byref_params[(param_name, defined_param_pos)] = None
                
        if (context.call_stack.count(call_info) > 0):
            log.warn("Recursive infinite loop detected. Aborting call " + str(call_info))
            return "NULL"

        context.call_stack.append(call_info)
        
        do_const_assignments(self.statements, context)
        
        for param_name in call_info.keys():
            param_val, param_type = call_info[param_name]
            context.set(param_name, param_val, var_type=param_type, force_local=True)
        
        old_global_scope = context.global_scope
        context.global_scope = False
        
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('evaluating Function %s(%s)' % (self.name, params))
        context.clear_error()
        for s in self.statements:
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug('Function %s eval statement: %s' % (self.name, s))
            if (isinstance(s, VBA_Object)):
                s.eval(context=context)

            if (context.exit_func):
                break

            if (context.must_handle_error()):
                break
            context.clear_error()

            if (is_loop_statement(s)):
                context.goto_executed = False
            
            if (context.goto_executed or s.exited_with_goto):
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("GOTO executed. Go to next loop iteration.")
                break
            
        context.global_scope = old_global_scope
        
        context.handle_error(params)
        
        if (self.bogus_if is not None):
            self.bogus_if.eval(context=context)

        del context.call_stack[-1]
        
        context.goto_executed = False

        caller_context.got_error = context.got_error
        
        context.exit_func = False
        try:

            for byref_param in self.byref_params.keys():
                if (context.contains(byref_param[0].lower())):
                    self.byref_params[byref_param] = context.get(byref_param[0].lower())

            return_value = context.get(self.name, local_only=True)
            if ((return_value is None) or (isinstance(return_value, Function))):
                return_value = ''
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug('Function %s: return value = %r' % (self.name, return_value))

            if ((self.return_type == "String") and (not isinstance(return_value, str))):
                return_value = coerce_to_string(return_value)

            if (array_indices is not None):

                if (isinstance(return_value, list)):

                    all_int = True
                    for i in array_indices:
                        if (not isinstance(i, int)):
                            all_int = False
                            break
                    if (all_int):

                        for i in array_indices:
                            return_value = return_value[i]

                    else:
                        log.warn("Array indices " + str(array_indices) + " are invalid. " + \
                                 "Not doing array access of function return value.")

                else:
                    log.warn(str(self) + " does not return an array. Not doing array access.")
                    
            for global_var in context.globals.keys():
                caller_context.globals[global_var] = context.globals[global_var]
                    
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("Returning from func " + str(self))
            return return_value

        except KeyError:
            
            return ''

function_start = Optional(CaselessKeyword('Static')) + Optional(public_private) + Optional(CaselessKeyword('Static')) + \
                 CaselessKeyword('Function').suppress() + TODO_identifier_or_object_attrib('function_name') + \
                 Optional(params_list_paren) + Optional(function_type2("return_type")) + EOS.suppress()
function_start_single = Optional(CaselessKeyword('Static')) + Optional(public_private) + Optional(CaselessKeyword('Static')) + \
                        CaselessKeyword('Function').suppress() + TODO_identifier_or_object_attrib('function_name') + \
                        Optional(params_list_paren) + Optional(function_type2) + Suppress(':')

function_end = (CaselessKeyword('End') + CaselessKeyword('Function') + EOS).suppress() | \
               (bogus_simple_for_each_statement + Suppress(EOS))
simple_function_end = (CaselessKeyword('End') + CaselessKeyword('Function')).suppress()
function_end_single = Optional(Suppress(':')) + (CaselessKeyword('End') + CaselessKeyword('Function') + EOS).suppress()

multiline_function = (function_start + \
                      Group(ZeroOrMore(statements_line)).setResultsName('statements') + \
                      Optional(bad_if_statement('bogus_if')) + \
                      Suppress(Optional(bad_next_statement)) + \
                      function_end)
simple_multiline_function = (function_start + \
                             Group(ZeroOrMore(statements_line)).setResultsName('statements') + \
                             Optional(bad_if_statement('bogus_if')) + \
                             Suppress(Optional(bad_next_statement)) + \
                             simple_function_end)

singleline_function = function_start_single + simple_statements_line('statements') + function_end_single
function = singleline_function | multiline_function
simple_function = simple_multiline_function            
function.setParseAction(Function)
simple_function.setParseAction(Function)

function_start_line = public_private + CaselessKeyword('Function').suppress() + lex_identifier('function_name') \
                 + Optional(params_list_paren) + Optional(function_type2) + EOS.suppress()
function_start_line.setParseAction(Function)


class PropertyLet(Sub):

    def __init__(self, original_str, location, tokens):
        super(PropertyLet, self).__init__(original_str, location, tokens)
        self.name = tokens.property_name
        self.params = tokens.params
        self.min_param_length = len(self.params)
        for param in self.params:
            if (param.is_optional):
                self.min_param_length -= 1
        self.statements = tokens.statements
        try:
            len(self.statements)
        except:
            self.statements = [self.statements]
        visitor = tagged_block_finder_visitor()
        self.accept(visitor)
        self.tagged_blocks = visitor.blocks
        log.info('parsed %r' % self)

    def __repr__(self):
        return 'Property Let %s (%s): %d statement(s)' % (self.name, self.params, len(self.statements))


property_let = Optional(CaselessKeyword('Static')) + public_private + Optional(CaselessKeyword('Static')) + \
               CaselessKeyword('Property').suppress() + CaselessKeyword('Let').suppress() + \
               lex_identifier('property_name') + params_list_paren + \
               Group(ZeroOrMore(statements_line)).setResultsName('statements') + \
               (CaselessKeyword('End') + CaselessKeyword('Property') + EOS).suppress()
property_let.setParseAction(PropertyLet)

extend_statement_grammar()


