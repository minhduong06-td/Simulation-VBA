#!/usr/bin/env python
"""
SimulationVBA: VBA Grammar - Base class for all VBA objects

SimulationVBA is a specialized engine to parse, analyze and interpret Microsoft
VBA macros (Visual Basic for Applications), mainly for malware analysis.

Author: Philippe Lagadec - http://www.decalage.info
License: BSD, see source code or documentation

Project Repository:
https://github.com/decalage2/ViperMonkey
"""

# === LICENSE ==================================================================




__version__ = '0.08'



import logging
from logger import log
import re
from curses_ascii import isprint
import traceback
import string
import gc
import hashlib

from inspect import getouterframes, currentframe
import sys
from datetime import datetime
import pyparsing

import expressions
from var_in_expr_visitor import *
from function_call_visitor import *
from lhs_var_visitor import *
from utils import safe_print
import utils
from let_statement_visitor import *
from vba_context import *
import excel

max_emulation_time = None

class VbaLibraryFunc(object):
    """
    Marker class to tell if a class implements a VBA function.
    """

    def num_args(self):
        """
        Get the # of arguments (minimum) required by the functio.
        """
        log.warning("Using default # args of 1 for " + str(type(self)))
        return 1

    def return_type(self):
        """
        Get the python type returned from the emulated function ('INTEGER' or 'STRING').
        """
        log.warning("Using default return type of 'INTEGER' for " + str(type(self)))
        return "INTEGER"

def excel_col_letter_to_index(x): 
    x = x.upper()
    return (reduce(lambda s,a:s*26+ord(a)-ord('A')+1, x, 0) - 1)

def limits_exceeded(throw_error=False):
    """
    Check to see if we are about to exceed the maximum recursion depth. Also check to 
    see if emulation is taking too long (if needed).
    """

    level = len(getouterframes(currentframe(1)))
    recursion_exceeded = (level > (sys.getrecursionlimit() * .50))
    time_exceeded = False

    if (max_emulation_time is not None):
        time_exceeded = (datetime.now() > max_emulation_time)

    if (recursion_exceeded):
        log.error("Call recursion depth approaching limit.")
        if (throw_error):
            raise RuntimeError("The SimulationVBA recursion depth will be exceeded. Aborting analysis.")
    if (time_exceeded):
        log.error("Emulation time exceeded.")
        if (throw_error):
            raise RuntimeError("The SimulationVBA emulation time limit was exceeded. Aborting analysis.")
        
    return (recursion_exceeded or time_exceeded)

class VBA_Object(object):
    """
    Base class for all VBA objects that can be evaluated.
    """

    loop_upper_bound = 10000000
    
    def __init__(self, original_str, location, tokens):
        """
        VBA_Object constructor, to be called as a parse action by a pyparsing parser

        :param original_str: original string matched by the parser
        :param location: location of the match
        :param tokens: tokens extracted by the parser
        :return: nothing
        """
        self.original_str = original_str
        self.location = location
        self.tokens = tokens
        self._children = None
        self.is_useless = False
        self.is_loop = False
        self.exited_with_goto = False
        
    def eval(self, context, params=None):
        """
        Evaluate the current value of the object.

        :param context: Context for the evaluation (local and global variables)
        :return: current value of the object
        """
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug(self)

    def full_str(self):
        return str(self)
        
    def get_children(self):
        """
        Return the child VBA objects of the current object.
        """

        limits_exceeded(throw_error=True)
        
        if ((hasattr(self, "_children")) and (self._children is not None)):
            return self._children
        r = []
        for _, value in self.__dict__.iteritems():
            if (isinstance(value, VBA_Object)):
                r.append(value)
            if ((isinstance(value, list)) or
                (isinstance(value, pyparsing.ParseResults))):
                for i in value:
                    if (isinstance(i, VBA_Object)):
                        r.append(i)
            if (isinstance(value, dict)):
                for i in value.values():
                    if (isinstance(i, VBA_Object)):
                        r.append(i)
        self._children = r
        return r
                        
    def accept(self, visitor, no_embedded_loops=False):
        """
        Visitor design pattern support. Accept a visitor.
        """

        limits_exceeded(throw_error=True)
        
        if (no_embedded_loops and
            hasattr(visitor, "in_loop") and
            visitor.in_loop and
            self.is_loop):
            return

        if (not hasattr(visitor, "in_loop")):
            visitor.in_loop = self.is_loop

        if ((not visitor.in_loop) and (self.is_loop)):
            visitor.in_loop = True

        visit_status = visitor.visit(self)
        if (not visit_status):
            return

        old_in_loop = visitor.in_loop
        
        for child in self.get_children():
            child.accept(visitor, no_embedded_loops=no_embedded_loops)

        visitor.in_loop = old_in_loop

    def to_python(self, context, params=None, indent=0):
        """
        JIT compile this VBA object to Python code for direct emulation.
        """
        raise NotImplementedError("to_python() not implemented in " + str(type(self)))

def _read_from_excel(arg, context):
    """
    Try to evaluate an argument by reading from the loaded Excel spreadsheet.
    """

    if ("MemberAccessExpression" not in str(type(arg))):
        return None        
    arg_str = str(arg)
    if (("sheets(" in arg_str.lower()) and
        (("range(" in arg_str.lower()) or ("cells(" in arg_str.lower()))):
        
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Try as Excel cell read...")

        return arg.eval(context)
            
        tmp_arg_str = arg_str.lower()
        start = tmp_arg_str.index("sheets(") + len("sheets(")
        end = start + tmp_arg_str[start:].index(")")
        sheet_name = arg_str[start:end].strip().replace('"', "").replace("'", "").replace("//", "")
        
        start = None
        if ("range(" in arg_str.lower()):
            start = tmp_arg_str.index("range(") + len("range(")
        else:
            start = tmp_arg_str.index("cells(") + len("cells(")
        end = start + tmp_arg_str[start:].index(")")
        cell_index = arg_str[start:end].strip().replace('"', "").replace("'", "").replace("//", "")
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Sheet name = '" + sheet_name + "', cell index = " + cell_index)
        
        try:
            
            sheet = context.loaded_excel.sheet_by_name(sheet_name)
            

            index_pat = r"(\d+)\s*,\s*(\d+)"
            if (re.search(index_pat, cell_index) is not None):
                indices = re.findall(index_pat, cell_index)[0]
                row = int(indices[0]) - 1
                col = int(indices[1]) - 1

            else:
                col = ""
                row = ""
                for c in cell_index:
                    if (c.isalpha()):
                        col += c
                    else:
                        row += c
                    
                row = int(row) - 1
                col = excel_col_letter_to_index(col)
            
            val = str(sheet.cell_value(row, col))
            
            log.info("Read cell (" + str(cell_index) + ") from sheet " + str(sheet_name))
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("Cell value = '" + str(val) + "'")
            return val
        
        except Exception as e:
            context.report_general_error("Cannot read cell from Excel spreadsheet. " + str(e))

def _read_from_object_text(arg, context):
    """
    Try to read in a value from the text associated with a object like a Shape.
    """

    arg_str = str(arg)
    arg_str_low = arg_str.lower().strip()

    if (("shapes(" in arg_str_low) and 
        (not isinstance(arg, expressions.Function_Call)) and
        (not type(arg).__name__ == 'Concatenation')):

        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("eval_arg: Try to get as ....TextFrame.TextRange.Text value: " + arg_str.lower())

        lhs = "Shapes('1')"
        if ("inlineshapes" in arg_str_low):
            lhs = "InlineShapes('1')"
        if ("MemberAccessExpression" in str(type(arg))):

            lhs = arg.lhs
            if ((str(lhs) == "ActiveDocument") or (str(lhs) == "ThisDocument")):
                lhs = arg.rhs[0]
        
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("eval_obj_text: Old member access lhs = " + str(lhs))
            if ((hasattr(lhs, "eval")) and
                (not isinstance(lhs, pyparsing.ParseResults))):
                lhs = lhs.eval(context)
            else:

                var_name = str(lhs)
                try:
                    lhs = context.get(var_name)
                except KeyError:
                    lhs = var_name

            if (lhs == "NULL"):
                lhs = "Shapes('1')"
            if ("inlineshapes" in arg_str_low):
                lhs = "InlineShapes('1')"
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("eval_obj_text: Evaled member access lhs = " + str(lhs))
        
        doc_var_name = str(lhs) + ".TextFrame.TextRange.Text"
        doc_var_name = doc_var_name.replace(".TextFrame.TextFrame", ".TextFrame")
        if (("InlineShapes(" in doc_var_name) and (not doc_var_name.startswith("InlineShapes("))):
            doc_var_name = doc_var_name[doc_var_name.index("InlineShapes("):]
        elif (("Shapes(" in doc_var_name) and
              (not doc_var_name.startswith("Shapes(")) and
              ("InlineShapes(" not in doc_var_name)):
            doc_var_name = doc_var_name[doc_var_name.index("Shapes("):]
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("eval_obj_text: Looking for object text " + str(doc_var_name))
        val = context.get_doc_var(doc_var_name.lower())
        if (val is not None):
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("eval_obj_text: Found " + str(doc_var_name) + " = " + str(val))
            return val

        lhs_str = str(lhs)
        if ("'" not in lhs_str):
            return None
        new_lhs = lhs_str[:lhs_str.index("'") + 1] + "1" + lhs_str[lhs_str.rindex("'"):]
        doc_var_name = new_lhs + ".TextFrame.TextRange.Text"
        doc_var_name = doc_var_name.replace(".TextFrame.TextFrame", ".TextFrame")
        if (("InlineShapes(" in doc_var_name) and (not doc_var_name.startswith("InlineShapes("))):
            doc_var_name = doc_var_name[doc_var_name.index("InlineShapes("):]
        elif (("Shapes(" in doc_var_name) and
              (not doc_var_name.startswith("Shapes(")) and
              ("InlineShapes(" not in doc_var_name)):
            doc_var_name = doc_var_name[doc_var_name.index("Shapes("):]
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("eval_arg: Fallback, looking for object text " + str(doc_var_name))
        val = context.get_doc_var(doc_var_name.lower())
        return val

def contains_excel(arg):
    """
    See if a given expression contains Excel book or sheet objects.
    """

    if (isinstance(arg, excel.ExcelSheet) or
        isinstance(arg, excel.ExcelBook)):
        return True
    
    if (not isinstance(arg, expressions.Function_Call)):
        return False

    excel_funcs = set(["usedrange", "sheets", "specialcells"])
    return (str(arg.name).lower() in excel_funcs)
    
constant_expr_cache = {}

def get_cached_value(arg):
    """
    Get the cached value of an all constant numeric expression if we have it.
    """

    if (isinstance(arg, int) or
        isinstance(arg, dict)):
        return arg

    if contains_excel(arg):
        return None

    arg_str = str(arg)
    if (arg_str not in constant_expr_cache.keys()):
        return None
    return constant_expr_cache[arg_str]

def set_cached_value(arg, val):
    """
    Set the cached value of an all constant numeric expression.
    """

    if ((not isinstance(val, int)) and
        (not isinstance(val, float)) and
        (not isinstance(val, complex))):
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.warning("Expression '" + str(val) + "' is a " + str(type(val)) + ", not an int. Not caching.")
        return

    if contains_excel(arg):
        return
        
    arg_str = str(arg)
    try:
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Cache value of " + arg_str + " = " + str(val))
    except UnicodeEncodeError:
        pass
    constant_expr_cache[arg_str] = val
    
def is_constant_math(arg):
    """
    See if a given expression is a simple math expression with all literal numbers.
    """

    if (isinstance(arg, VBA_Object)):
        var_visitor = var_in_expr_visitor()
        arg.accept(var_visitor)
        if (len(var_visitor.variables) > 0):
            return False

    if (isinstance(arg, dict) or
        contains_excel(arg)):
        return False
        
    try:
        import rure as local_re
    except ImportError:
        import re as local_re

    base_pat = "(?:\\s*\\d+(?:\\.\\d+)?\\s*[+\\-\\*/]\\s*)*\\s*\\d+"
    paren_pat = base_pat + "|(?:\\((?:\\s*" + base_pat + "\\s*[+\\-\\*\\\\]\\s*)*\\s*" + base_pat + "\\))"
    arg_str = str(arg).strip()
    try:
        arg_str = unicode(arg_str)
    except UnicodeDecodeError:
        arg_str = filter(isprint, arg_str)
        arg_str = unicode(arg_str)
    return (local_re.match(unicode(paren_pat), arg_str) is not None)

meta = None

def _boilerplate_to_python(indent):
    """
    Get starting boilerplate code for VB to Python JIT code.
    """
    indent_str = " " * indent
    boilerplate = indent_str + "import core.vba_library\n"
    boilerplate = indent_str + "import core.vba_context\n"
    boilerplate += indent_str + "from core.utils import safe_print\n"
    boilerplate += indent_str + "from core.utils import plus\n"
    boilerplate += indent_str + "from core.utils import eq\n"
    boilerplate += indent_str + "from core.utils import neq\n"
    boilerplate += indent_str + "import core.utils\n"
    boilerplate += indent_str + "from core.vba_object import update_array\n"
    boilerplate += indent_str + "from core.vba_object import coerce_to_num\n"
    boilerplate += indent_str + "from core.vba_object import coerce_to_int\n"
    boilerplate += indent_str + "from core.vba_object import coerce_to_str\n"
    boilerplate += indent_str + "from core.vba_object import coerce_to_int_list\n\n"
    boilerplate += indent_str + "try:\n"
    boilerplate += indent_str + " " * 4 + "vm_context\n"
    boilerplate += indent_str + "except (NameError, UnboundLocalError):\n"
    boilerplate += indent_str + " " * 4 + "vm_context = context\n"
    return boilerplate

def _get_local_func_type(expr, context):
    """
    Get the return type of a locally defined funtion given a call
    to the function.
    """

    if (not isinstance(expr, expressions.Function_Call)):
        return None

    func_def = None
    try:
        func_def = context.get(expr.name)
    except KeyError:
        return None

    if (hasattr(func_def, "return_type")):
        return func_def.return_type
    return None
        
def _infer_type_of_expression(expr, context):
    """
    Try to determine if a given expression is an "INTEGER" or "STRING" expression.
    """

    import operators
    import vba_library


    if (hasattr(expr, "return_type")):
        return expr.return_type()

    if (isinstance(expr, expressions.Function_Call)):

        if (expr.name.lower() in vba_library.VBA_LIBRARY):
            builtin = vba_library.VBA_LIBRARY[expr.name.lower()]
            if (hasattr(builtin, "return_type")):
                return builtin.return_type()

        r = _get_local_func_type(expr, context)
        return r
        
    if (isinstance(expr, operators.Xor) or
        isinstance(expr, operators.And) or
        isinstance(expr, operators.Or) or
        isinstance(expr, operators.Not) or
        isinstance(expr, operators.Neg) or
        isinstance(expr, operators.Subtraction) or
        isinstance(expr, operators.Multiplication) or
        isinstance(expr, operators.Power) or
        isinstance(expr, operators.Division) or
        isinstance(expr, operators.MultiDiv) or
        isinstance(expr, operators.FloorDivision) or
        isinstance(expr, operators.Mod) or        
        isinstance(expr, operators.Xor)):
        return "INTEGER"

    if (isinstance(expr, operators.Concatenation)):
        return "STRING"
    
    if (isinstance(expr, operators.AddSub) or
        isinstance(expr, expressions.BoolExpr) or
        isinstance(expr, expressions.BoolExprItem)):

        if ((hasattr(expr, "operators")) and ("-" in expr.operators)):
            return "INTEGER"
        
        r_type = None
        for child in expr.get_children():
            child_type = _infer_type_of_expression(child, context)
            if (child_type is not None):
                r_type = child_type
        return r_type

    return None
    
def _infer_type(var, code_chunk, context):
    """
    Try to infer the type of an undefined variable based on how it is used ("STRING" or "INTEGER").

    This is currently purely a heuristic.

    returns a tuple, 1st element is the inferred type ("STRING" or "INTEGER") and the 2nd element 
    is a flag indicating if we are sure of the type (True) or just guessing (False).
    """

    visitor = let_statement_visitor(var)
    code_chunk.accept(visitor)

    str_funcs = ["cstr(", "chr(", "left(", "right(", "mid(", "join(", "lcase(",
                 "replace(", "trim(", "ucase(", "chrw(", " & "]
    for assign in visitor.let_statements:

        poss_type = _infer_type_of_expression(assign.expression, context)
        if ((poss_type is not None) and (poss_type != "UNKNOWN")):
            return (poss_type, True)
        
        rhs = str(assign.expression).lower()
        for str_func in str_funcs:
            if (str_func in rhs):
                return ("STRING", True)

    return ("INTEGER", False)

def _get_var_vals(item, context, global_only=False):
    """
    Get the current values for all of the referenced VBA variables that appear in the 
    given VBA object.

    Returns a dict mapping var names to values.
    """

    import procedures
    import statements


    var_visitor = var_in_expr_visitor(context)
    item.accept(var_visitor, no_embedded_loops=False)
    var_names = var_visitor.variables

    lhs_visitor = lhs_var_visitor()
    item.accept(lhs_visitor, no_embedded_loops=False)
    lhs_var_names = lhs_visitor.variables
    
    var_names = var_names.union(lhs_var_names)
    tmp = set()
    for var in var_names:
        tmp.add(var)
        if ("." in var):
            tmp.add(var[:var.index(".")])
    var_names = tmp

    if (context.with_prefix_raw is not None):
        var_names.add(str(context.with_prefix_raw))
    
    r = {}
    zero_arg_funcs = set()
    for var in var_names:

        if ("(" in var):
            continue

        val = None
        orig_val = None
        try:

            val = context.get(var, global_only=global_only)
            orig_val = val
            
            if (global_only and (var in lhs_var_names)):
                continue
            
            if ((val == "__FUNC_ARG__") or
                (val == "__ALREADY_SET__") or
                (val == "__LOOP_VAR__")):
                continue
            
            if (isinstance(val, procedures.Function) or
                isinstance(val, procedures.Sub) or
                isinstance(val, statements.External_Function) or
                isinstance(val, VbaLibraryFunc)):

                val = None
                
                if (var not in lhs_var_names):
                    zero_arg_funcs.add(var)

                    context.set("__ORIG__" + var, orig_val, force_local=True)
                    context.set("__ORIG__" + var, orig_val, force_global=True)
                    continue

            val_str = None
            try:
                val_str = str(val).strip()
            except UnicodeEncodeError:
                val_str = filter(isprint, val).strip()
            if ((val_str == "inf") or
                (val_str == "-inf")):
                val = None

            if (val_str == "NULL"):
                val = None

            if ("core.vba_library.run_function" in val_str):
                val = 0
            
        except KeyError:
            if global_only:
                continue

        if (val is None):

            var_type, certain_of_type = _infer_type(var, item, context)
            if (var_type == "INTEGER"):
                val = 0
                if certain_of_type:
                    context.set_type(var, "Integer")
            elif (var_type == "STRING"):
                val = ""
                if certain_of_type:
                    context.set_type(var, "String")
            else:
                log.warning("Type '" + str(var_type) + "' of var '" + str(var) + "' not handled." + \
                            " Defaulting initial value to 0.")
                val = 0

        var = utils.fix_python_overlap(var)
            
        r[var] = val

        if (utils.safe_str_convert(val) == "RegExp"):
            if (context.contains("RegExp.pattern")):
                pval = to_python(context.get("RegExp.pattern"), context)
                if (pval.startswith('"')):
                    pval = pval[1:]
                if (pval.endswith('"')):
                    pval = pval[:-1]
                r[var + ".Pattern"] = pval
            if (context.contains("RegExp.global")):
                gval = to_python(context.get("RegExp.global"), context)
                gval = gval.replace('"', "")
                if (gval == "True"):
                    gval = True
                if (gval == "False"):
                    gval = False
                r[var + ".Global"] = gval
        
        context.set(var, "__ALREADY_SET__", force_local=True)
        context.set(var, "__ALREADY_SET__", force_global=True)
        
        if (orig_val is None):
            orig_val = val
        context.set("__ORIG__" + var, orig_val, force_local=True)
        context.set("__ORIG__" + var, orig_val, force_global=True)
        
    return (r, zero_arg_funcs)

def _loop_vars_to_python(loop, context, indent):
    """
    Set up initialization of variables used in a loop in Python.
    """
    indent_str = " " * indent
    loop_init = ""
    init_vals, _ = _get_var_vals(loop, context)
    sorted_vars = list(init_vals.keys())
    sorted_vars.sort()
    for var in sorted_vars:
        val = to_python(init_vals[var], context)
        var_name = str(var)
        if ((not var_name.endswith(".Pattern")) and
            (not var_name.endswith(".Global"))):
            var_name = var_name.replace(".", "")
        loop_init += indent_str + var_name + " = " + val + "\n"
    try:
        hash_object = hashlib.md5(str(loop).encode())
    except UnicodeDecodeError:
        hash_object = hashlib.md5(filter(isprint, str(loop)).encode())

    prog_var = "pct_" + hash_object.hexdigest()
    loop_init += indent_str + prog_var + " = 0\n"
    loop_init = indent_str + "# Initialize variables read in the loop.\n" + loop_init
    return (loop_init, prog_var)

def to_python(arg, context, params=None, indent=0, statements=False):
    """
    Call arg.to_python() if arg is a VBAObject, otherwise just return arg as a str.
    """
        
    r = None
    _arg_type = type(arg).__name__
    _has_tp = hasattr(arg, "to_python")
    _tp_type_str = str(type(arg.to_python)) if _has_tp else "N/A"
    _check_result = (_has_tp and
                     ((_tp_type_str == "<type 'method'>") or
                      (_tp_type_str == "<type 'instancemethod'>")))
    if (_check_result):
        r = arg.to_python(context, params=params, indent=indent)
    elif (_has_tp and isinstance(arg, VBA_Object)):
        log.info(
            "to_python type check FAILED for %s: has_tp=%s _tp_type=%s fallback=str()",
            _arg_type, _has_tp, _tp_type_str,
        )

    elif (isinstance(arg, str)):

        the_str = str(arg)
        the_str = str(the_str).\
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
            the_str = the_str.replace(chr(i), repl)
        for i in range(11, 13):
            repl = hex(i).replace("0x", "")
            if (len(repl) == 1):
                repl = "0" + repl
            repl = "\\x" + repl
            the_str = the_str.replace(chr(i), repl)
        for i in range(14, 32):
            repl = hex(i).replace("0x", "")
            if (len(repl) == 1):
                repl = "0" + repl
            repl = "\\x" + repl
            the_str = the_str.replace(chr(i), repl)
        for i in range(127, 255):
            repl = hex(i).replace("0x", "")
            if (len(repl) == 1):
                repl = "0" + repl
            repl = "\\x" + repl
            the_str = the_str.replace(chr(i), repl)
        r = " " * indent + '"' + the_str + '"'

    elif ((isinstance(arg, list) or
           isinstance(arg, pyparsing.ParseResults)) and statements):
        r = ""
        indent_str = " " * indent
        for statement in arg:
            r += indent_str + "try:\n"
            try:
                r += to_python(statement, context, indent=indent+4) + "\n"
            except Exception as e:
                return "ERROR! to_python failed! " + str(e)
            r += indent_str + "except Exception as e:\n"
            if (log.getEffectiveLevel() == logging.DEBUG):
                r += indent_str + " " * 4 + "safe_print(\"ERROR: \" + str(e))\n"
            else:
                r += indent_str + " " * 4 + "pass\n"

    else:
        arg_str = None
        try:
            arg_str = str(arg)
        except UnicodeEncodeError:
            arg_str = filter(isprint, arg)
        r = " " * indent + arg_str

        
    return r

def _check_for_iocs(loop, context, indent):
    """
    Check the variables modified in a loop to see if they were
    set to interesting IOCs.
    """
    indent_str = " " * indent
    lhs_visitor = lhs_var_visitor()
    loop.accept(lhs_visitor)
    lhs_var_names = lhs_visitor.variables
    ioc_str = indent_str + "# Check for IOCs in intermediate variables.\n"
    for var in lhs_var_names:
        py_var = utils.fix_python_overlap(var)
        ioc_str += indent_str + "try:\n"
        ioc_str += indent_str + " "*4 + "vm_context.save_intermediate_iocs(" + py_var + ")\n"
        ioc_str += indent_str + "except:\n"
        ioc_str += indent_str + " "* 4 + "pass\n"
    return ioc_str

def _updated_vars_to_python(loop, context, indent):
    """
    Save the variables updated in a loop in Python.
    """
    import statements
    
    indent_str = " " * indent
    lhs_visitor = lhs_var_visitor()
    loop.accept(lhs_visitor)
    lhs_var_names = lhs_visitor.variables
    if (context.with_prefix_raw is not None):
        lhs_var_names.add(str(context.with_prefix_raw))
    if (isinstance(loop, statements.For_Statement)):
        lhs_var_names.add(str(loop.name))
    var_dict_str = "{"
    first = True
    for var in lhs_var_names:
        py_var = utils.fix_python_overlap(var)
        if (not first):
            var_dict_str += ", "
        first = False
        var = var.replace(".", "")
        var_dict_str += '"' + var + '" : ' + py_var
    var_dict_str += "}"
    save_vals = indent_str + "try:\n"
    save_vals += indent_str + " " * 4 + "var_updates\n"
    save_vals += indent_str + " " * 4 + "var_updates.update(" + var_dict_str + ")\n"
    save_vals += indent_str + "except (NameError, UnboundLocalError):\n"
    save_vals += indent_str + " " * 4 + "var_updates = " + var_dict_str + "\n"
    save_vals += indent_str + 'var_updates["__shell_code__"] = core.vba_library.get_raw_shellcode_data()\n'
    save_vals = indent_str + "# Save the updated variables for reading into SimulationVBA.\n" + save_vals
    if (log.getEffectiveLevel() == logging.DEBUG):
        save_vals += indent_str + "print \"UPDATED VALS!!\"\n"
        save_vals += indent_str + "print var_updates\n"
    return save_vals

def _get_all_called_funcs(item, context):
    """
    Get all of the local functions called in the given VBA object.
    """

    call_visitor = function_call_visitor()
    item.accept(call_visitor)
    func_names = call_visitor.called_funcs

    tmp_context = Context(context=context, _locals=context.locals, copy_globals=True)
    _, zero_arg_funcs = _get_var_vals(item, tmp_context)
    func_names.update(zero_arg_funcs)
    
    local_funcs = []
    for func_name in func_names:
        if (context.contains(func_name)):
            curr_func = context.get(func_name)
            if (isinstance(curr_func, VBA_Object)):
                local_funcs.append(curr_func)

    return local_funcs

def _called_funcs_to_python(loop, context, indent):
    """
    Convert all the functions called in the loop to Python.
    """
    
    local_funcs = _get_all_called_funcs(loop, context)
    local_func_hashes = set()
    for curr_func in local_funcs:
        curr_func_hash = hashlib.md5(str(curr_func).encode()).hexdigest()
        local_func_hashes.add(curr_func_hash)
        
    seen_funcs = set()
    funcs_to_handle = list(local_funcs)
    while (len(funcs_to_handle) > 0):

        curr_func = funcs_to_handle.pop()
        curr_func_hash = hashlib.md5(str(curr_func).encode()).hexdigest()
        
        if (curr_func_hash in seen_funcs):
            continue
        seen_funcs.add(curr_func_hash)

        curr_local_funcs = _get_all_called_funcs(curr_func, context)

        for new_func in curr_local_funcs:
            new_func_hash = hashlib.md5(str(new_func).encode()).hexdigest()
            if (new_func_hash not in local_func_hashes):
                local_func_hashes.add(new_func_hash)
                local_funcs.append(new_func)
                funcs_to_handle.append(new_func)
                
    r = ""
    for local_func in local_funcs:
        r += to_python(local_func, context, indent=indent) + "\n"

    indent_str = " " * indent
    r = indent_str + "# VBA Local Function Definitions\n" + r
    return r

jit_cache = {}

def _eval_python(loop, context, params=None, add_boilerplate=False, namespace=None):
    """
    Convert the loop to Python and emulate the loop directly in Python.
    """

    if (not context.do_jit):
        return False

    code_vba = str(loop).replace("\n", "\\n")[:20]
    log.info("Starting JIT emulation of '" + code_vba + "...' ...")
    if (("Execute(" in str(loop)) or
        ("ExecuteGlobal(" in str(loop)) or
        ("Eval(" in str(loop))):
        log.warning("Loop Execute()s dynamic code. Not JIT emulating.")
        return False
    
    code_python = ""
    try:

        tmp_context = Context(context=context, _locals=context.locals, copy_globals=True)
        
        log.info("Generating Python JIT code...")
        code_python = to_python(loop, tmp_context)
        if add_boilerplate:
            var_inits, _ = _loop_vars_to_python(loop, tmp_context, 0)
            func_defns = _called_funcs_to_python(loop, tmp_context, 0)
            code_python = _boilerplate_to_python(0) + "\n" + \
                          func_defns + "\n" + \
                          var_inits + "\n" + \
                          code_python + "\n" + \
                          _check_for_iocs(loop, tmp_context, 0) + "\n" + \
                          _updated_vars_to_python(loop, tmp_context, 0)
        if (log.getEffectiveLevel() == logging.DEBUG):
            safe_print("JIT CODE!!")
            safe_print(code_python)
        log.info("Done generating Python JIT code.")

        if (not context.is_vbscript):
            
            non_ascii_pat = r'"[^"]*[\x7f-\xff][^"]*"'
            non_ascii_pat1 = r'"[^"]*(?:\\x7f|\\x[89a-f][0-9a-f])[^"]*"'
            if ((re.search(non_ascii_pat1, code_python) is not None) or
                (re.search(non_ascii_pat, code_python) is not None)):
                log.warning("VBA code contains Microsoft specific extended ASCII strings. Not JIT emulating.")
                return False

        if (('"Execute", ' in code_python) or
            ('"ExecuteGlobal", ' in code_python) or
            ('"Eval", ' in code_python)):
            log.warning("Functions called by loop Execute() dynamic code. Not JIT emulating.")
            return False
        

        if (code_python in jit_cache):
            var_updates = jit_cache[code_python]
            log.info("Using cached JIT loop results.")
            if (var_updates == "ERROR"):
                log.error("Previous run of Python JIT loop emulation failed. Using fallback emulation for loop.")
                return False

        elif (namespace is None):
            log.info("Evaluating Python JIT code...")
            exec code_python in locals()
        else:
            exec(code_python, namespace)
            var_updates = namespace["var_updates"]
        log.info("Done JIT emulation of '" + code_vba + "...' .")

        jit_cache[code_python] = var_updates
        
        try:
            for updated_var in var_updates.keys():
                if (updated_var == "__shell_code__"):
                    continue
                context.set(updated_var, var_updates[updated_var])
        except (NameError, UnboundLocalError):
            log.warning("No variables set by Python JIT code.")

        import vba_context
        vba_context.shellcode = var_updates["__shell_code__"]

    except NotImplementedError as e:
        log.error("Python JIT emulation of loop failed. " + str(e) + ". Using fallback emulation method for loop...")
        return False

    except Exception as e:

        jit_cache[code_python] = "ERROR"
        
        if ("Infinite Loop" in str(e)):
            log.warning("Detected infinite loop. Terminating loop.")
            return True

        log.error("Python JIT emulation of loop failed. " + str(e) + ". Using fallback emulation method for loop...")
        if (log.getEffectiveLevel() == logging.DEBUG):
            traceback.print_exc(file=sys.stdout)
            safe_print("-*-*-*-*-\n" + code_python + "\n-*-*-*-*-")
        return False

    return True

def eval_arg(arg, context, treat_as_var_name=False):
    """
    evaluate a single argument if it is a VBA_Object, otherwise return its value
    """

    limits_exceeded(throw_error=True)

    if (log.getEffectiveLevel() == logging.DEBUG):
        log.debug("try eval arg: %s (%s, %s, %s)" % (arg, type(arg), isinstance(arg, VBA_Object), treat_as_var_name))

    got_constant_math = is_constant_math(arg)
    
    cached_val = get_cached_value(arg)
    if (cached_val is not None):
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("eval_arg: Got cached value %r = %r" % (arg, cached_val))
        return cached_val
    
    excel_val = _read_from_excel(arg, context)
    if (excel_val is not None):
        if got_constant_math: set_cached_value(arg, excel_val)
        return excel_val

    obj_text_val = _read_from_object_text(arg, context)
    if (obj_text_val is not None):
        if got_constant_math: set_cached_value(arg, obj_text_val)
        return obj_text_val
    
    if ((isinstance(arg, VBA_Object)) or (isinstance(arg, VbaLibraryFunc))):

        if ((".run(" in str(arg).lower()) and (context.contains("run"))):

            if ("MemberAccessExpression" in str(type(arg))):
                arg_evaled = arg.eval(context)
                if got_constant_math: set_cached_value(arg, arg_evaled)
                return arg_evaled

        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("eval_arg: eval as VBA_Object %s" % arg)
        r = arg.eval(context=context)
        
        poss_shape_txt = ""
        if (isinstance(r, VBA_Object) or isinstance(r, str)):
            try:
                poss_shape_txt = str(r)
            except:
                pass
        if ((poss_shape_txt.startswith("Shapes(")) or (poss_shape_txt.startswith("InlineShapes("))):
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("eval_arg: Handling intermediate Shapes() access for " + str(r))
            r = eval_arg(r, context)
            if got_constant_math: set_cached_value(arg, r)
            return r
        
        if got_constant_math: set_cached_value(arg, r)
        return r

    else:
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("eval_arg: not a VBA_Object: %r" % arg)

        if (isinstance(arg, str)):

            try:
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("eval_arg: Try as variable name: %r" % arg)
                r = context.get(arg)
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("eval_arg: Got %r = %r" % (arg, r))
                if got_constant_math: set_cached_value(arg, r)
                return r
            except:
                    
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("eval_arg: Not found as variable name: %r" % arg)
                pass
            else:
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("eval_arg: Do not try as variable name: %r" % arg)

            if ("nodetypedvalue" in arg.lower()):
                try:
                    tmp = arg.lower().replace("nodetypedvalue", "text")
                    if (log.getEffectiveLevel() == logging.DEBUG):
                        log.debug("eval_arg: Try to get as " + tmp + "...")
                    val = context.get(tmp)
    
                    if (log.getEffectiveLevel() == logging.DEBUG):
                        log.debug("eval_arg: Try base64 decode of '" + str(val) + "'...")
                    val_decode = utils.b64_decode(val)
                    if (val_decode is not None):
                        if got_constant_math: set_cached_value(arg, val_decode)
                        return val_decode
                except KeyError:
                    if (log.getEffectiveLevel() == logging.DEBUG):
                        log.debug("eval_arg: Not found as .text.")
                    pass

            elif (".selecteditem" in arg.lower()):
                try:
                    tmp = arg.lower().replace(".selecteditem", ".rapt.value")
                    if (log.getEffectiveLevel() == logging.DEBUG):
                        log.debug("eval_arg: Try to get as " + tmp + "...")
                    val = context.get(tmp)
                    if got_constant_math: set_cached_value(arg, val)
                    return val

                except KeyError:
                    if (log.getEffectiveLevel() == logging.DEBUG):
                        log.debug("eval_arg: Not found as .rapt.value.")
                    pass

            elif ("." in arg.lower()):

                doc_var_val = context.get_doc_var(arg)
                if (doc_var_val is not None):
                    if got_constant_math: set_cached_value(arg, doc_var_val)
                    return doc_var_val

                arg_peeled = arg
                while ("." in arg_peeled):
                
                    curr_var_attempt = arg_peeled.lower()
                    try:
                        if (log.getEffectiveLevel() == logging.DEBUG):
                            log.debug("eval_arg: Try to load as variable " + curr_var_attempt + "...")
                        val = context.get(curr_var_attempt)
                        if (val != str(arg)):
                            if got_constant_math: set_cached_value(arg, val)
                            return val

                    except KeyError:
                        if (log.getEffectiveLevel() == logging.DEBUG):
                            log.debug("eval_arg: Not found as variable")
                        pass

                    arg_peeled = arg_peeled[arg_peeled.index(".") + 1:]

                func_name = arg.lower()
                func_name = func_name[func_name.rindex(".")+1:]
                try:

                    if (log.getEffectiveLevel() == logging.DEBUG):
                        log.debug("eval_arg: Try to run as function '" + func_name + "'...")
                    func = context.get(func_name)
                    r = func
                    import procedures
                    if (isinstance(func, procedures.Function) or
                        isinstance(func, procedures.Sub) or
                        ('simulation_vba.core.vba_library.' in str(type(func)))):
                        r = eval_arg(func, context, treat_as_var_name=True)
                        
                    if (r != func):

                        if got_constant_math: set_cached_value(arg, r)
                        return r

                    if got_constant_math: set_cached_value(arg, arg)
                    return arg

                except KeyError:
                    if (log.getEffectiveLevel() == logging.DEBUG):
                        log.debug("eval_arg: Not found as function")

                except Exception as e:
                    if (log.getEffectiveLevel() == logging.DEBUG):
                        log.debug("eval_arg: Failed. Not a function. " + str(e))
                    traceback.print_exc()

                tmp = arg.lower().strip()
                if (tmp.startswith("activedocument.item(")):

                    prop = tmp.replace("activedocument.item(", "").replace(")", "").replace("'","").strip()

                    if (meta is None):
                        log.error("BuiltInDocumentProperties: Metadata not read.")
                        return ""
                
                    if (not hasattr(meta, prop.lower())):
                        log.error("BuiltInDocumentProperties: Metadata field '" + prop + "' not found.")
                        return ""

                    r = getattr(meta, prop.lower())
                    if (log.getEffectiveLevel() == logging.DEBUG):
                        log.debug("BuiltInDocumentProperties: return %r -> %r" % (prop, r))
                    return r

                if ((tmp.startswith("thisdocument.builtindocumentproperties(")) or
                    (tmp.startswith("activeworkbook.builtindocumentproperties("))):

                    var = tmp.replace("thisdocument.builtindocumentproperties(", "").replace(")", "").replace("'","").strip()
                    var = var.replace("activeworkbook.builtindocumentproperties(", "")
                    val = context.get_doc_var(var)
                    if (val is not None):
                        return val

                    val = context.read_metadata_item(var)
                    if (val is not None):
                        return val
                    
                if (tmp.startswith("activedocument.variables(")):

                    if (log.getEffectiveLevel() == logging.DEBUG):
                        log.debug("eval_arg: handle expression as doc var lookup '" + tmp + "'")
                    var = tmp.replace("activedocument.variables(", "").\
                          replace(")", "").\
                          replace("'","").\
                          replace('"',"").\
                          replace('.value',"").\
                          replace("(", "").\
                          strip()
                    if (log.getEffectiveLevel() == logging.DEBUG):
                        log.debug("eval_arg: look for '" + var + "' as document variable...")
                    val = context.get_doc_var(var)
                    if (val is not None):
                        if (log.getEffectiveLevel() == logging.DEBUG):
                            log.debug("eval_arg: got it as document variable.")
                        return val
                    else:
                        if (log.getEffectiveLevel() == logging.DEBUG):
                            log.debug("eval_arg: did NOT get it as document variable.")

                if (tmp.startswith("activedocument.customdocumentproperties(")):

                    var = tmp.replace("activedocument.customdocumentproperties(", "").\
                          replace(")", "").\
                          replace("'","").\
                          replace('"',"").\
                          replace('.value',"").\
                          replace("(", "").\
                          strip()
                    val = context.get_doc_var(var)
                    if (val is not None):
                        return val
                    
                wild_name = tmp[:tmp.index(".")] + "*"
                for i in range(0, 11):
                    tmp_name = wild_name + str(i)
                    try:
                        val = context.get(tmp_name)
                        if (log.getEffectiveLevel() == logging.DEBUG):
                            log.debug("eval_arg: Found '" + tmp + "' as wild card form variable '" + tmp_name + "'")
                        return val
                    except:
                        pass


        if (treat_as_var_name and (re.match(r"[a-zA-Z_][\w\d]*", str(arg)) is not None)):

            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("eval_arg: return 'NULL'")
            return "NULL"

        if ((str(arg).lower().endswith(".tag")) or
            (str(arg).lower().endswith(".boundvalue")) or
            (str(arg).lower().endswith(".column")) or
            (str(arg).lower().endswith(".caption")) or
            (str(arg).lower().endswith(".groupname")) or
            (str(arg).lower().endswith(".seltext")) or
            (str(arg).lower().endswith(".controltiptext")) or
            (str(arg).lower().endswith(".passwordchar")) or
            (str(arg).lower().endswith(".controlsource")) or
            (str(arg).lower().endswith(".value"))):
            return ""
        
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("eval_arg: return " + str(arg))
        return arg

def eval_args(args, context, treat_as_var_name=False):
    """
    Evaluate a list of arguments if they are VBA_Objects, otherwise return their value as-is.
    Return the list of evaluated arguments.
    """
    try:
        iterator = iter(args)
    except TypeError:
        log.info("eval_args: args not iterable (type=%s), returning as-is", type(args).__name__)
        return args

    got_vba_objects = False
    for arg in args:
        if (isinstance(arg, VBA_Object)):
            got_vba_objects = True
    if (not got_vba_objects):
        log.info("eval_args: no VBA_Objects found in args (types=%s), returning as-is",
                 [type(a).__name__ for a in args])
        return args
    r = list(map(lambda arg: eval_arg(arg, context=context, treat_as_var_name=treat_as_var_name), args))
    log.info("eval_args: evaluated %d args (result types=%s)",
             len(r), [type(x).__name__ for x in r])
    return r

def update_array(old_array, indices, val):
    """
    Add an item to a Python list.
    """

    if (not isinstance(old_array, list)):
        old_array = []

    if (len(indices) == 1):
        
        index = int(indices[0])
        if (index >= len(old_array)):
            old_array.extend([0] * (index - len(old_array) + 1))
        old_array[index] = val

    elif (len(indices) == 2):

        index = int(indices[0])
        index1 = int(indices[1])
        if (index >= len(old_array)):
            for i in range(0, (index - len(old_array) + 1)):
                old_array.append([])
        if (index1 >= len(old_array[index])):
            old_array[index].extend([0] * (index1 - len(old_array[index]) + 1))
        old_array[index][index1] = val
        
    return old_array

def coerce_to_int_list(obj):
    """
    Coerce a constant string VBA object to a list of ASCII codes.
    :param obj: VBA object
    :return: list
    """

    if (isinstance(obj, list)):
        return obj
    
    s = coerce_to_str(obj)

    r = []
    for c in s:
        r.append(ord(c))
    return r

def coerce_to_str(obj, zero_is_null=False):
    """
    Coerce a constant VBA object (integer, Null, etc) to a string.
    :param obj: VBA object
    :return: string
    """

    if ((obj is None) or (obj == "NULL")):
        return ''

    if (zero_is_null and (obj == 0)):
        return ''
    

    if (isinstance(obj, basestring)):

        if (isinstance(obj, unicode)):
            try:
                return obj.encode('utf-8')
            except:
                pass
            
        return obj
    
    if (isinstance(obj, list)):
        r = ""
        bad = False
        for c in obj:

            if (c == 0):
                continue
            try:
                r += chr(c)
            except:

                bad = True
                break

        if (not bad):
            return r

    if (isinstance(obj, dict) and ("value" in obj)):

        return (coerce_to_str(obj["value"]))
        
    try:
        return str(obj)
    except:
        return ''

def coerce_args_to_str(args):
    """
    Coerce a list of arguments to strings.
    Return the list of evaluated arguments.
    """
    return [coerce_to_str(arg) for arg in args]

def coerce_to_int(obj):
    """
    Coerce a constant VBA object (integer, Null, etc) to a int.
    :param obj: VBA object
    :return: int
    """

    if ((obj is None) or (obj == "NULL")):
        return 0

    if (isinstance(obj, int)):
        return obj
    
    if (isinstance(obj, str)):

        if (obj.count('\x00') == len(obj)):
            return 0
        
        obj = obj.replace("\x00", "")
        
        if ("." in obj):
            try:
                obj = float(obj)
                return int(obj)
            except:
                pass
            
        hex_pat = r"&h[0-9a-f]+"
        if (re.match(hex_pat, obj.lower()) is not None):
            return int(obj.lower().replace("&h", "0x"), 16)

    if (isinstance(obj, dict) and ("value" in obj)):

        return (coerce_to_int(obj["value"]))
        
    try:
        return int(obj)
    except ValueError as e:

        log.error("int conversion failed. Returning NULL. " + str(e))
        return 0

def coerce_to_num(obj):
    """
    Coerce a constant VBA object (integer, Null, etc) to a int or float.
    :param obj: VBA object
    :return: int
    """
    if ((obj is None) or (obj == "NULL")):
        return 0

    if ((isinstance(obj, float)) or (isinstance(obj, int))):
        return obj
    
    if (isinstance(obj, str)):

        dumb_pat = r"(?:\d+,)+\d+"
        if (re.match(dumb_pat, obj) is not None):
            obj = obj[:obj.index(",")]
        
        if ("." in obj):
            try:
                obj = float(obj)
                return obj
            except:
                pass

        if (obj.count('\x00') == len(obj)):
            return 0

        hex_pat = r"&h[0-9a-f]+"
        if (re.match(hex_pat, obj.lower()) is not None):
            return int(obj.lower().replace("&h", "0x"), 16)

    if (isinstance(obj, dict) and ("value" in obj)):

        return (coerce_to_num(obj["value"]))
        
    return int(obj)

def coerce_args_to_int(args):
    """
    Coerce a list of arguments to ints.
    Return the list of evaluated arguments.
    """
    return [coerce_to_int(arg) for arg in args]

def coerce_args(orig_args, preferred_type=None):
    """
    Coerce all of the arguments to either str or int based on the most
    common arg type.

    preferred_type = Preferred type to coerce things if possible.
    """

    if (len(orig_args) == 0):
        return orig_args

    args = []
    for arg in orig_args:
        if (arg is None):
            args.append("NULL")
        else:
            args.append(arg)
            
    first_type = None
    have_other_type = False
    all_null = True
    all_types = set()
    for arg in args:

        if (arg == "NULL"):
            continue
        all_null = False
        if (isinstance(arg, str)):
            all_types.add("str")
            if (first_type is None):
                first_type = "str"
            continue
        elif (isinstance(arg, int)):
            all_types.add("int")
            if (first_type is None):
                first_type = "int"
            continue
        else:
            have_other_type = True
            break

    if (all_null):
        first_type = "int"
        
    if (have_other_type):
        return args

    if (first_type is None):
        return args

    if (preferred_type in all_types):
        first_type = preferred_type
    
    if (first_type == "str"):

        new_args = []
        for arg in args:
            if (args == "NULL"):
                new_args.append('')
            else:
                new_args.append(arg)

        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Coerce to str " + str(new_args))
        return coerce_args_to_str(new_args)

    else:

        new_args = []
        for arg in args:
            if (args == "NULL"):
                new_args.append(0)
            else:
                new_args.append(arg)
                
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Coerce to int " + str(new_args))
        return coerce_args_to_int(new_args)
