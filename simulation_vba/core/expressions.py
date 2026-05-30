#!/usr/bin/env python

"""@package expressions VBA Grammar - Expressions
"""

"""
SimulationVBA: VBA Grammar - Expressions

SimulationVBA is a specialized engine to parse, analyze and interpret Microsoft
VBA macros (Visual Basic for Applications), mainly for malware analysis.

Author: Philippe Lagadec - http://www.decalage.info
License: BSD, see source code or documentation

Project Repository:
https://github.com/decalage2/ViperMonkey
"""

# === LICENSE ==================================================================


__version__ = '0.03'


import traceback
import logging
import re
import sys
import os
import array
from hashlib import sha256
import string
import base64
import unidecode

from pyparsing import CaselessKeyword, CaselessLiteral, Combine, FollowedBy, Forward, Group, infixNotation, \
    Keyword, Literal, NotAny, oneOf, OneOrMore, opAssoc, Optional, ParseException, Regex, \
    Suppress, White, Word, ZeroOrMore, delimitedList
import pyparsing

from identifiers import lex_identifier, reserved_identifier, TODO_identifier_or_object_attrib, \
    strict_reserved_keywords, unrestricted_name, enum_val_id, identifier, typed_name, \
    TODO_identifier_or_object_attrib_loose
from lib_functions import StrReverse, Environ, Asc, Chr, chr_, asc, expression, strReverse
from literals import date_string, decimal_literal, float_literal, literal, \
    quoted_string_keep_quotes, integer, quoted_string
from operators import AddSub, And, Concatenation, Eqv, FloorDivision, Mod, MultiDiv, Neg, \
    Not, Or, Power, Sum, Xor
import procedures
from vba_object import eval_arg, eval_args, to_python, coerce_to_int, coerce_to_str, \
    VbaLibraryFunc, VBA_Object
import vba_context
import utils

from logger import log

def _vba_to_python_op(op, is_boolean):
    """
    Convert a VBA boolean operator to a Python boolean operator.
    """
    op_map = {
        "Not" : "not",
        "And" : "and",
        "AndAlso" : "and",
        "Or" : "or",
        "OrElse" : "or",
        "Eqv" : "|eq|",
        "=" : "|eq|",
        ">" : ">",
        "<" : "<",
        ">=" : ">=",
        "=>" : ">=",
        "<=" : "<=",
        "=<" : "<=",
        "<>" : "|neq|",
        "is" : "|eq|"
    }
    if (not is_boolean):
        op_map["Not"] = "~"
        op_map["And"] = "&"
        op_map["AndAlso"] = "&"
        op_map["Or"] = "|"
        op_map["OrElse"] = "|"
    return op_map[op]


file_pointer = Suppress('#') + expression + NotAny("#")
file_pointer.setParseAction(lambda t: "#" + str(t[0]))
file_pointer_loose = (decimal_literal ^ lex_identifier)
file_pointer_loose.setParseAction(lambda t: "#" + str(t[0]))


missed_var_count = {}
class SimpleNameExpression(VBA_Object):
    """
    Identifier referring to a variable within a VBA expression:
    single identifier with no qualification or argument list
    """

    def __init__(self, original_str, location, tokens, name=None):
        super(SimpleNameExpression, self).__init__(original_str, location, tokens)
        if (name is not None):
            self.name = name
        else:
            self.name = tokens.name
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed "%r" as SimpleNameExpression' % self)

    def __repr__(self):
        return '%s' % self.name

    def to_python(self, context, params=None, indent=0):

        if (self.name == "RegExp"):
            return "core.utils.vb_RegExp()"

        value = None
        try:
            value = context.get(self.name)
        except KeyError:
            pass
        if (value == "__ALREADY_SET__"):
            try:
                value = context.get("__ORIG__" + str(self.name))
            except KeyError:
                pass
        
        import vba_library
        if ((self.name.lower() in vba_library.VBA_LIBRARY) and
            (isinstance(value, VbaLibraryFunc)) and
            (value.num_args() == 0)):

            args = "[]"
            r = "core.vba_library.run_function(\"" + str(self.name) + "\", vm_context, " + args + ")"
            return r

        var_name = str(self)
        var_name = utils.fix_python_overlap(var_name)
        
        if (isinstance(value, procedures.Function) and
            (value.min_param_length == 0)):
            return var_name + "()"
        
        return var_name
    
    def eval(self, context, params=None):

        import statements
        
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('try eval variable/function %r' % self.name)
        try:
            value = context.get(self.name)
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug('get variable %r = %r' % (self.name, value))
            if (isinstance(value, procedures.Function) or
                isinstance(value, procedures.Sub) or
                isinstance(value, statements.External_Function) or
                isinstance(value, VbaLibraryFunc)):

                if ((isinstance(value, procedures.Function) or
                     isinstance(value, procedures.Sub)) and
                    (value.min_param_length > 0)):
                    return "NULL"

                if (not context.throttle_logging):
                    log.info("calling Function: " + str(value) + "()")
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug('evaluating function %r' % value)
                value = value.eval(context)
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug('evaluated function %r = %r' % (self.name, value))
            return value
        except KeyError:

            var_name = str(self.name)
            global missed_var_count
            if (var_name not in missed_var_count.keys()):
                missed_var_count[var_name] = 0
            missed_var_count[var_name] += 1
            if (missed_var_count[var_name] < 20):
                log.warning('Variable %r not found' % self.name)
            if (self.name.startswith("%") and self.name.endswith("%")):
                return self.name.upper()
            return "NULL"


simple_name_expression = Optional(CaselessKeyword("ByVal").suppress()) + \
                         (TODO_identifier_or_object_attrib('name') | enum_val_id('name'))
simple_name_expression.setParseAction(SimpleNameExpression)

unrestricted_name_expression = unrestricted_name('name')
unrestricted_name_expression.setParseAction(SimpleNameExpression)

placeholder = Keyword("***PLACEHOLDER***")
placeholder.setParseAction(lambda t: str(t[0]))


class InstanceExpression(VBA_Object):
    """
    An instance expression consists of the keyword "Me".
    It represents the current instance of the type defined by the
    enclosing class module and has this type as its value type.
    """

    def __init__(self, original_str, location, tokens):
        super(InstanceExpression, self).__init__(original_str, location, tokens)
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed %r as InstanceExpression' % self)

    def __repr__(self):
        return 'Me'

    def eval(self, context, params=None):
        raise NotImplementedError




instance_expression = CaselessKeyword('Me').suppress()
instance_expression.setParseAction(InstanceExpression)





class MemberAccessExpression(VBA_Object):
    """
    Handle member access expressions.
    """

    def __init__(self, original_str, location, tokens, raw_fields=None):

        self.is_loop = False
        if (raw_fields is not None):
            self.lhs = raw_fields[0]
            self.rhs = raw_fields[1]
            self.rhs1 = raw_fields[2]
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug('Manually created MemberAccessExpression %r' % self)

        else:
            super(MemberAccessExpression, self).__init__(original_str, location, tokens)
            tokens = tokens[0][0]
            self.rhs = tokens[1:]
            self.lhs = tokens.lhs
            if ((isinstance(self.lhs, list) or isinstance(self.lhs, pyparsing.ParseResults)) and (len(self.lhs) > 0)):
                self.lhs = self.lhs[0]
            self.rhs1 = ""
            if (hasattr(tokens, "rhs1")):
                self.rhs1 = tokens.rhs1
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug('parsed %r as MemberAccessExpression' % self)

    def __repr__(self):
        r = str(self.lhs)
        for t in self.rhs:
            r += "." + str(t)
        if (len(self.rhs1) > 0):
            r += "." + str(self.rhs1)
        return r

    def _to_python_handle_listbox_list(self, context, indent):
        """
        Handle List() object method calls like foo.List(bar).
        foo is (currently) a ListBox object.
        """

        func = self.rhs
        if (isinstance(func, list)):
            func = func[-1]
            
        if ((not isinstance(func, Function_Call)) or
            (func.name != "List")):

            if (str(func) == "List"):
                return str(self.lhs)

            return None

        the_list = self._get_with_prefix_value(context)
        if (the_list is None):
            the_list = context.get(str(self.lhs))
            if (the_list == "__ALREADY_SET__"):
                the_list = context.get("__ORIG__" + str(self.lhs))
        if ((the_list is None) or (not isinstance(the_list, list))):
            return None

        r = str(the_list) + "[coerce_to_int(" + to_python(func.params[0], context) + ")]"
        return r
        
    def _get_with_prefix_value(self, context):
        """
        Get the value of the With prefix. None is returned if there is no
        With prefix.
        """
        with_value = None
        if ((context.with_prefix_raw is not None) and
            (context.contains(str(context.with_prefix_raw)))):
            with_value = context.get(str(context.with_prefix_raw))
            if (with_value == "__ALREADY_SET__"):

                with_value = context.get("__ORIG__" + str(context.with_prefix_raw))

        return with_value
    
    def _to_python_handle_add(self, context, indent):
        """
        Handle Add() object method calls like foo.Add(bar, baz). 
        foo is (currently) a Scripting.Dictionary object.
        """

        with_dict = self._get_with_prefix_value(context)
        if ((with_dict is None) or (not isinstance(with_dict, dict))):
            return None

        expr_str = str(self)
        if (".Add(" not in expr_str):
            return None
            
        tmp_var = SimpleNameExpression(None, None, None, name=str(context.with_prefix_raw))
        new_add = Function_Call(None, None, None, old_call=self.rhs[0])
        tmp = [tmp_var]
        for p in new_add.params:
            tmp.append(p)
        new_add.params = tmp
        indent_str = " " * indent
        r = indent_str + to_python(new_add, context)        
        return r

    def _convert_nested_methods_to_func_call(self, context):
        """
        Given a member access expression like foo(1).bar(2).baz(3)
        return (conceptually) baz(3, bar(2, foo(1))).
        """

        import vba_library
        
        obj_stack = []
        obj_stack.append(self.lhs)
        if isinstance(self.rhs, list):
            for obj in self.rhs:
                obj_stack.append(obj)
        else:
            obj_stack.append(self.rhs)

        prev_func = None
        curr_func = None
        res_func = None
        while (len(obj_stack) > 0):

            curr_obj = obj_stack.pop()
            obj_name = None
            curr_func = curr_obj
            if isinstance(curr_obj, SimpleNameExpression):
                obj_name = str(curr_obj)
                curr_func = function_call.parseString(obj_name + "()", parseAll=True)[0]
                curr_func.params = []
            elif isinstance(curr_obj, Function_Call):
                obj_name = str(curr_obj.name)
                curr_func = Function_Call(None, None, None, old_call=curr_obj)
            else:
                return None
                
            if (obj_name.lower() not in vba_library.VBA_LIBRARY):

                if ((context.contains(obj_name)) and (len(obj_stack) == 0)):
                    curr_func = context.get(obj_name)
                else:
                    return None

            if (prev_func is not None):
                prev_func.params.append(curr_func)
            else:
                res_func = curr_func
            prev_func = curr_func
            
        return res_func
        
    def _to_python_nested_methods(self, context, indent):
        """
        Given a member access expression like foo(1).bar(2).baz(3)
        return (conceptually) baz(3, bar(2, foo(1))), but in Python.
        """
            
        res_func = self._convert_nested_methods_to_func_call(context)
        if (res_func is None):
            return None
        r = to_python(res_func, context)
        return r

    def _to_python_handle_regex(self, context, indent):
        """Handle RegEx() object method calls like Replace() and Test().

        """
        if (len(self.rhs) == 0):
            return None

        raw_last_func = str(self.rhs[-1]).replace("('", "(").replace("')", ")").strip()
        if (not ((raw_last_func.startswith("Test(")) or
                 (raw_last_func.startswith("Replace(")) or
                 (raw_last_func == "Global") or
                 (raw_last_func == "Pattern"))):
            return None
            
        exp_str = str(self)
        call_str = raw_last_func
        if (isinstance(self.rhs[-1], Function_Call)):
            the_call = self.rhs[-1]
            call_str = str(the_call.name) + "("
            first = True
            for p in the_call.params:
                if (not first):
                    call_str += ", "
                first = False
                call_str += to_python(p, context)
            call_str += ")"
        r = exp_str[:exp_str.rindex(".")] + "." + call_str
        return r
        
    def to_python(self, context, params=None, indent=0):

        add_code = self._to_python_handle_add(context, indent)
        if (add_code is not None):
            return add_code

        add_code = self._to_python_handle_listbox_list(context, indent)
        if (add_code is not None):
            return add_code

        add_code = self._to_python_handle_regex(context, indent)
        if (add_code is not None):
            return add_code

        add_code = self._to_python_nested_methods(context, indent)
        if (add_code is not None):
            return add_code
        
        if (len(self.rhs) > 0):

            raw_last_func = str(self.rhs[-1]).replace("('", "(").replace("')", ")").strip()
            if (raw_last_func.startswith("SpecialCells(")):

                new_special_cells = Function_Call(None, None, None, old_call=self.rhs[-1])
                cells = self._eval_cell_range(context, just_expr=True)
                tmp = [cells, new_special_cells.params[0]]
                new_special_cells.params = tmp
                
                r = to_python(new_special_cells, context, params)
                return r

            last_rhs = to_python(self.rhs[-1], context, params)
            
            if (last_rhs.lower() == '"name"'):
                lhs_str = to_python(self.lhs, context, params)
                if (lhs_str.startswith('"')):
                    lhs_str = lhs_str[1:]
                if (lhs_str.endswith('"')):
                    lhs_str = lhs_str[:-1]
                last_rhs = lhs_str + "['name']"

            if ((("(" not in last_rhs) and ("[" not in last_rhs)) or
                (last_rhs.lower().startswith("address("))):

                if ((last_rhs.lower() != "value") and
                    (last_rhs.lower() != "row") and
                    (not last_rhs.lower().startswith("address(")) and
                    (last_rhs.lower() != "col") and
                    context.contains(str(self))):

                    r = str(self).replace(".", "")
                    return r

                if (last_rhs.lower().startswith("address(")):
                    last_rhs = "index"
                
                lhs_str = to_python(self.lhs, context, params)
                last_rhs = "core.vba_library.member_access(" + lhs_str + ", \"" + last_rhs + "\", globals())"

                pat = r"(core\.vba_library\.member_access\(core\.vba_library\.run_function\(\"Range\", vm_context, \[)(.+)(\]\), \"(?:Column|Row)\",)"
                last_rhs = re.sub(pat, r"\1\2, True\3", last_rhs)
                
            return last_rhs
        
        return ""
    
    def _handle_indexed_pages_access(self, context):
        """
        Handle getting the caption of a Page object referenced via index.
        """

        page_pat = r".+\.Pages\('(\d+)'\)\.Caption"
        index = re.findall(page_pat, str(self))
        if (len(index) == 0):
            return None
        index = int(index[0]) + 1

        var_name = "Page" + str(index) + ".Caption"
        if (context.contains(var_name)):
            return context.get(var_name)
        return None
    
    def _handle_table_cell(self, context):
        """
        Handle reading a value from a table cell.
        """

        pat = r"\w+\.Tables\(\s*'(\w+)'\s*\)\.Cell\(\s*'(\w+)\s*,\s*(\w+)'\s*\).*"
        indices = re.findall(pat, str(self))
        if (len(indices) == 0):
            return None
        indices = indices[0]

        table_index = None
        try:

            obj = expression.parseString(indices[0], parseAll=True)[0]
            
            table_index = obj
            if (isinstance(table_index, VBA_Object)):
                table_index = table_index.eval(context)
            if (isinstance(table_index, str)):
                table_index = int(table_index.replace("'", ""))
            table_index -= 1

        except ParseException:
            log.error("Parse error. Cannot evaluate '" + indices[0] + "'")
            return None
        except Exception as e:
            log.error("Comment index '" + str(indices[0]) + "' not int. " + str(e))
            return None
        cell_index_row = None
        try:

            obj = expression.parseString(indices[1], parseAll=True)[0]
            
            cell_index_row = obj
            if (isinstance(cell_index_row, VBA_Object)):
                cell_index_row = cell_index_row.eval(context)
            if (isinstance(cell_index_row, str)):
                cell_index_row = int(cell_index_row.replace("'", ""))
            cell_index_row -= 1

        except ParseException:
            log.error("Parse error. Cannot evaluate '" + indices[1] + "'")
            return None
        except Exception as e:
            log.error("Comment index '" + str(indices[1]) + "' not int. " + str(e))
            return None
        cell_index_col = None
        try:

            obj = expression.parseString(indices[2], parseAll=True)[0]
            
            cell_index_col = obj
            if (isinstance(cell_index_col, VBA_Object)):
                cell_index_col = cell_index_col.eval(context)
            if (isinstance(cell_index_row, str)):
                cell_index_col = int(cell_index_row.replace("'", ""))
            cell_index_col -= 1

        except ParseException:
            log.error("Parse error. Cannot evaluate '" + indices[2] + "'")
            return None
        except Exception as e:
            log.error("Comment index '" + str(indices[2]) + "' not int. " + str(e))
            return None

        tables = context.get("__DOC_TABLE_CONTENTS__")
        if (table_index >= len(tables)):
            return None
        table = tables[table_index]
        if (cell_index_row >= len(table)):
            return None
        row = table[cell_index_row]
        if (cell_index_col >= len(row)):
            return None
        cell = str(row[cell_index_col]) + "  "
        return cell
    
    def _handle_paragraphs(self, context):
        """
        Handle references to the .Paragraphs field of the current doc.
        """

        if (str(self).lower().endswith(".paragraphs")):
            return context.get("ActiveDocument.Paragraphs".lower())

        if (len(self.rhs) == 0):
            return None
        first_rhs = self.rhs[0]
        if (not str(first_rhs).startswith("Paragraphs('")):
            return None

        r = eval_arg(first_rhs, context)
        return r

    def _handle_comments(self, context):
        """
        Handle references to the .Comments field of the current doc.
        """

        me_str = str(self)
        if (".comments" not in me_str.lower()):
            return None

        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Try _handle_comments() eval of " + me_str)
        if (me_str.lower().endswith(".comments")):
            return context.get("ActiveDocument.Comments".lower())

        ref_pat = r".Comments\(\s*(.+)\s*\)"
        ids = re.findall(ref_pat, me_str)
        if (len(ids) == 0):
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("No comment index found.")
            return None


        index = None
        try:

            obj = expression.parseString(ids[0], parseAll=True)[0]
            
            index = obj
            if (isinstance(index, VBA_Object)):
                index = index.eval(context)
            index = int(index.replace("'", "")) - 1

        except ParseException:
            log.error("Parse error. Cannot evaluate '" + str(ids[0]) + "'")
            return None
        except Exception as e:
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("Comment index '" + str(ids[0]) + "' not int. " + str(e))
            return None

        comments = context.get("ActiveDocument.Comments".lower())
        if (index >= len(comments)):
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("Comment index " + str(ids[0]) + " out of range")
            return None
        return comments[index]
            
    def _handle_count(self, context, curr_item):
        """
        Handle references to the .Count field of the current item.
        """
        if ((".count" in str(self).lower()) and (isinstance(curr_item, list))):
            return len(curr_item)

    def _handle_item(self, context, curr_item):
        """
        Handle accessing a list item.
        """

        if (not isinstance(curr_item, list)):
            return None

        if (".item(" not in str(self).lower()):
            return None

        tmp_rhs = self.rhs
        if (isinstance(tmp_rhs, list) and (len(tmp_rhs) > 0)):
            tmp_rhs = tmp_rhs[0]
        if ((not isinstance(tmp_rhs, Function_Call)) or
            (tmp_rhs.name != "Item")):
            return None
        index = eval_arg(tmp_rhs.params[0], context)
        if (not isinstance(index, int)):
            return None
        if (index >= len(curr_item)):
            return "NULL"

        return curr_item[index]
        
    def _handle_oslanguage(self, context):
        """
        Handle references to the OSlanguage field.
        """
        if (str(self).lower().endswith(".oslanguage")):
            return context.get("oslanguage")
    
    def _handle_application_run(self, context):
        """
        Handle functions called with Application.Run()
        """

        if ((not str(self).startswith("Application.Run(")) and
            (not str(self).lower().startswith("thisdocument.run("))):
            return None
        
        if (len(self.rhs[0].params) == 0):
            return None

        func_args = None
        if (isinstance(self.rhs[0].params[0], Function_Call)):
            func_name = self.rhs[0].params[0].name
            func_args = self.rhs[0].params[0].params

        else:
            func_name = str(self.rhs[0].params[0])
            func_args = []
            if (len(self.rhs[0].params) > 1):
                func_args = self.rhs[0].params[1:]
            func_args = eval_args(func_args, context)

        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Try indirect run of function '" + func_name + "'")
        r = "NULL"
        try:

            s = func_name
            while ((isinstance(s, str)) or (isinstance(s, SimpleNameExpression))):
                s = context.get(str(s))
                if (isinstance(s, procedures.Function) or
                    isinstance(s, procedures.Sub) or
                    isinstance(s, VbaLibraryFunc)):
                    s = s.eval(context=context, params=func_args)
                    r = s

            if ((str(self).lower().startswith("thisdocument.run(")) and (r != "NULL")):
                context.report_action('Execute Command', r, 'ThisDocument.Run', strip_null_bytes=True)
            return r
        
        except KeyError:
            if (r != "NULL"):
                return r
            return None

    def _handle_set_clipboard(self, context):
        """
        Handle calls like objHTML.ParentWindow.clipboardData.setData(...).
        """

        if (".setdata(" not in str(self).lower()):
            return None
        
        func = self.rhs[-1]
        if (not isinstance(func, Function_Call)):
            return None
        if (len(func.params) < 2):
            return None
        val = func.params[1]
        val = str(eval_arg(val, context))

        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Save clipboard text '" + val + "'")
        context.set("** CLIPBOARD **", val, force_global=True)
        return True

    def _handle_get_clipboard(self, context):
        """
        Handle calls like objHTML.ParentWindow.clipboardData.getData(...).
        """

        if (".getdata(" not in str(self).lower()):
            return None
        
        if (context.contains("** CLIPBOARD **")):
            return context.get("** CLIPBOARD **")
        return None
        
    def _handle_docprops_read(self, context):
        """
        Handle data reads with ActiveDocument.BuiltInDocumentProperties(...).
        """

        if ((not str(self).startswith("ActiveDocument.BuiltInDocumentProperties(")) and
            (not str(self).startswith("ThisDocument.BuiltInDocumentProperties("))):
            return None

        if (len(self.rhs[0].params) == 0):
            return None
        field_name = eval_arg(self.rhs[0].params[0], context)

        r = context.get_doc_var(field_name)
        if (r is not None):
            return r

        return context.read_metadata_item(field_name)

    def _handle_control_read(self, context):
        """
        Handle data reads with StreamName.Controls(...).Value.
        """

        pat = r".+\.Controls\(\s*'([^']+)'\s*\)(?:\.Value)?"
        my_text = str(self)
        if (re.match(pat, str(self)) is None):
            return None

        list_name = my_text.replace(".Value", "")
        list_name = list_name[:list_name.rindex("(")]
        list_vals = None
        if (list_name.endswith("('")):
            list_name = list_name[:-2]
        try:
            list_vals = context.get(list_name)
        except KeyError:
            return None

        index = re.findall(pat, my_text)[0]

        try:

            obj = expression.parseString(index, parseAll=True)[0]
            
            index = obj
            if (isinstance(index, VBA_Object)):
                index = index.eval(context)
            index = int(index)

        except ParseException:
            log.error("Parse error. Cannot evaluate '" + index + "'")
            return None
        except:
            return None

        if (index < len(list_vals)):
            return list_vals[index]
        return None

    def _handle_docvars_read(self, context):
        """
        Handle data reads from a document variable.
        """

        tmp = self.__repr__().lower()
        if (tmp.startswith("activedocument.variables(")):
            return eval_arg(self.__repr__(), context)
        
        if ("(" in tmp):
            tmp = tmp[:tmp.rindex("(")]
        val = context.get_doc_var(tmp, search_wildcard=False)
        
        if (("(" in str(self.lhs)) and
            (isinstance(self.lhs, Function_Call)) and
            (val is not None) and
            (isinstance(val, list))):

            if (len(self.lhs.params) > 0):
                index = eval_arg(self.lhs.params[0], context)
                if ((isinstance(index, int)) and (index < len(val))):
                    if ((index >= len(val)) or (index < 0)):
                        return None
                    val = val[index]

        rhs = str(self.rhs).lower().replace("'", "").replace("[", "").replace("]", "")
        if ((isinstance(val, dict)) and (rhs in val)):
            val = val[rhs]

        if (isinstance(val, procedures.Function) or
            isinstance(val, procedures.Sub) or
            isinstance(val, VbaLibraryFunc)):
            return None
            
        return val

    def _handle_text_file_read(self, context):
        """
        Handle OpenTextFile(...).ReadAll() calls.
        """

        tmp = self.__repr__().lower()
        if (("opentextfile(" not in tmp) or ("readall" not in tmp)):
            return None

        if (len(self.rhs) < 2):
            return None
        read_call = self.rhs[-2]
        if (not isinstance(read_call, Function_Call)):
            return None
        read_file = str(eval_arg(read_call.params[0], context))

        try:
            f = open(read_file, 'rb')
            r = f.read()
            f.close()
            return r
        except Exception as e:

            if (read_file.startswith("C:\\")):
                read_file = read_file.replace("C:\\", "")

            try:
                f = open(read_file, 'rb')
                r = f.read()
                f.close()
                return r
            except Exception as e:
                
                log.error("ReadAll('" + read_file + "') failed. " + str(e))
                return None

    def _handle_docvar_value(self, lhs, rhs):
        """
        Handle reading .Name and .Value fields from doc vars.
        """
        
        if ((isinstance(rhs, list)) and (len(rhs) > 0)):
            rhs = rhs[0]
        rhs = str(rhs).strip()
            
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("_handle_docvar_value(): lhs = " + str(lhs) + ", rhs = '" + str(rhs) + "'")
            
        if (not isinstance(lhs, tuple)):
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("_handle_docvar_value(): LHS not tuple")
            return None
        if (len(lhs) < 2):
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("_handle_docvar_value(): LHS not 2 element tuple")
            return None
        
        if (rhs == "Name"):
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("_handle_docvar_value(): return name = '" + str(lhs[0]) + "'")
            return lhs[0]

        if (rhs == "Value"):
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("_handle_docvar_value(): return value = '" + str(lhs[1]) + "'")
            return lhs[1]

        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("_handle_docvar_value(): not getting name or value of '" + str(self) + "'")
        return None

    def _handle_file_close(self, context, lhs, rhs):
        """
        Handle close of file object foo like foo.Close().
        """

        if ((isinstance(rhs, list)) and (len(rhs) > 0)):
            rhs = rhs[0]
        if (str(rhs) != "Close"):
            return None
        from vba_library import Close
        file_close = Close()
            
        return file_close.eval(context, [str(lhs)])
    
    def _handle_replace(self, context, lhs, rhs):
        """
        Handle string replaces of the form foo.Replace(bar, baz). foo is a RegExp object.
        """

        if ((isinstance(rhs, list)) and (len(rhs) > 0)):
            rhs = rhs[0]
        if (not isinstance(rhs, Function_Call)):
            return None
        if (rhs.name != "Replace"):
            return None
        if (not str(lhs).lower().endswith("regexp")):
            return None

        pat_name = str(self.lhs) + ".pattern"
        if (not context.contains(pat_name)):
            pat_name = ".pattern"
            if (not context.contains(pat_name)):
                return None
        repl = context.get(pat_name)
        
        new_replace = Function_Call(None, None, None, old_call=rhs)
        tmp = [new_replace.params[0]]
        tmp.append(repl)
        tmp.append(new_replace.params[1])
        tmp.append("<-- USE REGEX -->")
        new_replace.params = tmp
        
        r = new_replace.eval(context)
        return r

    def _handle_add(self, context, lhs, rhs):
        """
        Handle Add() object method calls like foo.Add(bar, baz). 
        foo is (currently) a Scripting.Dictionary object.
        """

        if (isinstance(lhs, str) and
            lhs.startswith("{") and
            lhs.endswith("}")):
            try:
                lhs = eval(lhs)
            except SyntaxError:
                pass

        if ((context.with_prefix_raw is not None) and
            (context.contains(str(context.with_prefix_raw)))):
            lhs = context.get(str(context.with_prefix_raw))
            
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("_handle_add(): lhs = " + str(lhs) + ", rhs = " + str(rhs))
        if ((isinstance(rhs, list)) and (len(rhs) > 0)):
            rhs = rhs[0]
        if (not isinstance(rhs, Function_Call)):
            return None
        if (rhs.name != "Add"):
            return None
        if (not isinstance(lhs, dict)):
            return None

        new_add = Function_Call(None, None, None, old_call=rhs)
        tmp = [lhs]
        for p in new_add.params:
            tmp.append(p)
        new_add.params = tmp
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Add() func = " + str(new_add))
        
        new_dict = new_add.eval(context)

        if (context.with_prefix_raw is not None):
            context.set(str(context.with_prefix_raw), new_dict, do_with_prefix=False)
        if (context.contains(str(self.lhs))):
            context.set(str(self.lhs), new_dict, do_with_prefix=False)

        return new_dict

    def _handle_listbox_list(self, context, lhs, rhs):
        """
        Handle List() object method calls like foo.List(bar).
        foo is (currently) a ListBox object.
        """

        if (isinstance(lhs, str) and
            lhs.startswith("[") and
            lhs.endswith("]")):
            try:
                lhs = eval(lhs)
            except SyntaxError:
                pass

        if ((context.with_prefix_raw is not None) and
            (context.contains(str(context.with_prefix_raw)))):
            lhs = context.get(str(context.with_prefix_raw))
            
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("_handle_listbox_list(): lhs = " + str(lhs) + ", rhs = " + str(rhs))
        if ((isinstance(rhs, list)) and (len(rhs) > 0)):
            rhs = rhs[0]
        if ((not isinstance(rhs, Function_Call)) or
            (rhs.name != "List")):

            lhs_str = str(self.lhs)
            if ((str(rhs) == "List") and (context.contains(lhs_str))):                
                return context.get(lhs_str)

            return None
        if (not isinstance(lhs, list)):
            return None

        index = eval_arg(rhs.params[0], context)

        if ((not isinstance(index, int)) or
            (index < 0) or
            (index > (len(lhs) - 1))):
            return None
        return lhs[index]
    
    def _handle_listbox_additem(self, context, lhs, rhs):
        """
        Handle AddItem() object method calls like foo.AddItem(bar).
        foo is (currently) a ListBox object.
        """

        if (isinstance(lhs, str) and
            lhs.startswith("[") and
            lhs.endswith("]")):
            try:
                lhs = eval(lhs)
            except SyntaxError:
                pass

        if ((context.with_prefix_raw is not None) and
            (context.contains(str(context.with_prefix_raw)))):
            lhs = context.get(str(context.with_prefix_raw))
            
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("_handle_listbox_additem(): lhs = " + str(lhs) + ", rhs = " + str(rhs))
        if ((isinstance(rhs, list)) and (len(rhs) > 0)):
            rhs = rhs[0]
        if (not isinstance(rhs, Function_Call)):
            return None
        if (rhs.name != "AddItem"):
            return None

        if ((lhs is None) or (lhs == "NULL")):
            lhs = []
        if (not isinstance(lhs, list)):
            return None

        new_add = Function_Call(None, None, None, old_call=rhs)
        tmp = [lhs]
        for p in new_add.params:
            tmp.append(p)
        new_add.params = tmp
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("AddItem() func = " + str(new_add))
        
        new_list = new_add.eval(context)

        if (context.with_prefix_raw is not None):
            context.set(str(context.with_prefix_raw), new_list, do_with_prefix=False)
        context.set(str(self.lhs), new_list, do_with_prefix=False, force_global=True)

        return new_list

    def _handle_exists(self, context, lhs, rhs):
        """
        Handle Exists() object method calls like foo.Exists(bar). 
        foo is (currently) a Scripting.Dictionary object.
        """

        if (isinstance(lhs, str) and
            lhs.startswith("{") and
            lhs.endswith("}")):
            try:
                lhs = eval(lhs)
            except SyntaxError:
                pass

        if ((context.with_prefix_raw is not None) and
            (context.contains(str(context.with_prefix_raw)))):
            lhs = context.get(str(context.with_prefix_raw))
            
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("_handle_exists(): lhs = " + str(lhs) + ", rhs = " + str(rhs))
        if ((isinstance(rhs, list)) and (len(rhs) > 0)):
            rhs = rhs[0]
        if (not isinstance(rhs, Function_Call)):
            return None
        if (rhs.name != "Exists"):
            return None
        if (not isinstance(lhs, dict)):
            return None

        new_exists = Function_Call(None, None, None, old_call=rhs)
        tmp = [lhs]
        for p in new_exists.params:
            tmp.append(p)
        new_exists.params = tmp
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Exists() func = " + str(new_exists))
        
        r = new_exists.eval(context)
        return r

    def _handle_adodb_writes(self, lhs_orig, lhs, rhs, context):
        """
        Handle expressions like "foo.Write(...)" where foo = "ADODB.Stream".
        """

        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("_handle_adodb_writes(): lhs_orig = " + str(lhs_orig) + ", lhs = " + str(lhs) + ", rhs = " + str(rhs))
        rhs_str = str(rhs).strip()
        if (("write(" not in rhs_str.lower()) and ("writetext(" not in rhs_str.lower())):
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("Not a Write() call.")
            return False
        
        lhs_str = str(lhs)
        if ((lhs_str.lower() != "ADODB.Stream".lower()) and
            (not lhs_str.lower().startswith("cdo.message."))):

            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug(lhs_str + " is not an ADODB.Stream")
            if ((not isinstance(self.rhs, list)) or (len(self.rhs) < 2)):
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("Done (1).")
                return False

            for field in self.rhs[:-1]:
                lhs_orig += "." + str(field)

            if (str(eval_arg(lhs_orig, context)) == str(lhs_orig)):
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("Done (2).")
                return False
        
        txt = None
        rhs_val = eval_arg(rhs.params[0], context)
        try:
            txt = str(rhs_val)
        except UnicodeEncodeError:
            txt = ''.join(filter(lambda x:x in string.printable, rhs_val))

        if (".GetEncodedContentStream.WriteText(" in str(self)):

            type_name = lhs_str[:lhs_str.index("GetEncodedContentStream")] + "ContentTransferEncoding"
            typ = None
            try:
                typ = context.get(type_name)
            except KeyError:
                pass
            if (typ.lower() == "base64"):                
                decoded = utils.b64_decode(txt)
                if (decoded is not None):
                    txt = decoded
            

        var_name = str(lhs_orig) + ".ReadText"
        if (not context.contains(var_name)):
            context.set(var_name, "", force_global=True)
        final_txt = context.get(var_name) + txt
        context.set(var_name, final_txt, force_global=True)

        var_name = "ADODB.Stream.ReadText"
        if (not context.contains(var_name)):
            context.set(var_name, "", force_global=True)
        final_txt = context.get(var_name) + txt
        context.set(var_name, final_txt, force_global=True)
        
        return True

    def _handle_excel_read(self, context, rhs):
        """
        Handle Excel reads like worksheets.cells(1,2).
        """

        return None

    def _handle_0_arg_call(self, context, rhs=None):
        """
        Handle calls to 0 argument functions.
        """

        if (rhs is None):
            if (len(self.rhs1) > 0):
                rhs = self.rhs1
            else:
                rhs = self.rhs[len(self.rhs) - 1]
        
        if (((not isinstance(rhs, str)) and (not isinstance(rhs, SimpleNameExpression))) or
            (not context.contains(str(rhs)))):
            return None
        func = context.get(str(rhs))
        if ((not isinstance(func, procedures.Sub)) and
            (not isinstance(func, procedures.Function)) and
            (not isinstance(func, VbaLibraryFunc))):
            return None

        num_params = 100
        if (hasattr(func, "params")):
            num_params = len(func.params)
        if (hasattr(func, "num_args")):
            num_params = func.num_args()
        if (num_params > 0):
            return None

        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('evaluating function %r' % func)
        r = func.eval(context)
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('evaluated function %r = %r' % (str(func), r))
        return r

    def _handle_loadxml(self, context, load_xml_result):
        """
        Handle things like kXMeYOrbWn.LoadXML(VuvMyknuKxHFAK). This is 
        specifically targeting BASE64 XML elements used for base64 decoding.
        """

        memb_str = str(self)
        if (".LoadXML(" not in memb_str):
            return False

        var_name = memb_str[:memb_str.index(".")] + ".text"
        context.set(var_name, load_xml_result)
        var_name = memb_str[:memb_str.index(".")] + ".selectsinglenode('b64decode').text"
        context.set(var_name, load_xml_result)
        var_name = memb_str[:memb_str.index(".")] + ".nodetypedvalue"
        context.set(var_name, load_xml_result)
        var_name = memb_str[:memb_str.index(".")] + ".selectsinglenode('b64decode').nodetypedvalue"
        context.set(var_name, load_xml_result)
        
        return True

    def _handle_savetofile(self, context, filename):
        """
        Handle things like TvfSKqpfj.SaveToFile oFyFLFCozNUyE, 2.
        """

        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("_handle_savetofile(): filename = " + str(filename) + ", self = " + str(self))
        memb_str = str(self)
        if (".savetofile(" not in memb_str.lower()):
            return False

        var_name = memb_str[:memb_str.lower().index(".savetofile")] + ".ReadText"
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("var_name = " + var_name)
        val = None
        try:
            val = context.get(var_name)
        except KeyError:
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("var_name '" + var_name + "' not found.")

        if (val == None):
            var_name = memb_str[:memb_str.lower().index(".savetofile")] + ".text"
            if ("CreateObject(" in var_name):
                var_name = var_name[:var_name.index("CreateObject(")] + ".text"
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("var_name 1 = " + var_name)
            val = None
            try:
                val = context.get(var_name)
            except KeyError:
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("var_name 1 '" + var_name + "' not found.")
                return False
            

        out_dir = vba_context.out_dir
        if (not os.path.isdir(out_dir)):
            os.makedirs(out_dir)
        
        if ("/" in filename):
            filename = filename[filename.rindex("/") + 1:]
        if ("\\" in filename):
            filename = filename[filename.rindex("\\") + 1:]        
        fname = out_dir + "/" + filename
        fname = fname.replace("\x00", "").replace("..", "")
        fname = ''.join(filter(lambda x:x in string.printable, fname))
        fname = re.sub(r"[^ -~]", "__", fname)
        try:

            f = open(fname, 'wb')
            f.write(val)
            f.close()
            context.report_action('Write File', filename, 'ADODB.Stream SaveToFile()', strip_null_bytes=True)

            raw_data = array.array('B', val).tostring()
            h = sha256()
            h.update(raw_data)
            file_hash = h.hexdigest()
            context.report_action("Dropped File Hash", file_hash, 'File Name: ' + filename)

            context.set(var_name, "")
            
        except Exception as e:
            log.error("Writing " + fname + " failed. " + str(e))
            return False
        
        return True

    def _handle_path_access(self):
        """
        See if this is accessing the Path field of a file/folder object.
        """
        tmp = str(self.rhs).lower().replace("'", "").replace("[", "").replace("]", "")
        if (tmp == "path"):

            return "C:\\Users\\admin\\"

    def _handle_indexed_form_access(self, context):
        """
        See if this is accessing a control in a form by index.
        """

        self_str = str(self)
        if (".Controls(" not in self_str):
            return None

        controls_str = (self_str[:self_str.index(".Controls(")] + ".Controls").lower()
        control_vals = None
        try:
            control_vals = context.get(controls_str)
        except KeyError:

            return None

        pat = r".+\.Controls\(\s*'([^']+)'\s*\)"
        vals = re.findall(pat, self_str)
        if (len(vals) == 0):
            return None
        index = vals[0]

        try:

            obj = expression.parseString(index, parseAll=True)[0]
            
            index = obj
            if (isinstance(index, VBA_Object)):
                index = index.eval(context)
            index = int(index)

        except ParseException:
            log.error("Parse error. Cannot evaluate '" + index + "'")
            return None
        except:
            return None

        if (index >= len(control_vals)):
            return None

        control_val = control_vals[index]
        if ((not isinstance(self.rhs, list)) or (len(self.rhs) < 2)):
            return None
        field = str(self.rhs[1]).lower()
        if (field not in control_val):
            return None
        r = control_val[field]
        return r

    def _handle_regex_execute(self, context, tmp_lhs):
        """
        Handle application of a RegEx object to a string via the RegEx object's Execute() method.
        """

        if (str(tmp_lhs).lower() != "vbscript.regexp"):
            return None

        if (".Execute(" not in str(self)):
            return None


        pat_var = str(self.lhs).lower() + ".pattern"
        pat = None
        try:
            pat = context.get(pat_var)
        except KeyError:

            return None

        tmp_rhs = self.rhs
        if (isinstance(tmp_rhs, list) and (len(tmp_rhs) > 0)):
            tmp_rhs = tmp_rhs[0]
        if ((not isinstance(tmp_rhs, Function_Call)) or
            (tmp_rhs.name != "Execute")):
            return None
        mod_str = tmp_rhs.params[0]
        try:
            str_val = context.get(mod_str)
            mod_str = str_val
        except KeyError:

            return None

        r = re.findall(pat, mod_str)
        return r
        
    def _read_member_expression_as_var(self, context, tmp_lhs):

        if (isinstance(tmp_lhs, dict)):

            key = str(self.rhs).replace("[", "").replace("]", "").replace("'", "")
            if (key.lower() in tmp_lhs.keys()):

                return tmp_lhs[key.lower()]

            if (key.lower() == "text"):
                key = "value"
                if (key.lower() in tmp_lhs.keys()):

                    return tmp_lhs[key.lower()]
        
        try:
            r = context.get(str(self), search_wildcard=False)
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("Member access " + str(self) + " stored as variable = " + str(r))
            return r
        except KeyError:
            pass

        try:
            r = context.get(str(self), search_wildcard=True)
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("Member access " + str(self) + " resolved via wildcard = " + str(r))
            log.info("Resolved %s via wildcard UserForm property lookup (value starts with %r, %d chars).",
                      str(self), str(r)[:30], len(str(r)))
            return r
        except KeyError:
            pass

        expr_list = [self.lhs]
        if (isinstance(self.rhs, list)):
            expr_list += self.rhs
        if (isinstance(self.rhs1, list)):
            expr_list += self.rhs1
        memb_str = ""
        first = True
        for expr in expr_list:

            if (not first):
                memb_str += "."
            first = False
            if (not isinstance(expr, Function_Call)):
                memb_str += str(expr)
                continue

            evaled_params = eval_args(expr.params, context)

            func_str = str(expr.name) + "("
            func_first = True
            for param in evaled_params:
                if (not func_first):
                    func_str += ", "
                func_first = False
                if (isinstance(param, float)):
                    param = int(param)
                if (isinstance(param, int)):
                    param = "'" + str(param) + "'"
                func_str += coerce_to_str(param)
            func_str += ")"
            memb_str += func_str
                        
        try:
            r = context.get(memb_str, search_wildcard=False)
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("Member access " + str(self) + " stored as variable = " + str(r))
            return r
        except KeyError:
            pass

        try:
            r = context.get(memb_str, search_wildcard=True)
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("Member access " + str(self) + " resolved via wildcard (2) = " + str(r))
            log.info("Resolved %s via wildcard UserForm property lookup (value starts with %r, %d chars).",
                      str(self), str(r)[:30], len(str(r)))
            return r
        except KeyError:
            pass

        if ("Pages(" in memb_str):
            tmp_str = memb_str[memb_str.index("Pages("):]
            try:
                r = context.get(tmp_str, search_wildcard=False)
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("Member access " + str(self) + " stored as variable = " + str(r))
                return r
            except KeyError:
                pass
            
        return None

    def _handle_usedrange_call(self, context):
        """
        Handle things like ActiveSheet.UsedRange or Sheets(a).UsedRange.
        """

        rhs = None
        if (len(self.rhs1) > 0):
            rhs = self.rhs1
        else:
            rhs = self.rhs[len(self.rhs) - 1]
        if (str(rhs) != "UsedRange"):
            return None

        sheet = None
        if (isinstance(self.lhs, Function_Call) and
            (self.lhs.name == "Sheets")):
            sheet = eval_arg(self.lhs, context)

        new_usedrange = None
        try:
            new_usedrange = function_call.parseString("UsedRange()", parseAll=True)[0]
            new_usedrange.params = []
        except ParseException as e:
            log.error("Parsing synthetic UsedRange() failed. " + str(e))
            return None
        if (sheet is not None):
            new_usedrange.params.append(sheet)

        return eval_arg(new_usedrange, context)
    
    def _eval_cell_range(self, context, just_expr=False):
        """
        Evaluate a member access expression that results in a range of Excel cells.
        """

        range_exp_str = str(self.lhs).replace("'", "")
        for exp in self.rhs[:-1]:
            range_exp_str += "." + str(exp).replace("'", "")
        cells = None
        try:

            obj = expression.parseString(range_exp_str, parseAll=True)[0]
            if just_expr:
                return obj

            cells = eval_arg(obj, context)
            return cells
        except ParseException:
            log.warning("Parse error. Cannot parse cell range expression '" + range_exp_str + "'")            
            return None
        except Exception as e:
            log.error("Cannot eval cell range expression '" + range_exp_str + "'. " + str(e))
            return None
        
    def _handle_specialcells_call(self, context):
        """
        Handle things like ActiveSheet.UsedRange.SpecialCells(xlCellTypeConstants)
        """

        rhs = None
        if (len(self.rhs1) > 0):
            rhs = self.rhs1
        else:
            rhs = self.rhs[len(self.rhs) - 1]
        if ((not isinstance(rhs, Function_Call)) or
            (rhs.name != "SpecialCells")):
            return None

        cells = self._eval_cell_range(context)
        if (cells is None):
            return None

        new_special_cells = Function_Call(None, None, None, old_call=rhs)
        tmp = [cells, new_special_cells.params[0]]
        new_special_cells.params = tmp
        r = new_special_cells.eval(context)
        
        return r

    def _eval_nested_methods(self, context):
        """
        Given a member access expression like foo(1).bar(2).baz(3)
        evaluate this as nested function calls (conceptually) like baz(3, bar(2, foo(1))).
        """

        res_func = self._convert_nested_methods_to_func_call(context)
        if (res_func is None):
            return None

        r = eval_arg(res_func, context)
        return r

    def _handle_stringbuilder_method(self, context, lhs_val):
        """
        Handle string builder object appends like 'foo.Append_3 "aaa"' and
        string builder string conversions like 'foo.ToString'.
        """

        if (not str(lhs_val).lower().endswith("stringbuilder")):
            return None


        rhs = self.rhs
        if (isinstance(rhs, list)):
            rhs = rhs[0]
        if (isinstance(rhs, Function_Call) and (str(rhs.name) == "Append_3")):

            synth_var = str(self.lhs) + ".__BUFFER__"

            buffer_val = ""
            if (context.contains(synth_var)):
                buffer_val = context.get(synth_var)

            if (len(rhs.params) == 0):
                return None
            str_val = eval_arg(rhs.params[0], context)

            buffer_val += str_val
            context.set(synth_var, buffer_val, force_global=True)

            return buffer_val

        if (str(rhs) == "ToString"):

            synth_var = str(self.lhs) + ".__BUFFER__"

            buffer_val = ""
            if (context.contains(synth_var)):
                buffer_val = context.get(synth_var)

            return buffer_val
                
        return None

    def _handle_parentdirectory(self, context):
        """Handle reading the ParentFolder property for things like
        undertakesPurposes.GetSpecialFolder(2).ParentFolder.

        """

        self_str = str(self).strip()
        if (not self_str.endswith(".ParentFolder")):
            return None

        child_folder = ""
        if (isinstance(self.rhs, list) and (len(self.rhs) > 1)):
            child_folder = str(eval_arg(self.rhs[-2], context))
        return child_folder + "\.."

    def _handle_exec(self, context):
        """Handle calling the WSCriptShell Exec() method.

        """

        if (isinstance(self.rhs, list) and
            (len(self.rhs) > 0) and
            (str(self.rhs[0]).startswith("Exec("))):
            eval_arg(self.rhs[0], context)            
    
    def eval(self, context, params=None):

        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("MemberAccess eval of " + str(self))

        tmp_lhs = None
        if (self.lhs is not None):
            if isinstance(self.lhs, SimpleNameExpression):
                try:
                    tmp_lhs = context.get(str(self.lhs))
                except KeyError:
                    tmp_lhs = "NULL"
            else:
                tmp_lhs = eval_arg(self.lhs, context)
        else:
            tmp_lhs = eval_arg(context.with_prefix, context)

        self._handle_exec(context)
            
        r = self._handle_usedrange_call(context)
        if (r is not None):
            return r
            
        r = self._handle_0_arg_call(context)
        if (r is not None):
            return r

        r = self._handle_stringbuilder_method(context, tmp_lhs)
        if (r is not None):
            return r

        r = self._handle_parentdirectory(context)
        if (r is not None):
            return r
        
        r = self._read_member_expression_as_var(context, tmp_lhs)
        if (r is not None):
            return r

        rhs = None
        if (len(self.rhs1) > 0):
            rhs = self.rhs1
        else:
            rhs = self.rhs[len(self.rhs) - 1]
            if ((str(rhs) == "Text") and (len(self.rhs) > 1)):
                rhs = self.rhs[len(self.rhs) - 2]

        calling_func = isinstance(rhs, Function_Call)
        if (not calling_func):
            try:
                func = context.get(str(rhs), search_wildcard=False)
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("Member access " + str(self) + " got RHS = " + str(func))
                calling_func = (isinstance(func, procedures.Function) or isinstance(func, procedures.Sub))
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("Member access " + str(self) + " calling function = " + str(calling_func))
            except KeyError:
                pass

        call_retval = self._handle_specialcells_call(context)
        if (call_retval is not None):
            return call_retval
        
        call_retval = self._handle_indexed_pages_access(context)
        if (call_retval is not None):
            return call_retval
            
        call_retval = self._handle_indexed_form_access(context)
        if (call_retval is not None):
            return call_retval
        
        call_retval = self._handle_control_read(context)
        if (call_retval is not None):
            return call_retval        
        
        call_retval = self._handle_oslanguage(context)
        if (call_retval is not None):
            return call_retval

        call_retval = self._handle_paragraphs(context)
        if (call_retval is not None):
            return call_retval

        call_retval = self._handle_comments(context)
        if (call_retval is not None):
            return call_retval
        
        call_retval = self._handle_application_run(context)
        if (call_retval is not None):
            return call_retval

        call_retval = self._handle_docprops_read(context)
        if (call_retval is not None):
            return call_retval
        
        if (not calling_func):
            call_retval = self._handle_docvars_read(context)
            if (call_retval is not None):
                return call_retval

        call_retval = self._handle_set_clipboard(context)
        if (call_retval is not None):
            return call_retval

        call_retval = self._handle_get_clipboard(context)
        if (call_retval is not None):
            return call_retval
                    
        call_retval = self._handle_table_cell(context)
        if (call_retval is not None):
            return call_retval
                        
        call_retval = self._handle_count(context, tmp_lhs)
        if (call_retval is not None):
            return call_retval

        call_retval = self._handle_item(context, tmp_lhs)
        if (call_retval is not None):
            return call_retval

        call_retval = self._handle_regex_execute(context, tmp_lhs)
        if (call_retval is not None):
            return call_retval

        call_retval = self._handle_0_arg_call(context, rhs)
        if (call_retval is not None):
            return call_retval
        
        call_retval = self._handle_text_file_read(context)
        if (call_retval is not None):
            return call_retval

        if (self._handle_adodb_writes(self.lhs, tmp_lhs, rhs, context)):
            return "NULL"

        call_retval = self._handle_path_access()
        if (call_retval is not None):
            return call_retval

        call_retval = self._handle_listbox_list(context, tmp_lhs, self.rhs)
        if (call_retval is not None):
            return call_retval

        call_retval = self._handle_replace(context, tmp_lhs, self.rhs)
        if (call_retval is not None):
            return call_retval
        
        call_retval = self._eval_nested_methods(context)
        if (call_retval is not None):
            return call_retval
        
        if (calling_func):
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug('rhs {!r} is a Function_Call'.format(rhs))

            rhs_name = str(rhs)
            if (hasattr(rhs, "name")):
                rhs_name = rhs.name
            if (context.contains_user_defined(rhs_name)):
                for func in Function_Call.log_funcs:
                    if (rhs_name.lower() == func.lower()):
                        return str(self)

            call_retval = self._handle_add(context, tmp_lhs, self.rhs)
            if (call_retval is not None):
                return call_retval

            call_retval = self._handle_listbox_additem(context, tmp_lhs, self.rhs)
            if (call_retval is not None):
                return call_retval

            call_retval = self._handle_exists(context, tmp_lhs, self.rhs)
            if (call_retval is not None):
                return call_retval

            call_retval = self._handle_excel_read(context, self.rhs)
            if (call_retval is not None):
                return call_retval
                    
            tmp_rhs = eval_arg(rhs, context)

            if (self._handle_loadxml(context, tmp_rhs)):
                return "NULL"

            if (self._handle_savetofile(context, tmp_rhs)):
                return "NULL"

            return tmp_rhs

        elif (isinstance(rhs, Function_Call_Array_Access)):

            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug('rhs {!r} is a Function_Call_Array_Access'.format(rhs))
            tmp_rhs = eval_arg(rhs, context)
            return tmp_rhs
            
        elif (str(self.lhs) != str(tmp_lhs)):

            if (((isinstance(tmp_lhs, str)) or (isinstance(tmp_lhs, SimpleNameExpression))) and
                (str(tmp_lhs) != "NULL") and
                (not "Shapes(" in str(tmp_lhs)) and
                (not "Close" in str(self.rhs)) and
                (not context.contains(str(self.lhs)))):

                return str(tmp_lhs)

            call_retval = self._handle_docvar_value(tmp_lhs, self.rhs)
            if (call_retval is not None):
                return call_retval

            call_retval = self._handle_file_close(context, tmp_lhs, self.rhs)
            if (call_retval is not None):
                return call_retval

            if ((isinstance(tmp_lhs, procedures.Function)) and
                (len(tmp_lhs.params) == 0)):

                r = tmp_lhs.eval(context)
                return r

            if (((str(self.rhs) == "['Text']") or (str(self.rhs).lower() == "['value']")) and (isinstance(tmp_lhs, str))):
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("Returning .Text value.")
                return tmp_lhs
            
            r = MemberAccessExpression(None, None, None, raw_fields=(tmp_lhs, self.rhs, self.rhs1))
            
            call_retval = r._handle_docvars_read(context)
            if (call_retval is not None):
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("MemberAccess: Found " + str(r) + " = '" + str(call_retval) + "'") 
                return call_retval

            tmp_rhs = eval_arg(rhs, context)
            var_pat = r"[A-za-z_0-9]+"
            if ((tmp_rhs != rhs) and
                ((re.match(var_pat, str(tmp_lhs)) is not None) or
                 (str(tmp_lhs).lower().endswith(".application"))) and
                (tmp_rhs != "NULL") and
                ("simulation_vba.core.vba_library" not in str(type(tmp_rhs)))):
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("Resolved member access variable.")
                return tmp_rhs        
            
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("MemberAccess: Return new access object " + str(r))
            return r

        elif (context.contains(rhs)):
            return context.get(rhs)
        
        else:
            return eval_arg(self.__repr__(), context)
        
l_expression = Forward()

function_call_limited = Forward()
func_call_array_access_limited = Forward()
function_call = Forward()
excel_expression = Forward()

member_object_limited = (
    ((Suppress("[") + (unrestricted_name_expression | decimal_literal) + Suppress("]")) | unrestricted_name_expression | excel_expression)
    + NotAny("(")
    + NotAny("#")
    + NotAny("$")
    + NotAny("!")
)
member_object_loose = Suppress(Literal("(")) + ((func_call_array_access_limited ^ function_call_limited) | member_object_limited) + Suppress(Literal(")")) | \
                      ((func_call_array_access_limited ^ function_call_limited) | member_object_limited)
member_object_strict = Suppress(Optional(".")) + NotAny(reserved_identifier) + member_object_loose

member_access_expression = Group(Group(member_object_strict("lhs") + OneOrMore((Suppress(".") | Suppress("!")) + member_object_loose("rhs"))))
member_access_expression.setParseAction(MemberAccessExpression)


member_access_expression_loose = Group(
    Group(
        Suppress(ZeroOrMore(" "))
        + member_object_strict("lhs")
        + OneOrMore(Suppress(".") + member_object_loose("rhs"))
    )
    + Suppress(ZeroOrMore(" "))
)
member_access_expression_loose.setParseAction(MemberAccessExpression)


member_access_expression_limited = Group(
    Group((
        member_object_strict("lhs")
        + NotAny(White())
        + Suppress(".")
        + NotAny(White())
        + member_object_limited("rhs")
        + Optional(
              NotAny(White())
              + Suppress(".")
              + NotAny(White())
              + member_object_limited("rhs1")
        )
    ).leaveWhitespace())
)
member_access_expression_limited.setParseAction(MemberAccessExpression)




procedure_pointer_expression = member_access_expression | simple_name_expression
addressof_expression = CaselessKeyword("addressof").suppress() + procedure_pointer_expression


argument_expression = (Optional(CaselessKeyword("byval")) + expression) | addressof_expression

class NamedArgument(VBA_Object):

    def __init__(self, original_str, location, tokens, name=None):
        super(NamedArgument, self).__init__(original_str, location, tokens)

        self.name = tokens.name
        self.value = tokens.value
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed "%r" as NamedArgument' % self)

    def __repr__(self):
        return '%s:=%s' % (self.name, self.value)

    def eval(self, context, params=None):
        try:
            return eval_arg(self.value, context)
        except:
            log.error("NamedArgument: Cannot eval " + self.__repr__() + ".")
            return ''
    
named_argument = unrestricted_name('name') + Suppress(":=") + argument_expression('value')
named_argument.setParseAction(NamedArgument)
named_argument_list = delimitedList(named_argument)
required_positional_argument = argument_expression
positional_argument = Optional(argument_expression)
positional_or_named_argument_list = Optional(delimitedList(positional_argument) + ",") \
                                    + (named_argument_list | required_positional_argument)
argument_list = Optional(positional_or_named_argument_list)



index_expression = simple_name_expression + Suppress("(") + simple_name_expression + Suppress(")")




dictionary_access_expression = l_expression + Suppress("!") + unrestricted_name



class With_Member_Expression(VBA_Object):
    
    def __init__(self, original_str, location, tokens, old_call=None):
        super(With_Member_Expression, self).__init__(original_str, location, tokens)
        old_call = old_call # pylint warning
        self.expr = tokens.expr
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed %r as With_Member_Expression' % self)

    def __repr__(self):
        return "." + str(self.expr)

    def to_python(self, context, params=None, indent=0):
        indent = indent # pylint warning
        
        with_dict = None
        if ((context.with_prefix_raw is not None) and
            (context.contains(str(context.with_prefix_raw)))):
            with_dict = context.get(str(context.with_prefix_raw))
            if (with_dict == "__ALREADY_SET__"):

                with_dict = context.get("__ORIG__" + str(context.with_prefix_raw))

            if (not isinstance(with_dict, dict)):                
                with_dict = None

        if (with_dict is None):
            return "ERROR: Only doing JIT on Scripting.Dictionary With blocks."

        expr_str = str(self)
        if ((not expr_str.startswith(".Exists")) and
            (not expr_str.startswith(".Items")) and
            (not expr_str.startswith(".Item")) and
            (not expr_str.startswith(".Count"))):
            return "ERROR: Only doing JIT on Scripting.Dictionary methods in With blocks."

        if (expr_str == ".Count"):
            return "(len(" + context.with_prefix_raw + ") - 1)"

        tmp_var = SimpleNameExpression(None, None, None, name=str(context.with_prefix_raw))
        new_exists = Function_Call(None, None, None, old_call=self.expr)
        tmp = [tmp_var]
        for p in new_exists.params:
            tmp.append(p)
        new_exists.params = tmp
        r = to_python(new_exists, context, params)
        return r
        
    def _handle_method_calls(self, context):
        """
        Handle Scripting.Dictionary...() calls.
        """

        expr_str = str(self)
        if ((not expr_str.startswith(".Exists")) and (not expr_str.startswith(".Count"))):
            return None

        if ((context.with_prefix_raw is None) or
            (not context.contains(str(context.with_prefix_raw)))):
            return None
        with_dict = context.get(str(context.with_prefix_raw))

        if (expr_str == ".Count"):
            return (len(with_dict) - 1)

        if ((not hasattr(self.expr, "name")) or
            (not hasattr(self.expr, "params"))):
            return None
        
        new_exists = Function_Call(None, None, None, old_call=self.expr)
        tmp = [with_dict]
        for p in new_exists.params:
            tmp.append(p)
        new_exists.params = tmp
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Dictionary NNNNNN() func = " + str(new_exists))
        
        r = new_exists.eval(context)
        return r
        
    def eval(self, context, params=None):

        call_retval = self._handle_method_calls(context)
        if (call_retval is not None):
            return call_retval

        return self.expr.eval(context, params)


with_member_access_expression = Suppress(".") + \
                                (simple_name_expression("expr") ^ function_call_limited("expr") ^ member_access_expression("expr")) 
with_member_access_expression.setParseAction(With_Member_Expression)
with_dictionary_access_expression = Suppress("!") + unrestricted_name
with_expression = with_member_access_expression | with_dictionary_access_expression



boolean_expression = Forward()
new_expression = Forward()
# pylint: disable=pointless-statement
l_expression << (with_expression ^ member_access_expression ^ new_expression ^ member_access_expression_loose) | \
    instance_expression | \
    dictionary_access_expression | \
    simple_name_expression


class Function_Call(VBA_Object):
    """
    Function call within a VBA expression
    """

    log_funcs = ["CreateProcessA", "CreateProcessW", "CreateProcess", ".run", "CreateObject",
                 "Open", ".Open", "GetObject", "Create", ".Create", "Environ",
                 "CreateTextFile", ".CreateTextFile", ".Eval", "Run",
                 "SetExpandedStringValue", "WinExec", "FileExists", "SaveAs",
                 "FileCopy", "Load", "ShellExecute", "FolderExists"]
    
    def __init__(self, original_str, location, tokens, old_call=None):
        super(Function_Call, self).__init__(original_str, location, tokens)

        if (old_call is not None):
            self.name = old_call.name
            if (hasattr(old_call.params, "copy")):
                self.params = old_call.params.copy()
            else:
                self.params = old_call.params
            return

        self.name = str(tokens.name)
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('Function_Call.name = %r' % self.name)
        assert isinstance(self.name, basestring)
        self.params = tokens.params

        array_pat = r"\w+\(.+\)"
        if ((self.name.lower() == "multibytetowidechar") and
            (len(self.params) == 6) and
            (re.match(array_pat, str(self.params[2])) is not None) and
            (isinstance(self.params[2], Function_Call))):

            array = None
            orig_array = self.params[2]
            try:
                array = expression.parseString(self.params[2].name, parseAll=True)[0]
            except ParseException:
                pass
            if (array is not None):
                self.params[2] = array
                log.warning("Rewrote MultiByteToWideChar() array reference '" + str(orig_array) + "' to '" + str(array) + "'.")

        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('Function_Call.params = %r' % self.params)
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed %r as Function_Call' % self)

    def __repr__(self):
        parms = ""
        first = True
        for parm in self.params:
            if (not first):
                parms += ", "
            first = False
            parms += str(parm)
        return '%s(%r)' % (self.name, parms)

    def eval(self, context, params=None):

        import vba_library
        vba_library.var_names = self.params

        log.info(
            "EVAL FUNCTION_CALL: name=%s params_count=%d",
            self.name, len(self.params) if hasattr(self.params, '__len__') else '?',
        )
        if hasattr(self.params, '__iter__'):
            for _pi, _p in enumerate(self.params):
                _p_type = type(_p).__name__
                _p_repr = repr(_p)[:200]
                log.info("  Eval raw param[%d]: type=%s repr=%s", _pi, _p_type, _p_repr)
        
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Function_Call: eval params: " + str(self.params))

        dll_func_name = context.get_true_name(self.name)
        is_external = False
        if (dll_func_name is not None):
            is_external = True
            self.name = dll_func_name

        params = None
        if (self.name == "CallByName"):
            params = eval_args(self.params[1:], context=context)
            params = [self.params[0]] + params
        else:
            if (self.name.lower() in ("write", "hdynjmt", "print")):
                _raw_types = [type(p).__name__ for p in self.params]
                _raw_strs = [str(p)[:100] for p in self.params]
                _raw_vba = [isinstance(p, VBA_Object) for p in self.params]
                log.info(
                    "Function_Call %s: raw params types=%s vba=%s preview=%r",
                    self.name, _raw_types, _raw_vba, _raw_strs,
                )
                for _pi, _p in enumerate(self.params):
                    _p_type = type(_p).__name__
                    _p_bases = [b.__name__ for b in type(_p).__mro__]
                    _p_repr = repr(_p)[:200]
                    _p_str = str(_p)[:100]
                    log.info(
                        "  raw param[%d]: type=%s bases=%s repr=%s str=%s",
                        _pi, _p_type, _p_bases, _p_repr, _p_str,
                    )
                    if isinstance(_p, VBA_Object) and hasattr(_p, 'arg'):
                        log.info("    arg content: %s", [str(a)[:40] for a in _p.arg])
                        for _ai, _a in enumerate(_p.arg):
                            log.info(
                                "      operand[%d]: type=%s repr=%s",
                                _ai, type(_a).__name__, repr(_a)[:80],
                            )
            params = eval_args(self.params, context=context)
            if (self.name.lower() in ("write", "hdynjmt", "print")):
                _eval_type = type(params).__name__
                _eval_iterable = hasattr(params, '__iter__')
                _eval_len = len(params) if _eval_iterable else 'N/A'
                log.info(
                    "  eval_args returned: type=%s iterable=%s len=%s",
                    _eval_type, _eval_iterable, _eval_len,
                )
                if _eval_iterable:
                    for _ei, _ep in enumerate(params):
                        log.info(
                            "    param[%d]: type=%s repr=%s",
                            _ei, type(_ep).__name__, repr(_ep)[:120],
                        )
        str_params = repr(params)[1:-1]
        if (len(str_params) > 80):
            str_params = str_params[:80] + "..."
            
        if (context.have_error()):
            log.warn('Short circuiting function call %s(%s) due to thrown VB error.' % (self.name, str_params))
            return None

        if (str(self.name).strip().lower() in ("application.goto", "goto") and
            (len(params) > 0) and hasattr(context, "set_excel_selection")):
            context.set_excel_selection(params[0])
            return "NULL"

        skip_report_functions = set(["cos", "tan"])
        if (str(self.name).lower() not in skip_report_functions):
            if (not context.throttle_logging):
                log.info('calling Function: %s(%s)' % (self.name, str_params))
        
        if (is_external):

            context.report_action("External Call", self.name + "(" + str(params) + ")", self.name, strip_null_bytes=True)

            try:
                s = context.get_lib_func(self.name)
                if (s is None):
                    raise KeyError("func not found")
                r = s.eval(context=context, params=params)
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("External function " + str(s.name) + " returns " + str(r))
                return r
            except KeyError:
                log.warning("External function " + str(self.name) + " not found.")
                return "NULL"

        # pylint: disable=protected-access
        if self.name.lower() in context._log_funcs \
                or any(self.name.lower().endswith(func.lower()) for func in Function_Call.log_funcs):
            if ("Scripting.Dictionary" not in str(params)):
                context.report_action(self.name, params, 'Interesting Function Call', strip_null_bytes=True)
        try:

            f = context.get(self.name)
            
            if (isinstance(f, dict)):

                if (len(params) > 0):
                    if (log.getEffectiveLevel() == logging.DEBUG):
                        log.debug('Dict Access: %r[%r]' % (f, params[0]))
                    index = params[0]
                    if (index in f):
                        return f[index]
                    return "NULL"
            
            if (isinstance(f, list)):

                if (len(params) > 0):
                    tmp = f
                    if (log.getEffectiveLevel() == logging.DEBUG):
                        log.debug('Array Access: %r[%r]' % (tmp, str(params)))
                    index = utils.int_convert(params[0])
                    index1 = None
                    if (len(params) > 1):
                        index1 = utils.int_convert(params[1])
                    try:

                        if (index1 is None):
                            r = tmp[index]
                        else:
                            r = tmp[index][index1]
                        if (log.getEffectiveLevel() == logging.DEBUG):
                            log.debug('Returning: %r' % r)
                        return r
                    except Exception as e:

                        msg = 'Array Access Failed: %r[%r] %r' % (tmp, str(params), str(e))
                        context.set_error(msg)
                        return 0

                else:

                    return f
                    
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug('Calling: %r' % f)

            if ((isinstance(f, str)) and (context.contains(f))):
                tmp_f = context.get(f)
                if (isinstance(tmp_f, VbaLibraryFunc)):
                    f = tmp_f

            if (f is not None):
                if (isinstance(f, (procedures.Function, procedures.Sub)) or
                    ("vba_library." in str(type(f)))):
                    try:

                        r = f.eval(context=context, params=params)                        
                        
                        if (hasattr(f, "byref_params")):
                            for byref_param_info in f.byref_params.keys():
                                try:
                                    arg_var_name = str(self.params[byref_param_info[1]])
                                    if (context.contains(arg_var_name)):

                                        if (not isinstance(f, (VbaLibraryFunc, procedures.Function, procedures.Sub))):
                                            context.set(arg_var_name, f.byref_params[byref_param_info])
                                except IndexError:
                                    break

                        context.exit_func = False
                                    
                        return r

                    except AttributeError as e:

                        log.error(str(f) + " has no eval() method. " + str(e))
                        return f

                elif ((isinstance(f, int)) and
                      (len(params) == 1) and
                      (isinstance(params[0], int))):
                    return (f + params[0])

                elif (len(params) > 0):

                    if (log.getEffectiveLevel() == logging.DEBUG):
                        log.debug("Looks like array access.")
                    try:

                        i = utils.int_convert(params[0])
                        r = f[i]
                        if (isinstance(f, str)):
                            r = ord(r)
                        if (log.getEffectiveLevel() == logging.DEBUG):
                            log.debug("Return " + str(r))
                        return r

                    except Exception as e:

                        log.error("Array access %r[%r] failed. %r" % (f, params[0], str(e)))
                        return 0
            else:

                log.error('Function %r resolves to None' % self.name)
                return None

        except KeyError:

            func_name = str(self.name)
            if ((func_name == "Application.Run") or (func_name == "Run")):

                new_func = params[0]

                new_params = params[1:]

                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("Try indirect run of function '" + new_func + "'")
                r = "NULL"
                try:

                    s = new_func
                    while (isinstance(s, str)):

                        s = context.get(s)
                        if (isinstance(s, (VbaLibraryFunc, procedures.Function, procedures.Sub))):
                            s = s.eval(context=context, params=new_params)
                            r = s

                    if ((str(self).lower().startswith("thisdocument.run(")) and (r != "NULL")):
                        context.report_action('Execute Command', r, 'ThisDocument.Run', strip_null_bytes=True)

                except KeyError:
                    pass

                if (r != "NULL"):
                    return r

            if (context.contains(self.name) and
                (len(params) == 1) and
                (isinstance(params[0], int))):

                var_val = context.get(self.name)

                if (isinstance(var_val, int)):

                    return (var_val + params[0])
                
            context.increase_general_errors()
            log.warning('Function %r not found' % self.name)
            return None

        return None
        
    def to_python(self, context, params=None, indent=0):
        indent = indent # pylint warning

        log.info(
            "FUNCTION_CALL.to_python: name=%s raw_params_count=%d",
            self.name, len(self.params) if hasattr(self.params, '__len__') else '?',
        )
        if hasattr(self.params, '__iter__'):
            for _pi, _p in enumerate(self.params):
                log.info(
                    "  raw param[%d]: type=%s repr=%s",
                    _pi, type(_p).__name__, repr(_p)[:200],
                )

        dll_func_name = context.get_true_name(self.name)
        func_name = self.name
        is_external = False
        if (dll_func_name is not None):
            is_external = True
            func_name = dll_func_name
        
        py_params = []
        old_bitwise = context.in_bitwise_expression
        context.in_bitwise_expression = True
        for p in self.params:
            log.info(
                "  -> to_python param p: type=%s repr=%s",
                type(p).__name__, repr(p)[:200],
            )
            _result = to_python(p, context, params)
            log.info(
                "  <- to_python result: type=%s repr=%s",
                type(_result).__name__, _result[:200] if isinstance(_result, str) else repr(_result)[:200],
            )
            py_params.append(_result)
        context.in_bitwise_expression = old_bitwise

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
            args += "]"

            r = None
            if is_internal:
                r = "core.vba_library.run_function(\"" + str(func_name) + "\", vm_context, " + args + ")"
            else:
                r = "core.vba_library.run_external_function(\"" + str(func_name) + "\", vm_context, " + args + ",\"\")"
            return r

        if (context.contains(func_name)):
            ref = context.get(func_name)
            ref1 = None
            try:
                ref1 = context.get("__ORIG__" + func_name)
            except KeyError:
                pass
            if ((isinstance(ref, list)) or
                (isinstance(ref1, list)) or
                (ref == "__FUNC_ARG__")):

                acc_str = ""
                for p in py_params:
                    acc_str += "[coerce_to_int(" + p + ")]"
                r = str(func_name) + acc_str
                return r
        
        r = str(func_name) + "("
        first = True
        for p in py_params:
            if (not first):
                r += ", "
            first = False
            r += p
        r += ")"

        return r
        

expr_item = Forward()
expr_item_strict = Forward()
expr_list_item = Optional(Suppress(CaselessKeyword("ByVal") | CaselessKeyword("ByRef"))) + \
                 expression ^ boolean_expression ^ member_access_expression_loose
expr_list_item_strict = Optional(Suppress(CaselessKeyword("ByVal") | CaselessKeyword("ByRef"))) + \
                        NotAny(CaselessKeyword("End")) + \
                        (expression ^ boolean_expression ^ member_access_expression_loose)
expr_list_item = (expr_item + FollowedBy(',')) | expr_list_item
expr_list_item_strict = (expr_item_strict + FollowedBy(',')) | expr_list_item_strict

def quick_parse_int_or_var(text):
    text = str(text).strip()

    if (text.isdigit()):
        return int(text)        

    if (re.match(r"[_a-zA-Z][_a-zA-Z\d]*", text) is not None):
        r = SimpleNameExpression(None, None, None, text)
        return r

    r = expression.parseString(text, parseAll=True)[0]
    return r
    

expr_list_fast = Regex("(?:\s*[0-9a-zA-Z_]+[ \t\f\v]*,[ \t\f\v]*){10,}[ \t\f\v]*[0-9a-zA-Z_]+[ \t\f\v]*")
expr_list_fast.setParseAction(lambda t: [quick_parse_int_or_var(i) for i in t[0].split(",")])

expr_list_slow = delimitedList(Optional(expr_list_item, default=""))

expr_list = (
    expr_list_item
    + NotAny(':=')
    + Optional(Suppress(",") + (expr_list_fast | expr_list_slow))
)
expr_list_strict = (
    expr_list_item_strict
    + NotAny(':=')
    + Optional(Suppress(",") + (expr_list_fast | expr_list_slow))
)

function_call <<= (
    CaselessKeyword("nothing")
    | (
        ~(strict_reserved_keywords + Literal("(")) +
        (
            (Suppress(Optional("#")) + (member_access_expression('name') ^ lex_identifier('name'))) |
            (Suppress('[') + lex_identifier('name') + Suppress(']'))
        ) +
        Suppress(
            Optional('$')
            + Optional('#')
            + Optional('!')
            + Optional('%')
            + Optional('@')
        )
        + ((Suppress('(') + Optional(expr_list('params')) + Suppress(')')) |
           (Suppress('[') + Optional(expr_list('params')) + Suppress(']')))
    )
    | (
        Suppress('[') +
        CaselessKeyword("Shell")('name') +
        Suppress(']') +
        expr_list('params')
    )
    | (
        Suppress('[') + lex_identifier('name') + Suppress('(') + expr_list('params') + Suppress(')') + Suppress(']')
    )
)
function_call.setParseAction(Function_Call)

function_call_limited <<= (
    CaselessKeyword("nothing")
    | (
        (lex_identifier('name') | (Suppress('[') + lex_identifier('name') + Suppress(']')))
        + Suppress(Optional('$'))
        + Suppress(Optional('#'))
        + Suppress(Optional('!'))
        + Suppress(Optional('%'))
        + Suppress(Optional('@'))
        + (
            (Suppress('(') + Optional(expr_list('params')) + Suppress(')'))
            | (Suppress('[') + Optional(expr_list('params')) + Suppress(']'))
            | (Suppress(Optional('$')) + NotAny(".") + NotAny("-") + NotAny(CaselessKeyword("step")) + expr_list('params'))
        )
    )
)
function_call_limited.setParseAction(Function_Call)


class Function_Call_Array_Access(VBA_Object):
    """
    Array access of the return value of a function call.
    """

    def __init__(self, original_str, location, tokens):
        super(Function_Call_Array_Access, self).__init__(original_str, location, tokens)
        self.array = tokens.array
        self.index = tokens.index
        self.other_indices = None
        if (hasattr(tokens, "other_indices")):
            self.other_indices = tokens.other_indices
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed %r as Function_Call_Array_Access' % self)

    def __repr__(self):
        r = str(self.array) + "(" + str(self.index) + ")"
        if (self.other_indices is not None):
            r = str(self.array) + "(" + str(self.index) + ", " + str(self.other_indices) + ")"
        return r

    def eval(self, context, params=None):
        params = params # pylint warning

        array_val = eval_arg(self.array, context=context)
        array_index = coerce_to_int(eval_arg(self.index, context=context))

        if (not isinstance(array_val, list)):
            log.error("%r is not a list. Cannot perform array access." % array_val)
            return ''

        if (not isinstance(array_index, int)):
            log.error("Index %r is not an integer. Cannot perform array access." % array_index)
            return ''
        if ((array_index >= len(array_val)) or (array_index < 0)):
            log.error("Index %r is outside array bounds. Cannot perform array access." % array_index)
            return ''

        return array_val[array_index]
            

func_call_array_access = function_call("array") + Suppress("(") + \
                         expression("index") + ZeroOrMore(Suppress(Literal(",")) + expression)("other_indices") + \
                         Suppress(")")
func_call_array_access.setParseAction(Function_Call_Array_Access)

func_call_array_access_limited <<= function_call_limited("array") + Suppress("(") + \
                                   expression("index") + ZeroOrMore(Suppress(Literal(",")) + expression)("other_indices") + \
                                   Suppress(")")
func_call_array_access_limited.setParseAction(Function_Call_Array_Access)



typeof_expression = Forward()
addressof_expression = Forward()
literal_list_expression = Forward()
literal_range_expression = Forward()
limited_expression = Forward()
bool_expr_item = Forward()
tuple_expression = Forward()
expr_item <<= (
    Optional(CaselessKeyword("ByVal").suppress())
    + (
        date_string
        | file_pointer
        | float_literal
        | named_argument
        | l_expression
        | (chr_ ^ function_call ^ func_call_array_access)
        | simple_name_expression
        | asc
        | strReverse
        | literal
        | placeholder
        | typeof_expression
        | addressof_expression
        | excel_expression
        | literal_range_expression
        | literal_list_expression
        | Suppress(Literal("(")) + boolean_expression + Suppress(Literal(")"))
        | tuple_expression
    )
)
expr_item_strict <<= (
    Optional(CaselessKeyword("ByVal").suppress())
    + NotAny(CaselessKeyword("End"))
    + (
        date_string
        | file_pointer
        | float_literal
        | named_argument
        | l_expression
        | (chr_ ^ function_call ^ func_call_array_access)
        | simple_name_expression
        | asc
        | strReverse
        | literal
        | placeholder
        | typeof_expression
        | addressof_expression
        | excel_expression
        | literal_range_expression
        | literal_list_expression
    )
)




expression <<= infixNotation(expr_item,
                             [(CaselessKeyword("not"), 1, opAssoc.RIGHT, Not),
                              ("-", 1, opAssoc.RIGHT, Neg),
                              ("^", 2, opAssoc.RIGHT, Power),
                              (Regex(re.compile("[*/]")), 2, opAssoc.LEFT, MultiDiv),
                              ("\\", 2, opAssoc.LEFT, FloorDivision),
                              (Regex(re.compile("mod", re.IGNORECASE)), 2, opAssoc.LEFT, Mod),
                              (Regex(re.compile('[-+]')), 2, opAssoc.LEFT, AddSub),
                              ("&", 2, opAssoc.LEFT, Concatenation),
                              (";", 2, opAssoc.LEFT, Concatenation),
                              (Regex(re.compile("and", re.IGNORECASE)), 2, opAssoc.LEFT, And),
                              (Regex(re.compile("or", re.IGNORECASE)), 2, opAssoc.LEFT, Or),
                              (Regex(re.compile("xor", re.IGNORECASE)), 2, opAssoc.LEFT, Xor),
                              (Regex(re.compile("eqv", re.IGNORECASE)), 2, opAssoc.LEFT, Eqv),])
expression.setParseAction(lambda t: t[0])

limited_expression <<= (infixNotation(expr_item,
                                      [("-", 1, opAssoc.RIGHT, Neg),
                                       ("^", 2, opAssoc.RIGHT, Power),
                                       (Regex(re.compile("[*/]")), 2, opAssoc.LEFT, MultiDiv),
                                       ("\\", 2, opAssoc.LEFT, FloorDivision),
                                       (CaselessKeyword("mod"), 2, opAssoc.RIGHT, Mod),
                                       (Regex(re.compile('[-+]')), 2, opAssoc.LEFT, AddSub),
                                       ("&", 2, opAssoc.LEFT, Concatenation),
                                       (CaselessKeyword("xor"), 2, opAssoc.LEFT, Xor),])) | \
                                       Suppress(Literal("(")) + expression + Suppress(")")
expression.setParseAction(lambda t: t[0])

expr_const = Forward()
chr_const = Suppress(
    Combine(CaselessLiteral('Chr') + Optional(Word('BbWw', max=1)) + Optional('$')) + '(') + expr_const + Suppress(')')
chr_const.setParseAction(Chr)
asc_const = Suppress(CaselessKeyword('Asc') + '(') + expr_const + Suppress(')')
asc_const.setParseAction(Asc)
strReverse_const = Suppress(CaselessLiteral('StrReverse') + Literal('(')) + expr_const + Suppress(Literal(')'))
strReverse_const.setParseAction(StrReverse)
environ_const = Suppress(CaselessKeyword('Environ') + '(') + expr_const + Suppress(')')
environ_const.setParseAction(Environ)
expr_const_item = (chr_const | asc_const | strReverse_const | environ_const | literal)
expr_const <<= infixNotation(expr_const_item,
                             [
                                 ("+", 2, opAssoc.LEFT, Sum),
                                 ("&", 2, opAssoc.LEFT, Concatenation),
                             ])


class BoolExprItem(VBA_Object):
    """
    A comparison expression or other item appearing in a boolean expression.
    """

    def __init__(self, original_str, location, tokens):
        super(BoolExprItem, self).__init__(original_str, location, tokens)
        assert (len(tokens) > 0)
        self.lhs = tokens[0]
        self.op = None
        self.rhs = None
        if (len(tokens) == 3):
            self.op = tokens[1]
            self.rhs = tokens[2]        
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed %r as BoolExprItem' % self)

    def __repr__(self):
        if (self.op is not None):
            return self.lhs.__repr__() + " " + self.op + " " + self.rhs.__repr__()
        elif (self.lhs is not None):
            return self.lhs.__repr__()
        log.error("BoolExprItem: Improperly parsed.")
        return ""

    def _vba_to_python_op(self, op, context):
        return _vba_to_python_op(op, not context.in_bitwise_expression)
        
    def to_python(self, context, params=None, indent=0):
        r = " " * indent
        expr_str = None
        got_op = True
        if (self.op is not None):
            expr_str = to_python(self.lhs, context, params) + " " + \
                       self._vba_to_python_op(self.op, context) + " " + \
                       to_python(self.rhs, context, params)
        elif (self.lhs is not None):
            got_op = False
            expr_str = to_python(self.lhs, context, params)
        else:
            log.error("BoolExprItem: Improperly parsed.")
            return ""

        if (context.in_bitwise_expression and got_op):
            r += "(-1 if " + expr_str + " else 0)"
        else:
            r += expr_str
        return r
        
    def eval(self, context, params=None):
        params = params # pylint warning
        
        lhs = self.lhs
        try:
            lhs = eval_arg(self.lhs, context)
        except AttributeError:
            pass

        if (self.op is None):

            return lhs

        rhs = self.rhs
        try:
            rhs = eval_arg(self.rhs, context)
        except AttributeError:
            pass

        if ((rhs == "NULL") or (rhs is None)):
            if (isinstance(lhs, str)):
                rhs = ''
            else:
                rhs = 0
            context.set(self.rhs, rhs)
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("Set unitinitialized " + str(self.rhs) + " = " + str(rhs))
        if ((lhs == "NULL") or (lhs is None)):
            if (isinstance(rhs, str)):
                lhs = ''
            else:
                lhs = 0
            context.set(self.lhs, lhs)
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("Set unitialized " + str(self.lhs) + " = " + str(lhs))

        if (isinstance(lhs, str) and isinstance(rhs, int)):

            try:
                lhs = int(lhs)
            # pylint: disable=bare-except
            except:
                pass

        if (isinstance(rhs, str) and isinstance(lhs, int)):

            try:
                rhs = int(rhs)
            # pylint: disable=bare-except
            except:
                pass

        if (isinstance(lhs, float) and isinstance(rhs, int)):
            rhs = rhs + 0.0
        if (isinstance(rhs, float) and isinstance(lhs, int)):
            lhs = lhs + 0.0

        if (isinstance(lhs, unicode)):
            lhs = ''.join(filter(lambda x:x in string.printable, lhs))
        if (isinstance(rhs, unicode)):
            rhs = ''.join(filter(lambda x:x in string.printable, rhs))
            
        rhs_invalid_type = ((not isinstance(rhs, int)) and (not isinstance(rhs, str)) and (not isinstance(rhs, float)))
        lhs_invalid_type = ((not isinstance(lhs, int)) and (not isinstance(lhs, str)) and (not isinstance(lhs, float)))
        if (rhs_invalid_type or lhs_invalid_type):

            lhs = str(lhs)
            rhs = str(rhs)

        rhs = utils.strip_nonvb_chars(rhs)
        lhs = utils.strip_nonvb_chars(lhs)
        rhs_str = str(rhs)
        lhs_str = str(lhs)
        if (("**MATCH ANY**" in lhs_str) or
            ("**MATCH ANY**" in rhs_str) or
            ("CURRENT_FILE_NAME" in lhs_str) or
            ("CURRENT_FILE_NAME" in rhs_str)):
            if (self.op == "<>"):
                if (context.in_bitwise_expression):
                    return 0
                return False
            if (context.in_bitwise_expression):
                return -1
            return True

        r = False
        if ((self.op.lower() == "=") or
            (self.op.lower() == "is")):            
            r = lhs == rhs
        elif (self.op == ">"):
            r = lhs > rhs
        elif (self.op == "<"):
            r = lhs < rhs
        elif ((self.op == ">=") or (self.op == "=>")):
            r = lhs >= rhs
        elif ((self.op == "<=") or (self.op == "=<")):
            r = lhs <= rhs
        elif (self.op == "<>"):
            r = lhs != rhs
        elif (self.op.lower() == "like"):

            rhs = str(rhs)
            lhs = str(lhs)
            try:
                rhs = rhs.replace("*", ".*")
                r = (re.match(rhs, lhs) is not None)
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("'" + lhs + "' Like '" + rhs + "' == " + str(r))
            except Exception as e:
                
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug(str(rhs) + " not valid python regex. " + str(e))
                r = (rhs == lhs)
        else:
            log.error("BoolExprItem: Unknown operator %r" % self.op)
            r = False

        if (context.in_bitwise_expression):
            if r:
                r = -1
            else:
                r = 0

        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Evaled '" + str(self) + "' == " + str(r))
                
        return r
        

bool_expr_item <<= (limited_expression + \
                    (oneOf(">= => <= =< <> = > < <>") | CaselessKeyword("Like") | CaselessKeyword("Is")) + \
                    limited_expression) | \
                    limited_expression
bool_expr_item.setParseAction(BoolExprItem)

class BoolExpr(VBA_Object):
    """
    A boolean expression.
    """

    def __init__(self, original_str, location, tokens):
        super(BoolExpr, self).__init__(original_str, location, tokens)
        tokens = tokens[0]
        if ((not hasattr(tokens, "length")) or (len(tokens) > 2)):
            self.lhs = tokens
            try:
                self.lhs = tokens[0]
            # pylint: disable=bare-except
            except:
                pass
            self.op = None
            self.rhs = None
            try:
                self.op = tokens[1]
                self.rhs = BoolExpr(original_str, location, [tokens[2:], None])
            # pylint: disable=bare-except
            except:
                pass

        else:
            self.op = tokens[0]
            self.rhs = tokens[1]
            self.lhs = None

        if (isinstance(self.lhs, pyparsing.ParseResults)):
            self.lhs = BoolExpr(None, None, [self.lhs])
        if (isinstance(self.rhs, pyparsing.ParseResults)):
            self.rhs = BoolExpr(None, None, [self.rhs])

        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed %r as BoolExpr' % self)

    def __repr__(self):
        if (self.op is not None):
            if (self.lhs is not None):
                return self.lhs.__repr__() + " " + self.op + " " + self.rhs.__repr__()
            return self.op + " " + self.rhs.__repr__()
        elif (self.lhs is not None):
            return self.lhs.__repr__()
        log.error("BoolExpr: Improperly parsed.")
        return ""

    def _vba_to_python_op(self, op, context):
        return _vba_to_python_op(op, not context.in_bitwise_expression)
        
    def to_python(self, context, params=None, indent=0):

        start_cast = ""
        end_cast = ""
        if context.in_bitwise_expression:
            start_cast = "coerce_to_int("
            end_cast = ")"

        r = " " * indent + "("
        if (self.op is not None):
            if (self.lhs is not None):
                r += start_cast + to_python(self.lhs, context, params) + end_cast + \
                     " " + self._vba_to_python_op(self.op, context) + " " + \
                     start_cast + to_python(self.rhs, context, params) + end_cast
            else:
                r += self._vba_to_python_op(self.op, context) + " " + \
                     start_cast + to_python(self.rhs, context, params) + end_cast
        elif (self.lhs is not None):
            r += to_python(self.lhs, context, params)
        else:
            log.error("BoolExpr: Improperly parsed.")
            return ""

        r += ")"
        return r
        
    def eval(self, context, params=None):
        params = params # pylint warning
        
        if (self.lhs is None):

            rhs = None
            try:
                rhs = eval_arg(self.rhs, context)
            except Exception as e:
                log.error("BoolExpr: Cannot eval " + self.__repr__() + ". " + str(e))
                return ''

            if ((isinstance(rhs, int)) and (not isinstance(rhs, bool))):
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("Bitwise boolean operation: " + str(self))
                if (self.op.lower() == "not"):
                    # pylint: disable=invalid-unary-operand-type
                    return (~ rhs)
                log.error("BoolExpr: Unknown bitwise unary op " + str(self.op))
                return 0
                
            if (self.op.lower() == "not"):
                return (not rhs)
            log.error("BoolExpr: Unknown boolean unary op " + str(self.op))
            return ''
                
        lhs = self.lhs
        try:
            lhs = eval_arg(self.lhs, context)
        except AttributeError:
            pass

        if (self.op is None):

            return lhs

        rhs = self.rhs
        try:
            rhs = eval_arg(self.rhs, context)
        except AttributeError:
            pass

        if ((isinstance(lhs, int) and isinstance(rhs, int)) and
            (not isinstance(lhs, bool) and not isinstance(rhs, bool))):

            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("Bitwise boolean operation: " + str(self))
            if ((self.op.lower() == "and") or (self.op.lower() == "andalso")):
                return lhs & rhs
            elif ((self.op.lower() == "or") or (self.op.lower() == "orelse")):
                return lhs | rhs
            elif (self.op.lower() == "xor"):
                return lhs ^ rhs

            log.error("BoolExpr: Unknown bitwise operator %r" % self.op)
            return 0
            
        if ((self.op.lower() == "and") or (self.op.lower() == "andalso")):
            return lhs and rhs
        elif ((self.op.lower() == "or") or (self.op.lower() == "orelse")):
            return lhs or rhs
        elif ((self.op.lower() == "eqv") or (self.op.lower() == "=")):
            return (lhs == rhs)

        log.error("BoolExpr: Unknown operator boolean %r" % self.op)
        return False

        
boolean_expression <<= infixNotation(bool_expr_item,
                                     [
                                         (CaselessKeyword("Not"), 1, opAssoc.RIGHT),
                                         (CaselessKeyword("And"), 2, opAssoc.LEFT),
                                         (CaselessKeyword("AndAlso"), 2, opAssoc.LEFT),
                                         (CaselessKeyword("Or"), 2, opAssoc.LEFT),
                                         (CaselessKeyword("OrElse"), 2, opAssoc.LEFT),
                                         (CaselessKeyword("Eqv"), 2, opAssoc.LEFT),
                                         (CaselessKeyword("="), 2, opAssoc.LEFT),
                                     ])
boolean_expression.setParseAction(BoolExpr)


class New_Expression(VBA_Object):
    def __init__(self, original_str, location, tokens):
        super(New_Expression, self).__init__(original_str, location, tokens)
        self.obj = tokens.expression
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed %r as New_Expression' % self)

    def __repr__(self):
        return ('New %r' % self.obj)

    def to_python(self, context, params=None, indent=0):
        context = context # pylint warning
        params = params # pylint warning
        indent = indent # pylint warning
        
        if (str(self.obj).strip().lower() == "regexp"):
            return "core.utils.vb_RegExp()"

        return "ERROR: Not emulating " + str(self)
    
    def eval(self, context, params=None):
        context = context # pylint warning
        params = params # pylint warning
        
        return self.obj

# pylint: disable=expression-not-assigned
new_expression << CaselessKeyword('New').suppress() + expression('expression')
new_expression.setParseAction(New_Expression)

any_expression = expression ^ boolean_expression


class TypeOf_Expression(VBA_Object):
    def __init__(self, original_str, location, tokens):
        super(TypeOf_Expression, self).__init__(original_str, location, tokens)
        self.item = tokens.item
        self.the_type = tokens.the_type
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed %r as TypeOf_Expression' % self)

    def __repr__(self):
        return "TypeOf " + str(self.item) + " Is " + str(self.the_type)

    def eval(self, context, params=None):
        context = context # pylint warning
        params = params # pylint warning
        
        return True


typeof_expression <<= CaselessKeyword("TypeOf") + expression("item") + CaselessKeyword("Is") + expression("the_type")
typeof_expression.setParseAction(TypeOf_Expression)


class AddressOf_Expression(VBA_Object):
    def __init__(self, original_str, location, tokens):
        super(AddressOf_Expression, self).__init__(original_str, location, tokens)
        self.item = tokens.item
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed %r as AddressOf_Expression' % self)

    def __repr__(self):
        return "AddressOf " + str(self.item)

    def eval(self, context, params=None):
        context = context # pylint warning
        params = params # pylint warning
        return "**MATCH ANY**"


addressof_expression <<= CaselessKeyword("AddressOf") + expression("item")
addressof_expression.setParseAction(AddressOf_Expression)


class Excel_Expression(VBA_Object):

    def __init__(self, original_str, location, tokens):
        super(Excel_Expression, self).__init__(original_str, location, tokens)
        self.row = tokens.row
        self.col = tokens.col
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed %r as Excel_Expression' % self)

    def __repr__(self):
        return "[" + str(self.row) + ":" + str(self.col) + "]"

    def eval(self, context, params=None):
        context = context # pylint warning
        params = params # pylint warning
        return "NULL"

    
excel_expression <<= Suppress(Literal("[")) + \
                     lex_identifier("row") + Suppress(Literal(":")) + lex_identifier("col") + \
                     Suppress(Literal("]"))
excel_expression.setParseAction(Excel_Expression)


class Literal_List_Expression(VBA_Object):

    def __init__(self, original_str, location, tokens):
        super(Literal_List_Expression, self).__init__(original_str, location, tokens)
        self.item = tokens.item
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed %r as Literal_List_Expression' % self)

    def __repr__(self):
        return "[" + str(self.item) + "]"

    def eval(self, context, params=None):
        params = params # pylint warning
        if (isinstance(self.item, str)):
            return self.item
        return self.item.eval(context)
    

literal_list_expression <<= Suppress("[") + (unrestricted_name | decimal_literal)("item") + Suppress("]")
literal_list_expression.setParseAction(Literal_List_Expression)


literal_range_expression <<= Suppress(Literal("[")) + decimal_literal + Suppress(Literal(":")) + decimal_literal + Suppress(Literal("]"))
literal_range_expression.setParseAction(lambda t: str(t[0]) + ":" + str(t[1]))

class Tuple_Expression(VBA_Object):

    def __init__(self, original_str, location, tokens):
        super(Tuple_Expression, self).__init__(original_str, location, tokens)
        self.expr_items = tokens.expr_items
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed %r as Tuple_Expression' % self)

    def __repr__(self):
        r = "("
        first = True
        for i in self.expr_items:
            if not first:
                r += ", "
            first = False
            r += str(i)
        r += ")"
        return r

    def eval(self, context, params=None):
        pass

    
tuple_expression <<= Suppress(Literal("(")) + \
                     (expression + OneOrMore(Suppress(Literal(",")) + expression))("expr_items") + \
                     Suppress(Literal(")"))
tuple_expression.setParseAction(Tuple_Expression)
