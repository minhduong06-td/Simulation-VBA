#!/usr/bin/env python
from __future__ import print_function
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

__version__ = '0.04'
import sys
import logging
import string
from pyparsing import *
import prettytable
import unidecode
import string
import subprocess
from logger import log
from procedures import Function
from procedures import Sub
from function_call_visitor import *
from function_defn_visitor import *
from function_import_visitor import *
from var_defn_visitor import *
import filetype
import read_ole_fields
import vba_object
from meta import FakeMeta


def list_startswith(_list, lstart):
    if _list is None:
        return False
    lenlist = len(_list)
    lenstart = len(lstart)
    if lenlist >= lenstart:
        return (_list[:lenstart] == lstart)
    else:
        return False



from vba_lines import *
from modules import *

import vba_library
from vba_library import *

from stubbed_engine import StubbedEngine

def pull_urls_excel_sheets(workbook):
    if (workbook is None):
        return []

    all_cells = excel.pull_cells_workbook(workbook)
    r = set()
    for cell in all_cells:

        value = None
        try:
            value = str(cell["value"]).strip()
        except UnicodeEncodeError:
            value = ''.join(filter(lambda x:x in string.printable, cell["value"])).strip()

        if (len(value) == 0):
            continue
        
        pat = r"[A-Za-z0-9_]{3,50}\.[A-Za-z]{2,10}/(?:[A-Za-z0-9_]{1,50}/)*[A-Za-z0-9_\.]{3,50}"
        if (re.search(pat, value) is not None):
            value = "http://" + value

        for url in re.findall(read_ole_fields.URL_REGEX, value):
            r.add(url.strip())

    return r

def pull_b64_excel_sheets(workbook):
    if (workbook is None):
        return []

    all_cells = excel.pull_cells_workbook(workbook)
    r = set()
    for cell in all_cells:

        value = None
        try:
            value = str(cell["value"]).strip()
        except UnicodeEncodeError:
            value = ''.join(filter(lambda x:x in string.printable, cell["value"])).strip()

        if (len(value) == 0):
            continue

        base64_pat_strict = r"(?:[A-Za-z0-9+/]{4}){10,}(?:[A-Za-z0-9+/]{0,4}=?=?)?"
        for b64 in re.findall(base64_pat_strict, value):
            r.add(b64.strip())

    return r


class SimulationVBA(StubbedEngine):

    def __init__(self, filename, data, do_jit=False):
        self.do_jit = do_jit
        self.comments = None
        self.metadata = None
        self.filename = filename
        self.data = data
        self.modules = []
        self.modules_code = []
        self.globals = {}
        self.externals = {}
        self.actions = []
        self.vba = None

        vba_pointer = self.filename
        is_data = False
        if ((self.filename is None) or (len(self.filename.strip()) == 0)):
            vba_pointer = self.data
            is_data = True
        self.is_vbscript = False
        if (filetype.is_office_file(vba_pointer, is_data)):
            self.is_vbscript = False
            log.info("Emulating an Office (VBA) file.")
        else:
            self.is_vbscript = True
            log.info("Emulating a VBScript file.")

        if (self.is_vbscript == True):
            vba_library.VBA_LIBRARY['vbCrLf'] = '\r\n'
            
        self.loaded_excel = None
        
        self.doc_vars = {}

        self.doc_text = ""

        self.doc_tables = []
        
        self.entry_points = ['autoopen', 'document_open', 'autoclose',
                             'document_close', 'auto_open', 'autoexec',
                             'autoexit', 'document_beforeclose', 'workbook_open',
                             'workbook_activate', 'auto_close', 'workbook_close',
                             'workbook_deactivate', 'documentopen', 'app_documentopen',
                             'main']

        self.callback_suffixes = ['_Activate',
                                  '_BeforeNavigate2',
                                  '_BeforeScriptExecute',
                                  '_Calculate',
                                  '_Change',
                                  '_DocumentComplete',
                                  '_DownloadBegin',
                                  '_DownloadComplete',
                                  '_FileDownload',
                                  '_GotFocus',
                                  '_Layout',
                                  '_LostFocus',
                                  '_MouseEnter',
                                  '_MouseHover',
                                  '_MouseLeave',
                                  '_MouseMove',
                                  '_NavigateComplete2',
                                  '_NavigateError',
                                  '_Painted',
                                  '_Painting',
                                  '_ProgressChange',
                                  '_PropertyChange',
                                  '_Resize',
                                  '_SetSecureLockIcon',
                                  '_StatusTextChange',
                                  '_TitleChange',
                                  '_Initialize',
                                  '_Click',
                                  '_OnConnecting',
                                  '_BeforeClose',
                                  '_OnDisconnected',
                                  '_OnEnterFullScreenMode',
                                  '_Zoom',
                                  '_Scroll',
                                  '_BeforeDropOrPaste']
                                  
    def set_metadata(self, dat):

        new_dat = dat
        if (isinstance(dat, dict)):
            new_dat = FakeMeta()
            for field in list(dat.keys()):
                setattr(new_dat, str(field), dat[field])
        self.metadata = new_dat
        
    def add_compiled_module(self, m):
        if (m is None):
            return
        self.modules.append(m)
        for name, _sub in list(m.subs.items()):
            if (name in self.globals):
                old_sub = self.globals[name]
                if (hasattr(old_sub, "statements")):
                    if (len(_sub.statements) < len(old_sub.statements)):
                        log.warning("Sub " + str(name) + " is already defined. Skipping new definition.")
                        continue
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug('(1) storing sub "%s" in globals' % name)
            self.globals[str(name).lower()] = _sub
            self.globals[name] = _sub
        for name, _function in list(m.functions.items()):
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug('(1) storing function "%s" in globals' % name)
            self.globals[str(name).lower()] = _function
            self.globals[name] = _function
        for name, _prop in list(m.functions.items()):
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug('(1) storing property let "%s" in globals' % name)
            self.globals[str(name).lower()] = _prop
            self.globals[name] = _prop
        for name, _function in list(m.external_functions.items()):
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug('(1) storing external function "%s" in globals' % name)
            self.globals[str(name).lower()] = _function
            self.externals[str(name).lower()] = _function
        for name, _var in list(m.global_vars.items()):
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug('(1) storing global var "%s" = %s in globals (1)' % (name, str(_var)))
            if (isinstance(name, str)):
                self.globals[str(name).lower()] = _var
            if (isinstance(name, list)):
                self.globals[name[0].lower()] = _var
                self.types[name[0].lower()] = name[1]
        
    def add_module(self, vba_code):

        vba_code = vba_collapse_long_lines(vba_code)

        try:
            m = module.parseString(vba_code, parseAll=True)[0]
            m.code = vba_code
            self.add_compiled_module(m)

        except ParseException as err:
            print('*** PARSING ERROR (1) ***')
            print(err.line)
            print(" " * (err.column - 1) + "^")
            print(err)

    def add_module2(self, vba_code):
        vba_code = vba_collapse_long_lines(vba_code)
        self.lines = vba_code.splitlines(True)
        tokens = []
        self.line_index = 0
        while self.lines:
            line_index, line, line_keywords = self.parse_next_line()
            if line_keywords is None:
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug('Empty line or comment: ignored')
                continue
            try:
                pub_priv = False
                if line_keywords[0] in ('public', 'private'):
                    pub_priv = True
                    line_keywords = line_keywords[1:]
                if line_keywords[0] == 'attribute':
                    l = header_statements_line.parseString(line, parseAll=True)
                elif line_keywords[0] in ('option', 'dim', 'declare'):
                    if (log.getEffectiveLevel() == logging.DEBUG):
                        log.debug('DECLARATION LINE')
                    l = declaration_statements_line.parseString(line, parseAll=True)
                elif line_keywords[0] == 'sub':
                    if (log.getEffectiveLevel() == logging.DEBUG):
                        log.debug('SUB')
                    l = sub_start_line.parseString(line, parseAll=True)
                    l[0].statements = self.parse_block(end=['end', 'sub'])
                elif line_keywords[0] == 'function':
                    if (log.getEffectiveLevel() == logging.DEBUG):
                        log.debug('FUNCTION')
                    l = function_start_line.parseString(line, parseAll=True)
                    l[0].statements = self.parse_block(end=['end', 'function'])
                elif line_keywords[0] == 'for':
                    if (log.getEffectiveLevel() == logging.DEBUG):
                        log.debug('FOR LOOP')
                    l = for_start.parseString(line)
                    l[0].statements = self.parse_block(end=['next'])
                else:
                    l = vba_line.parseString(line, parseAll=True)
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug(l)
                tokens.extend(l)
            except ParseException as err:
                print('*** PARSING ERROR (2) ***')
                print(err.line)
                print(" " * (err.column - 1) + "^")
                print(err)
            self.line_index += 1
        m = Module(original_str=vba_code, location=0, tokens=tokens)
        self.modules.append(m)
        for name, _sub in list(m.subs.items()):
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug('(2) storing sub "%s" in globals' % name)
            self.globals[str(name).lower()] = _sub
        for name, _function in list(m.functions.items()):
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug('(2) storing function "%s" in globals' % name)
            self.globals[str(name).lower()] = _function
        for name, _function in list(m.external_functions.items()):
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug('(2) storing external function "%s" in globals' % name)
            self.globals[str(name).lower()] = _function
            self.externals[str(name).lower()] = _function
        for name, _var in list(m.global_vars.items()):
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug('(2) storing global var "%s" in globals (2)' % name)
            
    def parse_next_line(self):
        line = self.lines.pop(0)
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('Parsing line %d: %s' % (self.line_index, line.rstrip()))
        self.line_index += 1
        line_keywords = line.lower().split(None, 2)
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('line_keywords: %r' % line_keywords)
        if len(line_keywords) == 0 or line_keywords[0].startswith("'"):
            return self.line_index-1, line, None
        return self.line_index-1, line, line_keywords

    def parse_block(self, end=['end', 'sub']):
        statements = []
        line_index, line, line_keywords = self.parse_next_line()
        while not list_startswith(line_keywords, end):
            try:
                l = vba_line.parseString(line, parseAll=True)
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug(l)
                statements.extend(l)
            except ParseException as err:
                print('*** PARSING ERROR (3) ***')
                print(err.line)
                print(" " * (err.column - 1) + "^")
                print(err)
            line_index, line, line_keywords = self.parse_next_line()
        return statements

    def _get_external_funcs(self):
        call_visitor = function_call_visitor()
        defn_visitor = function_defn_visitor()
        var_visitor = var_defn_visitor()
        import_visitor = function_import_visitor()
        for module in self.modules:
            module.accept(call_visitor)
            module.accept(defn_visitor)
            module.accept(var_visitor)
            module.accept(import_visitor)

        r = []
        for f in call_visitor.called_funcs:
            if ((f in defn_visitor.funcs) or
                (f in var_visitor.variables) or
                (len(f) == 0) or
                (("." in f) and (not "Shell" in f))):
                continue

            if (f in import_visitor.aliases):
                if (len(import_visitor.funcs[f]) > 0):
                     r.append(import_visitor.funcs[f])
                continue

            r.append(f)

        r.sort()
        return r
        
    def trace(self, entrypoint='*auto'):

        vba_context.intermediate_iocs = set()
        vba_context.num_b64_iocs = 0
        vba_context.shellcode = {}
        
        context = Context(_globals=self.globals,
                          engine=self,
                          doc_vars=self.doc_vars,
                          loaded_excel=self.loaded_excel,
                          filename=self.filename,
                          metadata=self.metadata)
        context.is_vbscript = self.is_vbscript
        context.do_jit = self.do_jit

        fname = self.filename
        is_data = False
        if ((fname is None) or (len(fname.strip()) == 0)):
            fname = self.data
            is_data = True
        direct_urls = read_ole_fields.pull_urls_office97(fname, is_data, self.vba)
        for url in direct_urls:
            context.save_intermediate_iocs(url)
        direct_urls = pull_urls_excel_sheets(self.loaded_excel)
        for url in direct_urls:
            context.save_intermediate_iocs(url)

        cell_b64_blobs = pull_b64_excel_sheets(self.loaded_excel)
        for cell_b64_blob in cell_b64_blobs:
            context.save_intermediate_iocs(cell_b64_blob)
            
        for func_name in list(self.externals.keys()):
            func = self.externals[func_name]
            context.dll_func_true_names[func.name] = func.alias_name

        context.globals["__DOC_TABLE_CONTENTS__"] = self.doc_tables
            
        context.globals["Range.Text".lower()] = "\n".join(self.doc_text)
        context.globals["Me.Content".lower()] = "\n".join(self.doc_text)
        context.globals["Me.Content.Text".lower()] = "\n".join(self.doc_text)
        context.globals["Me.Range.Text".lower()] = "\n".join(self.doc_text)
        context.globals["Me.Range".lower()] = "\n".join(self.doc_text)
        context.globals["Me.Content.Start".lower()] = 0
        context.globals["Me.Content.End".lower()] = len("\n".join(self.doc_text))
        context.globals["Me.Paragraphs".lower()] = self.doc_text
        context.globals["ActiveDocument.Content".lower()] = "\n".join(self.doc_text)
        context.globals["ActiveDocument.Content.Text".lower()] = "\n".join(self.doc_text)
        context.globals["ActiveDocument.Range.Text".lower()] = "\n".join(self.doc_text)
        context.globals["ActiveDocument.Range".lower()] = "\n".join(self.doc_text)
        context.globals["ActiveDocument.Content.Start".lower()] = 0
        context.globals["ActiveDocument.Content.End".lower()] = len("\n".join(self.doc_text))
        context.globals["ActiveDocument.Paragraphs".lower()] = self.doc_text
        context.globals["ThisDocument.Content".lower()] = "\n".join(self.doc_text)
        context.globals["ThisDocument.Content.Text".lower()] = "\n".join(self.doc_text)
        context.globals["ThisDocument.Range.Text".lower()] = "\n".join(self.doc_text)
        context.globals["ThisDocument.Range".lower()] = "\n".join(self.doc_text)
        context.globals["ThisDocument.Content.Start".lower()] = 0
        context.globals["ThisDocument.Content.End".lower()] = len("\n".join(self.doc_text))
        context.globals["ThisDocument.Paragraphs".lower()] = self.doc_text
        context.globals["['Me'].Content.Text".lower()] = "\n".join(self.doc_text)
        context.globals["['Me'].Range.Text".lower()] = "\n".join(self.doc_text)
        context.globals["['Me'].Range".lower()] = "\n".join(self.doc_text)
        context.globals["['Me'].Content.Start".lower()] = 0
        context.globals["['Me'].Content.End".lower()] = len("\n".join(self.doc_text))
        context.globals["['Me'].Paragraphs".lower()] = self.doc_text
        context.globals["['ActiveDocument'].Content.Text".lower()] = "\n".join(self.doc_text)
        context.globals["['ActiveDocument'].Range.Text".lower()] = "\n".join(self.doc_text)
        context.globals["['ActiveDocument'].Range".lower()] = "\n".join(self.doc_text)
        context.globals["['ActiveDocument'].Content.Start".lower()] = 0
        context.globals["['ActiveDocument'].Content.End".lower()] = len("\n".join(self.doc_text))
        context.globals["['ActiveDocument'].Paragraphs".lower()] = self.doc_text
        context.globals["['ThisDocument'].Content.Text".lower()] = "\n".join(self.doc_text)
        context.globals["['ThisDocument'].Range.Text".lower()] = "\n".join(self.doc_text)
        context.globals["['ThisDocument'].Range".lower()] = "\n".join(self.doc_text)
        context.globals["['ThisDocument'].Content.Start".lower()] = 0
        context.globals["['ThisDocument'].Content.End".lower()] = len("\n".join(self.doc_text))
        context.globals["['ThisDocument'].Paragraphs".lower()] = self.doc_text
        context.globals["['ActiveDocument'].Characters".lower()] = list("\n".join(self.doc_text))
        context.globals["ActiveDocument.Characters".lower()] = list("\n".join(self.doc_text))
        context.globals["ActiveDocument.Characters.Count".lower()] = long(len(self.doc_text))
        context.globals["Count".lower()] = 1
        context.globals[".Pages.Count".lower()] = 1
        context.globals["me.Pages.Count".lower()] = 1
        context.globals["['ThisDocument'].Characters".lower()] = list("\n".join(self.doc_text))
        context.globals["ThisDocument.Characters".lower()] = list("\n".join(self.doc_text))
        context.globals["ThisDocument.Sections".lower()] = list("\n".join(self.doc_text))
        context.globals["ActiveDocument.Sections".lower()] = list("\n".join(self.doc_text))

        doc_words = []
        for word in re.split(r"[ \n]", "\n".join(self.doc_text)):
            word = word.strip()
            if (word.startswith("-")):
                word = word[1:]
                doc_words.append("-")
            doc_words.append(word.strip())
        context.globals["ActiveDocument.Words".lower()] = doc_words
        context.globals["ThisDocument.Words".lower()] = doc_words
            
        if (self.comments is None):
            context.globals["ActiveDocument.Comments".lower()] = ["Comment 1", "Comment 2"]
            context.globals["ThisDocument.Comments".lower()] = ["Comment 1", "Comment 2"]
        else:
            context.globals["ActiveDocument.Comments".lower()] = self.comments
            context.globals["ThisDocument.Comments".lower()] = self.comments
            if (self.metadata is not None):
                all_comments = ""
                for comment in self.comments:
                    all_comments += comment + "/n"
                self.metadata.comments = all_comments
            
        self.actions = []

        self.external_funcs = self._get_external_funcs()
        context.external_funcs = self.external_funcs

        log.info("Emulating loose statements...")
        done_emulation = False
        for m in self.modules:
            if (m.eval(context=context)):
                context.dump_all_files(autoclose=True)
                done_emulation = context.got_actions
        
        for entry_point in self.entry_points:
            entry_point = entry_point.lower()
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("Trying entry point " + entry_point)
            if ((entry_point in self.globals) and
                (hasattr(self.globals[entry_point], "eval"))):
                context.report_action('Found Entry Point', str(entry_point), '')
                tmp_context = Context(context=context, _locals=context.locals, copy_globals=True)
                self.globals[entry_point].eval(context=tmp_context)
                tmp_context.dump_all_files(autoclose=True)
                context.got_actions = tmp_context.got_actions
                done_emulation = True

        for name in list(self.globals.keys()):

            for suffix in self.callback_suffixes:

                if (str(name).lower().endswith(suffix.lower())):

                    item = self.globals[name]
                    if (isinstance(item, Function) or isinstance(item, Sub)):

                        context.report_action('Found Entry Point', str(name), '')
                        tmp_context = Context(context=context, _locals=context.locals, copy_globals=True)
                        item.eval(context=tmp_context)
                        tmp_context.dump_all_files(autoclose=True)
                        context.got_actions = tmp_context.got_actions

        if (not done_emulation):

            log.warn("No entry points found. Using heuristics to find entry points...")
            
            zero_arg_subs = []
            for name in list(self.globals.keys()):
                item = self.globals[name]
                if ((isinstance(item, Sub)) and (len(item.params) == 0)):
                    zero_arg_subs.append(item)
                    
            for only_sub in zero_arg_subs:
                sub_name = only_sub.name
                context.report_action('Found Heuristic Entry Point', str(sub_name), '')
                only_sub.eval(context=context)
                context.dump_all_files(autoclose=True)
                
    def eval(self, expr):
        context = Context(_globals=self.globals,
                          engine=self,
                          doc_vars=self.doc_vars,
                          loaded_excel=self.loaded_excel)
        self.actions = []
        e = expression.parseString(expr)[0]
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('e=%r - type=%s' % (e, type(e)))
        value = e.eval(context=context)
        return value

    def dump_actions(self):
        t = prettytable.PrettyTable(('Action', 'Parameters', 'Description'))
        t.align = 'l'
        t.max_width['Action'] = 20
        t.max_width['Parameters'] = 25
        t.max_width['Description'] = 25
        for action in self.actions:
            str_action = str(action)
            if (len(str_action) > 50000):
                new_params = str(action[1])
                if (len(new_params) > 50000):
                    new_params = new_params[:25000] + "... <SNIP> ..." + new_params[-25000:]
                action = (action[0], new_params, action[2])
            t.add_row(action)
        return t

def scan_expressions(vba_code):
    context = Context()
    for m in expr_const.scanString(vba_code):
        e = m[0][0]
        if hasattr(e, 'eval'):
            yield (e, e.eval(context))


