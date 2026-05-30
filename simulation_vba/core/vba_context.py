


__version__ = '0.08'


import logging
import os
from hashlib import sha256
from datetime import datetime
from logger import log
import re
try:
    import rure as re2
except:
    import re as re2
import random
import string
import codecs
import copy
import struct
from curses_ascii import isascii

import vba_constants
import utils

def to_hex(s):
    """
    Convert a string to a VBA hex string.
    """

    r = ""
    for c in str(s):
        r += hex(ord(c)).replace("0x", "")
    return r

def is_procedure(vba_object):
    """
    Check if a VBA object is a procedure, e.g. a Sub or a Function.
    This is implemented by checking if the object has a statements
    attribute
    :param vba_object: VBA_Object to be checked
    :return: True if vba_object is a procedure, False otherwise
    """
    if hasattr(vba_object, 'statements'):
        return True
    else:
        return False

def add_shellcode_data(index, value, num_bytes):
    """
    Save injected shellcode data.
    """

    if ((not isinstance(index, int)) or
        (not isinstance(value, int)) or
        (not isinstance(num_bytes, int))):
        log.warning("Improperly typed argument passed to add_shellcode_data(). Skipping.")
        return

    if (num_bytes > 1):
        log.warning("Only handling single byte values in add_shellcode_data(). Skipping.")
        return
    
    shellcode[index] = value

def get_shellcode_data():
    """
    Get written shellcode bytes as a list.
    """

    if (len(shellcode) == 0):
        return []

    indices = shellcode.keys()
    indices.sort()
    last_i = None
    r = []
    for i in indices:

        if ((last_i is not None) and (last_i + 1 != i)):
            last_i += 1
            while (last_i != i):
                r.append(0x90)
                last_i += 1

        curr_val = shellcode[i]
        if (curr_val < 0):
            curr_val += 2**8
                    
        r.append(curr_val)
        last_i = i

    return r
    

VBA_LIBRARY = {}

out_dir = None

intermediate_iocs = set()

num_b64_iocs = 0

shellcode = {}

class Context(object):
    """
    a Context object contains the global and local named objects (variables, subs, functions)
    used to evaluate VBA statements.
    """

    def __init__(self,
                 _globals=None,
                 _locals=None,
                 context=None,
                 engine=None,
                 doc_vars=None,
                 loaded_excel=None,
                 filename=None,
                 copy_globals=False,
                 log_funcs=None,
                 expand_env_vars=True,
                 metadata=None):

        self.name_cache = {}

        self.curr_func_name = None
        
        self.last_saved_file = None
        
        self.in_bitwise_expression = False
        
        self.got_actions = False
        
        self.external_funcs = []

        self.has_change_handler = {}
        
        self.call_stack = []
        
        self.max_static_iters = 2

        self.is_vbscript = False

        self.do_jit = False

        self.throttle_logging = False
        
        if log_funcs:
            self._log_funcs = [func_name.lower() for func_name in log_funcs]
        else:
            self._log_funcs = []

        self.expand_env_vars = expand_env_vars
        
        self.skip_handlers = set()
        
        self.filename = filename
        
        self.got_error = False

        self.error_handler = None

        self.num_general_errors = 0
        
        self.dll_func_true_names = {}
        
        self.tagged_blocks = {}

        self.loaded_excel = loaded_excel
        
        self.open_files = {}
        self.file_id_map = {}

        self.closed_files = {}

        self.metadata = metadata
        
        self.global_scope = False

        self.in_procedure = False

        self.goto_executed = False

        self.types = {}

        self.with_prefix = ""
        self.with_prefix_raw = None
        
        if _globals is not None:
            if (copy_globals):
                self.globals = copy.deepcopy(_globals)
            else:
                self.globals = _globals

            for var in _globals.keys():
                self.save_intermediate_iocs(_globals[var])
                
        elif context is not None:
            if (copy_globals):
                self.globals = dict(context.globals)
            else:
                self.globals = context.globals
            self.in_bitwise_expression = context.in_bitwise_expression
            self.last_saved_file = context.last_saved_file
            self.curr_func_name = context.curr_func_name
            self.do_jit = context.do_jit
            self.has_change_handler = context.has_change_handler
            self.throttle_logging = context.throttle_logging
            self.is_vbscript = context.is_vbscript
            self.doc_vars = context.doc_vars
            self.types = dict(context.types)
            self.open_files = context.open_files
            self.file_id_map = context.file_id_map
            self.closed_files = context.closed_files
            self.loaded_excel = context.loaded_excel
            self.dll_func_true_names = context.dll_func_true_names
            self.filename = context.filename
            self.skip_handlers = context.skip_handlers
            self.call_stack = context.call_stack
            self.expand_env_vars = context.expand_env_vars
            self.metadata = context.metadata
            self.external_funcs = context.external_funcs
            self.num_general_errors = context.num_general_errors
            self.with_prefix = context.with_prefix
            self.with_prefix_raw = context.with_prefix_raw
        else:
            self.globals = {}
        if _locals is not None:
            self.locals = dict(_locals)
        else:
            self.locals = {}
        if engine is not None:
            self.engine = engine
        elif context is not None:
            self.engine = context.engine
        else:
            self.engine = None

        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Have xlrd loaded Excel file = " + str(self.loaded_excel is not None))
            
        if doc_vars is not None:

            self.doc_vars = doc_vars

            for var in doc_vars.keys():
                self.save_intermediate_iocs(doc_vars[var])

        elif context is not None:
            self.doc_vars = context.doc_vars
        else:
            self.doc_vars = {}
            
        self.loop_stack = []

        self.loop_object_stack = []
        
        self.exit_func = False
        
        self.globals["Now".lower()] = datetime.now()

        rand_name = ''.join(random.choice(string.ascii_uppercase + string.digits + " ") for _ in range(random.randint(10, 50)))
        self.globals["Application.UserName".lower()] = rand_name

        self.globals["ActiveDocument.AttachedTemplate.Path".lower()] = "C:\\Users\\" + rand_name + "\\AppData\\Roaming\\Microsoft\\Templates"
        self.globals["ThisDocument.AttachedTemplate.Path".lower()] = "C:\\Users\\" + rand_name + "\\AppData\\Roaming\\Microsoft\\Templates"

        if self.filename:
            self.globals["WSCRIPT.SCRIPTFULLNAME".lower()] = "C:\\" + self.filename
            self.globals["['WSCRIPT'].SCRIPTFULLNAME".lower()] = "C:\\" + self.filename
        
    def __repr__(self):
        r = ""
        r += "Locals:\n"
        r += str(self.locals) + "\n\n"
        return r
        
    def __eq__(self, other):
        if isinstance(other, Context):
            globals_eq = (self.globals == other.globals)
            if (not globals_eq):
                s1 = set()
                for i in self.globals.items():
                    s1.add(str(i))
                s2 = set()
                for i in other.globals.items():
                    s2.add(str(i))
                if (str(s1 ^ s2) == "set([])"):
                    globals_eq = True
            return ((self.call_stack == other.call_stack) and
                    globals_eq and
                    (self.locals == other.locals))
        return NotImplemented

    def __ne__(self, other):
        result = self.__eq__(other)
        if result is NotImplemented:
            return result
        return not result
        
    def read_metadata_item(self, var):

        if (self.metadata is None):
            log.error("BuiltInDocumentProperties: Metadata not read.")
            return ""
    
        var = var.lower().replace(" ", "_")
        if ("." in var):
            var = var[:var.index(".")]
    
        if (not hasattr(self.metadata, var)):
            log.error("BuiltInDocumentProperties: Metadata field '" + var + "' not found.")
            return ""

        r = getattr(self.metadata, var)

        r = r.replace("_x000d_.", "\r\n")
        r = r.replace("_x000d_", "\r")
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("BuiltInDocumentProperties: return %r -> %r" % (var, r))

        return r
            
    def get_error_handler(self):
        """
        Get the onerror goto error handler.
        """
        if (hasattr(self, "error_handler")):
            return self.error_handler
        return None

    def do_next_iter_on_error(self):
        """
        See if the error handler just calls Next to advance to next loop iteration.
        """

        handler = self.get_error_handler()
        if (handler is None):
            return False

        if (len(handler.block) == 0):

            return True
        first_cmd = str(handler.block[0]).strip()
        return (first_cmd == "Next")
    
    def have_error(self):
        """
        See if Visual Basic threw an error.
        """
        return (hasattr(self, "got_error") and
                self.got_error)

    def clear_error(self):
        """
        Clear out the error flag.
        """
        self.got_error = False
        
    def must_handle_error(self):
        """
        Check to see if there was are error raised during emulation and we have
        an error handler.
        """
        return (self.have_error() and
                hasattr(self, "error_handler") and
                (self.error_handler is not None))

    def handle_error(self, params):
        """
        Run the current error handler (if there is one) if there is an error.
        """

        if (self.must_handle_error()):
            log.warning("Running On Error error handler...")
            self.got_error = False
            self.error_handler.eval(context=self, params=params)

            self.got_error = False

    def set_error(self, reason):
        """
        Set that a VBA error has occurred.
        """

        self.got_error = True
        self.increase_general_errors()
        log.error("A VB error has occurred. Reason: " + str(reason))

    def report_general_error(self, reason):
        """
        Report and track general SimulationVBA errors. Note that these may not just be
        VBA errors.
        """
        self.num_general_errors += 1
        log.error(reason)

    def clear_general_errors(self):
        """
        Clear the count of general errors.
        """
        self.num_general_errors = 0

    def get_general_errors(self):
        """
        Get the number of reported general errors.
        """
        return self.num_general_errors

    def increase_general_errors(self):
        """
        Add one to the number of reported general errors.
        """
        self.num_general_errors += 1
        
    def get_true_name(self, name):
        """
        Get the true name of an aliased function imported from a DLL.
        """
        if (name in self.dll_func_true_names):
            return self.dll_func_true_names[name]
        return None

    def delete(self, name):
        """
        Delete a variable from the context.
        """

        if (not self.contains(name)):
            return self

        if name in self.locals:
            del self.locals[name]
        elif name in self.globals:
            del self.globals[name]

        return self

    def get_interesting_fileid(self):
        """
        Pick an 'interesting' looking open file and return its ID.
        """

        longest = ""
        cdrive = None
        for file_id in self.open_files.keys():
            if ((self.last_saved_file is not None) and (str(file_id).lower() == self.last_saved_file.lower())):
                cdrive = file_id
                break
            if (str(file_id).lower().startswith("c:")):
                cdrive = file_id
            if (len(str(file_id)) > len(longest)):
                longest = file_id

        if (cdrive is not None):
            return cdrive

        if (len(longest) > 0):
            return longest

        return None

    def file_is_open(self, fname):
        """
        Check to see if a file is already open.
        """
        fname = str(fname)
        fname = fname.replace(".\\", "").replace("\\", "/")

        return (fname in self.open_files.keys())
        
    def open_file(self, fname, file_id=""):
        """
        Simulate opening a file.

        fname - The name of the file.
        file_id - The numeric ID of the file.
        """
        fname = str(fname)
        fname = fname.replace(".\\", "").replace("\\", "/")

        if (fname in self.open_files.keys()):
            log.warning("File " + str(fname) + " is already open.")
            return

        self.open_files[fname] = b''
        if (file_id != ""):
            self.file_id_map[file_id] = fname
        log.info("Opened file " + fname)
        
    def write_file(self, fname, data):

        fname = str(fname)
        fname = fname.replace(".\\", "").replace("\\", "/")
        if fname not in self.open_files:

            if (fname in self.file_id_map.keys()):
                fname = self.file_id_map[fname]
            else:

                got_it = False
                if fname.startswith("#"):
                    var_name = fname[1:]
                    if self.contains(var_name):
                        fname = "#" + str(self.get(var_name))
                        if (fname in self.file_id_map.keys()):
                            got_it = True
                            fname = self.file_id_map[fname]
                
                if (not got_it):
                    log.error('File {} not open. Cannot write new data.'.format(fname))
                    return False
            
        if isinstance(data, str):

            if ((len(data.strip()) == 4) and (re.match('&H[0-9A-F]{2}', data, re.IGNORECASE))):
                data = chr(int(data.strip()[-2:], 16))

            self.open_files[fname] += data
            return True

        elif isinstance(data, list):
            for d in data:
                if (isinstance(d, int)):
                    self.open_files[fname] += chr(d)
                else:
                    self.open_files[fname] += str(d)
            return True

        elif isinstance(data, int):

            byte_list = struct.pack('<q', data)
            
            byte_size = utils.get_num_bytes(data)
            byte_list = byte_list[:byte_size]
            
            for b in byte_list:
                self.open_files[fname] += b
            return True
        
        else:
            log.error("Unhandled data type to write. " + str(type(data)) + ".")
            return False
        
    def dump_all_files(self, autoclose=False):
        for fname in self.open_files.keys():
            self.dump_file(fname, autoclose=autoclose)

    def get_num_open_files(self):
        """
        Get the # of currently open files being tracked.
        """
        return len(self.open_files)
            
    def close_file(self, fname):
        """
        Simulate closing a file.

        fname - The name of the file.

        Returns boolean indicating success.
        """
        global file_count
        
        fname = str(fname).replace(".\\", "").replace("\\", "/")
        file_id = None
        if fname not in self.open_files:

            if (fname in self.file_id_map.keys()):
                file_id = fname
                fname = self.file_id_map[fname]
            else:
                log.error('File {} not open. Cannot close.'.format(fname))
                return

        log.info("Closing file " + fname)

        data = self.open_files[fname]
        self.closed_files[fname] = data

        del self.open_files[fname]
        if (file_id is not None):
            del self.file_id_map[file_id]

        if out_dir:
            self.dump_file(fname)

    def dump_file(self, fname, autoclose=False):
        """
        Save the contents of a file dumped by the VBA to disk.

        fname - The name of the file.
        """
        if fname not in self.closed_files:
            if (not autoclose):
                log.error('File {} not closed. Cannot save.'.format(fname))
                return
            else:
                log.warning('File {} not closed. Closing file.'.format(fname))
                self.close_file(fname)
                
        raw_data = self.closed_files[fname]
        file_hash = sha256(raw_data).hexdigest()


        if not os.path.isdir(out_dir):
            os.makedirs(out_dir)

        try:
            fname = re.sub(r"[^ -~\r\n]", "__", fname)
            if ("/" in fname):
                fname = fname[fname.rindex("/") + 1:]
            if ("\\" in fname):
                fname = fname[fname.rindex("\\") + 1:]
            fname = fname.replace("\x00", "").replace("..", "")
            if (fname.startswith(".")):
                fname = "_dot_" + fname[1:]

            if (len(fname) > 50):
                fname = "REALLY_LONG_NAME_" + str(file_hash) + ".dat"
                log.warning("Filename of dropped file is too long, replacing with " + fname)

            fname = fname.strip()
            self.report_action("Dropped File Hash", file_hash, 'File Name: ' + fname)
            file_path = os.path.join(out_dir, os.path.basename(fname))
            orig_file_path = file_path
            count = 0
            while os.path.exists(file_path):
                count += 1
                file_path = '{} ({})'.format(orig_file_path, count)

            with open(file_path, 'wb') as f:
                f.write(raw_data)
            log.info("Wrote dumped file (hash {}) to {}.".format(file_hash, file_path))
        except Exception as e:
            log.error("Writing file {} failed with error: {}".format(fname, e))

    def get_lib_func(self, name):

        if (not isinstance(name, basestring)):
            raise KeyError('Object %r not found' % name)
        
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Looking for library function '" + name + "'...")
        name = name.lower()
        if name in VBA_LIBRARY:
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug('Found %r in VBA Library' % name)
            return VBA_LIBRARY[name]

        else:            
            raise KeyError('Library function %r not found' % name)

    def __get(self, name, case_insensitive=True, local_only=False, global_only=False):

        if (not isinstance(name, basestring)):
            raise KeyError('Object %r not found' % name)

        is_change_handler = (str(name).strip().lower().endswith("_change"))
        change_name = str(name).strip().lower()
        if is_change_handler: change_name = change_name[:-len("_change")]
        
        orig_name = name
        if (case_insensitive):
            name = name.lower()
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Looking for var '" + name + "'...")

        if (name.strip().endswith(".subfolders.count")):
            return -1
        
        if ((not global_only) and (name in self.locals)):
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug('Found %r in locals' % name)
            if is_change_handler: self.has_change_handler[change_name] = True
            self.name_cache[orig_name] = name
            return self.locals[name]

        elif ((not local_only) and (name in self.globals)):
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug('Found %r in globals' % name)
            if is_change_handler: self.has_change_handler[change_name] = True
            self.name_cache[orig_name] = name
            return self.globals[name]

        elif ((not local_only) and (name in VBA_LIBRARY)):
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug('Found %r in VBA Library' % name)
            if is_change_handler: self.has_change_handler[change_name] = True
            self.name_cache[orig_name] = name
            return VBA_LIBRARY[name]

        elif ((not local_only) and (name in self.doc_vars)):
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug('Found %r in VBA document variables' % name)
            if is_change_handler: self.has_change_handler[change_name] = True
            self.name_cache[orig_name] = name
            return self.doc_vars[name]

        elif vba_constants.is_constant(name):
            return vba_constants.get_constant(name)
        
        else:
            if is_change_handler: self.has_change_handler[change_name] = False
            raise KeyError('Object %r not found' % name)

    def _get(self, name, search_wildcard=True, case_insensitive=True, local_only=False, global_only=False):
        
        name = str(name)
        if (((name.lower() == "nodetypedvalue") or (name.lower() == ".nodetypedvalue")) and
            (not name in self.locals) and
            (".Text".lower() in self.locals)):
            return self.get(".Text")

        if (name in self.name_cache):
            cached_name = self.name_cache[name]
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("Cached name of '" + str(name) + "' is '" + str(cached_name) + "'")
            try:
                return self.__get(cached_name,
                                  case_insensitive=case_insensitive,
                                  local_only=local_only,
                                  global_only=global_only)
            except KeyError:
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("Cached lookup failed.")

        with_prefix = self.with_prefix
        if ((self.with_prefix_raw is not None) and
            (str(self.with_prefix_raw).startswith("ActiveDocument"))):
            with_prefix = self.with_prefix_raw
                    
        if (name.startswith(".")):
            
            tmp_name = str(with_prefix) + str(name)
            try:
                return self.__get(tmp_name,
                                  case_insensitive=case_insensitive,
                                  local_only=local_only,
                                  global_only=global_only)
            except KeyError:

                tmp_name = str(self.with_prefix) + str(name)
                try:
                    return self.__get(tmp_name,
                                      case_insensitive=case_insensitive,
                                      local_only=local_only,
                                      global_only=global_only)
                except KeyError:
                    pass

        try:
            return self.__get(str(name),
                              case_insensitive=case_insensitive,
                              local_only=local_only,
                              global_only=global_only)
        except KeyError:
            pass

        tmp_name = str(with_prefix) + "." + str(name)
        try:
            return self.__get(tmp_name,
                              case_insensitive=case_insensitive,
                              local_only=local_only,
                              global_only=global_only)
        except KeyError:

            if (isinstance(self.with_prefix, str) and
                (self.with_prefix_raw is not None) and
                ("Shapes" in str(self.with_prefix_raw)) and
                (str(name) == "Title")):
                return self.with_prefix
        
        if ("." in name):

            new_name = "me." + name[name.index(".")+1:]
            try:
                return self.__get(str(new_name),
                                  case_insensitive=case_insensitive,
                                  local_only=local_only,
                                  global_only=global_only)
            except KeyError:
                pass

            if (search_wildcard):
                new_name = name[:name.index(".")] + ".*"
                try:
                    r = self.__get(str(new_name),
                                   case_insensitive=case_insensitive,
                                   local_only=local_only,
                                   global_only=global_only)
                    if (log.getEffectiveLevel() == logging.DEBUG):
                        log.debug("Found wildcarded field value " + new_name + " = " + str(r))
                    return r
                except KeyError:
                    pass
            
        return self.__get(str(name) + "$",
                          case_insensitive=case_insensitive,
                          local_only=local_only,
                          global_only=global_only)

    def _get_all_metadata(self, name):
        """
        Return all items in something like ActiveDocument.BuiltInDocumentProperties.
        """

        if ((name != "ActiveDocument.BuiltInDocumentProperties") and
            (name != "ThisDocument.BuiltInDocumentProperties")):
            return None

        meta_names = [a for a in dir(self.metadata) if not a.startswith('__') and not callable(getattr(self.metadata, a))]

        for meta_name in meta_names:
            self.set(meta_name + ".Name", meta_name, force_global=True)
            self.set(meta_name + ".Value", getattr(self.metadata, meta_name), force_global=True)
            self.save_intermediate_iocs(getattr(self.metadata, meta_name))

        meta_names.append("Comments")
        comments = ""
        first = True
        for comment in self.get("ActiveDocument.Comments"):
            if (not first):
                comments += "\n"
            first = False
            comments += comment
        self.set("Comments.Name", "Comments", force_global=True)
        self.set("Comments.Value", comments, force_global=True)
        self.save_intermediate_iocs(comments)
        
        return meta_names
    
    def get(self, name, search_wildcard=True, local_only=False, global_only=False):

        if ((name is None) or
            (isinstance(name, str) and (len(name.strip()) == 0))):
            raise KeyError('Object %r not found' % name)
        
        if (str(name).strip().lower().endswith("_change")):

            orig_name = str(name).strip().lower()[:-len("_change")]
            if ((orig_name in self.has_change_handler) and (not self.has_change_handler[orig_name])):
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("Short circuited change handler lookup of " + name)
                raise KeyError('Object %r not found' % name)

        r = self._get_all_metadata(name)
        if (r is not None):
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("Read all metadata items.")
            return r
            
        r = None
        try:
            r = self._get(name,
                          search_wildcard=search_wildcard,
                          case_insensitive=False,
                          local_only=local_only,
                          global_only=global_only)
        except KeyError:
            r = self._get(name,
                          search_wildcard=search_wildcard,
                          case_insensitive=True,
                          local_only=local_only,
                          global_only=global_only)

        if ((r is None) or (r == "NULL")):

            tmp_name = "." + str(name)
            if (self.contains(tmp_name)):
                r = self._get(tmp_name)
            
        return r
            
    def contains(self, name, local=False):
        if (local):
            return (str(name).lower() in self.locals)
        try:
            self.get(name)
            return True
        except KeyError:
            return False

    def contains_user_defined(self, name):
        return ((name in self.locals) or (name in self.globals))

    def set_type(self, var, typ):
        var = var.lower()
        self.types[var] = typ
        
    def get_type(self, var):
        if (not isinstance(var, basestring)):
            return None
        var = var.lower()
        if (var not in self.types):
            return vba_constants.get_type(var)
        return self.types[var]

    def get_doc_var(self, var, search_wildcard=True):
        if (not isinstance(var, basestring)):
            return None

        var = var.lower()
        var = var.replace('!','').\
                    replace('^','').\
                    replace('%','').\
                    replace('&','').\
                    replace('@','').\
                    replace('#','').\
                    replace('$','')
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Looking up doc var " + var)

        if (var == "activedocument.variables"):

            r = []
            for var_name in self.doc_vars.keys():
                r.append((var_name, self.doc_vars[var_name]))                
            return r
        
        if (var not in self.doc_vars):

            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("doc var named " + var + " not found.")
            try:
                var_value = self.get(var, search_wildcard=search_wildcard)
                if ((var_value is not None) and
                    (str(var_value).lower() != str(var).lower())):
                    r = self.get_doc_var(var_value)
                    if (r is not None):
                        return r
                    return var_value
            except KeyError:
                pass

            if ((re.match(r"^[a-zA-Z_][\w\d]*$", str(var)) is not None) and
                ("*" in self.doc_vars)):
                return self.doc_vars["*"]

            if ("." in var):

                var = "activedocument." + var[var.index(".") + 1:]
                if (var in self.doc_vars):

                    r = self.doc_vars[var]
                    if (log.getEffectiveLevel() == logging.DEBUG):
                        log.debug("Found doc var " + var + " = " + str(r))
                    return r
                
            return None

        r = self.doc_vars[var]
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Found doc var " + var + " = " + str(r))
        return r

    def set_excel_selection(self, target):
        """Resolve Excel Application.Goto/Range targets into Selection.

        This is a best-effort bridge between static OOXML workbook data and the
        lightweight emulator. The OOXML loader stores named ranges and cell
        values under internal __excel_* keys; this method consumes those keys
        and updates common runtime variables such as Selection and ActiveCell.
        """
        if target is None:
            return None
        if isinstance(target, (list, tuple)) and len(target) > 0:
            target = target[0]
        target = str(target).strip()
        target = target.strip("[]")
        target = target.strip("'\"")
        if len(target) == 0:
            return None

        norm_target = target.replace("$", "")
        keys = [
            "__excel_defined_name." + target.lower(),
            "__excel_defined_name." + norm_target.lower(),
            "__excel_cell." + norm_target.lower(),
            norm_target.lower(),
            target.lower(),
        ]
        value = None
        resolved_key = None
        for key in keys:
            try:
                value = self.get(key, search_wildcard=False)
                resolved_key = key
                break
            except KeyError:
                pass

        if value is None:
            value = target
            log.info(
                "Application.Goto target %r not found in OOXML context keys; "
                "using target directly as Selection value.",
                target,
            )

        if isinstance(value, str) and "!" in value and "$" in value:
            from ooxml_context import _resolve_defined_name_value
            try:
                cell_ref = value.replace("$", "")
                sheet_cell_key = "__excel_cell." + cell_ref.lower()
                resolved = self.get(sheet_cell_key, search_wildcard=False)
                if resolved is not None:
                    log.info(
                        "Resolved defined-name target %r -> cell value (starts %r, %d chars).",
                        value, str(resolved)[:30], len(str(resolved)),
                    )
                    value = resolved
            except KeyError:
                pass

        log.info(
            "Application.Goto resolved %r -> Selection (starts %r, %d chars).",
            target, str(value)[:50], len(str(value)),
        )

        for name in ("Selection", "ActiveCell", "Application.Selection",
                     "Application.ActiveCell", "ActiveWindow.Selection"):
            self.set(name, value, force_global=True)
            self.set(name + ".Value", value, force_global=True)
            self.set(name + ".Value2", value, force_global=True)
            self.set(name + ".Text", value, force_global=True)
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Application.Goto resolved %r to Selection = %r" % (target, value))
        return value

    def save_intermediate_iocs(self, value):
        """
        Save variable values that appear to contain base64 encoded or URL IOCs.
        """

        global num_b64_iocs
        
        value = utils.strip_nonvb_chars(value)
        if (len(re.findall(r"NULL", str(value))) > 20):
            value = str(value).replace("NULL", "")

        got_ioc = False
        URL_REGEX = r'.*([hH][tT][tT][pP][sS]?://(([a-zA-Z0-9_\-]+\.[a-zA-Z0-9_\-\.]+(:[0-9]+)?)+(/([/\?&\~=a-zA-Z0-9_\-\.](?!http))+)?)).*'
        try:
            value = str(value).strip()
        except:
            return
        tmp_value = value
        if (len(tmp_value) > 100):
            tmp_value = tmp_value[:100] + " ..."
        if (re.match(URL_REGEX, value) is not None):
            if (value not in intermediate_iocs):
                got_ioc = True
                log.info("Found possible intermediate IOC (URL): '" + tmp_value + "'")

        if ((num_b64_iocs < 200) and (value not in intermediate_iocs)):
            uni_value = None
            try:
                uni_value = value.decode("utf-8")
            except UnicodeDecodeError:
                pass
            if (uni_value is not None):
                B64_REGEX = r"(?:[A-Za-z0-9+/]{4}){10,}(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?"
                b64_strs = re2.findall(unicode(B64_REGEX), uni_value)
                for curr_value in b64_strs:
                    if (len(curr_value) > 100):
                        got_ioc = True
                        num_b64_iocs += 1
                        log.info("Found possible intermediate IOC (base64): '" + curr_value + "'")

        if (not got_ioc):
            return

        iocs_to_delete = set()
        got_ioc = True
        for old_value in intermediate_iocs:
            if (value.startswith(old_value)):
                iocs_to_delete.add(old_value)
            if ((old_value.startswith(value)) and (len(old_value) > len(value))):
                got_ioc = False

        if (got_ioc):
            intermediate_iocs.add(value)
            
        for old_ioc in iocs_to_delete:
            intermediate_iocs.remove(old_ioc)

    def _set_excel_formula(self, name, value):
        """
        Handle setting an Excel cell to a formula.

        Sheets('dd').Cells('d, dd').FormulaLocal = ...
        """

        if ((name is None) or
            (value is None) or
            (value == "__ALREADY_SET__")):
            return False
        
        import expressions
        import vba_object
        if (not isinstance(name, expressions.MemberAccessExpression)):
            return False
        tmp_rhs = str(name.rhs)
        if (isinstance(name.rhs, list)):
            tmp_rhs = str(name.rhs[-1])
        if ((tmp_rhs.lower() != "formulalocal") and
            (tmp_rhs.lower() != "value") and
            (tmp_rhs.lower() != "name")):
            return False
        typ = tmp_rhs
        if (typ.lower() == "formulalocal"):
            typ = "Formula"
        
        row = "??"
        col = "??"

        if (isinstance(name.rhs, list) and
            (isinstance(name.rhs[0], expressions.Function_Call)) and
            (str(name.rhs[0]).startswith("Cells("))):

            cell_call = name.rhs[0]
            row = str(vba_object.eval_arg(cell_call.params[0], self))
            col = str(vba_object.eval_arg(cell_call.params[1], self))

        sheet = "??"
        if (isinstance(name.lhs, expressions.Function_Call) and
            (str(name.lhs).startswith("Sheets("))):

            sheet_call = name.lhs
            sheet = str(vba_object.eval_arg(sheet_call.params[0], self))
        
        r = "Sheet(" + sheet + ").Cell(" + row + ", " + col + ") = '" + str(value) + "'"
        self.report_action('Set Cell ' + typ, r, tmp_rhs, strip_null_bytes=True)
        return True

    def _handle_property_assignment(self, name, value):
        """
        If this is a property asignment, call the property handler.
        """

        import procedures
        
        if (not self.contains(name)):
            return False

        handler = self.get(name)
        if (not isinstance(handler, procedures.PropertyLet)):
            return False

        handler.eval(self, params=[value])

        return True
    
    def set(self,
            name,
            value,
            var_type=None,
            do_with_prefix=True,
            force_local=False,
            force_global=False,
            no_conversion=False,
            case_insensitive=True,
            no_overwrite=False):

        if (self._set_excel_formula(name, value)):
            return

        if (self._handle_property_assignment(name, value)):
            return

        import expressions
        if (isinstance(name, expressions.SimpleNameExpression)):
            name = str(name)
        
        orig_name = name
        if (not isinstance(name, basestring)):
            log.warning("context.set() " + str(name) + " is improper type. " + str(type(name)))
            name = str(name)

        if (value is None):
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("context.set() " + str(name) + " failed. Value is None.")
            return

        if (".." in name):
            self.set(name.replace("..", "."), value, var_type, do_with_prefix, force_local, force_global, no_conversion=no_conversion)

        if (no_overwrite and self.contains(name)):
            return
            
        self.save_intermediate_iocs(value)
        
        if (case_insensitive):
            tmp_name = name.lower()
            self.set(tmp_name, value, var_type, do_with_prefix, force_local, force_global, no_conversion=no_conversion, case_insensitive=False)

        name_str = str(name)
        if (("(" in name_str) and (")" in name_str)):

            name_str = name_str[:name_str.index("(")].strip()
            if (name_str in self.globals.keys()):
                force_global = True

        if ((not self.in_procedure) and (not force_global) and (not force_local)):
            self.set(name, value, force_global=True, do_with_prefix=do_with_prefix)
            return
                

        if (force_global):
            try:
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("Set global var " + str(name) + " = " + str(value))
            except:
                pass
            self.globals[name] = value

        elif ((name in self.locals) or force_local):
            try:
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("Set local var " + str(name) + " = " + str(value))
            except:
                pass
            self.locals[name] = value

        elif name in self.globals and not is_procedure(self.globals[name]):
            self.globals[name] = value
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("Set global var " + name + " = " + str(value))
            if ("." in name):
                text_name = name + ".text"
                self.globals[text_name] = value
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("Set global var " + text_name + " = " + str(value))

        else:
            if (not self.global_scope):
                try:
                    if (log.getEffectiveLevel() == logging.DEBUG):
                        log.debug("Set local var " + str(name) + " = " + str(value))
                except:
                    pass
                self.locals[name] = value
            else:
                self.globals[name] = value
                try:
                    if (log.getEffectiveLevel() == logging.DEBUG):
                        log.debug("Set global var " + name + " = " + str(value))
                except:
                    pass
                if ("." in name):
                    text_name = name + ".text"
                    self.globals[text_name] = value
                    if (log.getEffectiveLevel() == logging.DEBUG):
                        log.debug("Set global var " + text_name + " = " + str(value))
                    text_name = name[name.rindex("."):]
                    self.globals[text_name] = value
                    if (log.getEffectiveLevel() == logging.DEBUG):
                        log.debug("Set global var " + text_name + " = " + str(value))
                
        if (var_type is not None):
            self.types[name] = var_type

        if ((do_with_prefix) and (len(self.with_prefix) > 0)):
            tmp_name = str(self.with_prefix) + "." + str(name)
            self.set(tmp_name, value, var_type=var_type, do_with_prefix=False, no_conversion=no_conversion)

        if (no_conversion):
            return
            
        if (name.endswith(".text")):

            do_b64 = False
            node_type = name.replace(".text", ".datatype")
            try:

                val = str(self.get(node_type)).strip()
                if (val.lower() == "bin.base64"):
                    do_b64 = True

            except KeyError:
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("Did not find type var " + node_type)

            try:

                import expressions
                import vba_object
                node_type = orig_name
                if (isinstance(orig_name, expressions.MemberAccessExpression)):
                    node_type = orig_name.lhs
                else:
                    node_type = str(node_type).lower().replace(".text", "")
                val = vba_object.eval_arg(node_type, self)
                if (val == "Microsoft.XMLDOM"):
                    do_b64 = True

            except KeyError:
                pass
            
            if (do_b64):

                conv_val = utils.b64_decode(value)
                if (conv_val is not None):
                    val_name = name
                    self.set(val_name, conv_val, no_conversion=True, do_with_prefix=do_with_prefix)
                    val_name = name.replace(".text", ".nodetypedvalue")
                    self.set(val_name, conv_val, no_conversion=True, do_with_prefix=do_with_prefix)

        if (name.lower().endswith(".nodetypedvalue")):

            node_type = name[:name.rindex(".")] + ".datatype"
            try:

                val = str(self.get(node_type)).strip()
                if (val.lower() == "bin.hex"):

                    try:

                        conv_val = codecs.decode(str(value).strip(), "hex")
                        self.set(name, conv_val, no_conversion=True, do_with_prefix=do_with_prefix)
                    except Exception as e:
                        log.warning("hex conversion of '" + str(value) + "' FROM hex failed. Converting TO hex. " + str(e))
                        conv_val = to_hex(str(value).strip())
                        self.set(name, conv_val, no_conversion=True, do_with_prefix=do_with_prefix)
                        
            except KeyError:
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("Did not find type var " + node_type)

        if (name.endswith(".datatype")):

            node_value_name = name.replace(".datatype", ".nodetypedvalue")
            try:

                node_value = self.get(node_value_name)
                if (value.lower() == "bin.hex"):

                    try:

                        conv_val = codecs.decode(str(node_value).strip(), "hex")
                        self.set(node_value_name, conv_val, no_conversion=True, do_with_prefix=do_with_prefix)
                    except Exception as e:
                        log.warning("hex conversion of '" + str(node_value) + "' FROM hex failed. Converting TO hex. " + str(e))
                        conv_val = to_hex(str(node_value).strip())
                        self.set(node_value_name, conv_val, no_conversion=True, do_with_prefix=do_with_prefix)

                if (value.lower() == "bin.base64"):

                    
                    conv_val = utils.b64_decode(node_value)
                    if (conv_val is not None):
                        self.set(node_value_name, conv_val, no_conversion=True, do_with_prefix=do_with_prefix)
                        
            except KeyError:
                pass
            
    def _strip_null_bytes(self, item):
        r = item
        if (isinstance(item, str)):
            r = item.replace("\x00", "")
        if (isinstance(item, list)):
            r = []
            for s in item:
                if (isinstance(s, str)):
                    r.append(s.replace("\x00", ""))
                else:
                    r.append(s)
        return r
                    
    def report_action(self, action, params=None, description=None, strip_null_bytes=False):

        if (strip_null_bytes):

            action = utils.strip_nonvb_chars(action)
            new_params = utils.strip_nonvb_chars(params)
            if (isinstance(params, list)):
                new_params = []
                for p in params:
                    tmp_p = utils.strip_nonvb_chars(p)
                    if (len(re.findall(r"NULL", str(tmp_p))) > 20):
                        tmp_p = str(tmp_p).replace("NULL", "")
                    new_params.append(tmp_p)
            params = new_params
            description = utils.strip_nonvb_chars(description)

            if (len(re.findall(r"NULL", action)) > 20):
                action = action.replace("NULL", "")
            
        self.got_actions = True
        self.engine.report_action(action, params, description)

