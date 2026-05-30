


import re
from curses_ascii import isascii
from curses_ascii import isprint
import base64

import logging

try:
    from core.logger import log
except ImportError:
    from logger import log
try:
    from core.logger import CappedFileHandler
except ImportError:
    from logger import CappedFileHandler
from logging import LogRecord
from logging import FileHandler
import excel

def safe_str_convert(s):
    """
    Convert a string to ASCII without throwing a unicode decode error.
    """

    if (isinstance(s, dict) and ("value" in s)):
        s = s["value"]

    try:
        return str(s)
    except UnicodeDecodeError:
        return filter(isprint, s)
    except UnicodeEncodeError:
        return filter(isprint, s)

class Infix:
    """
    Used to define our own infix operators.
    """
    def __init__(self, function):
        self.function = function
    def __ror__(self, other):
        return Infix(lambda x, self=self, other=other: self.function(other, x))
    def __or__(self, other):
        return self.function(other)
    def __rlshift__(self, other):
        return Infix(lambda x, self=self, other=other: self.function(other, x))
    def __rshift__(self, other):
        return self.function(other)
    def __call__(self, value1, value2):
        return self.function(value1, value2)

def safe_plus(x,y):
    """
    Handle "x + y" where x and y could be some combination of ints and strs.
    """

    if excel.is_cell_dict(x):
        x = x["value"]
    if excel.is_cell_dict(y):
        y = y["value"]
    
    if (x == "NULL"):
        x = 0
    if (y == "NULL"):
        y = 0

    if (isinstance(x, str)):
        y = str_convert(y)
    if (isinstance(x, int)):
        y = int_convert(y)

    if ((isinstance(x, int) or isinstance(x, float)) and
        (isinstance(y, int) or isinstance(y, float))):
        return x + y
        
    if (isinstance(y, str)):

        if (x == 0):
            x = ""

        return str(x) + y

    if (isinstance(x, str)):

        if (y == 0):
            y = ""

        return x + str(y)

    return str(x) + str(y)

plus=Infix(lambda x,y: safe_plus(x, y))

def safe_equals(x,y):
    """
    Handle "x = y" where x and y could be some combination of ints and strs.
    """

    if (x == "NULL"):
        x = 0
    if (y == "NULL"):
        y = 0
    
    if (type(x) == type(y)):
        return x == y

    if ((isinstance(x, bool) and (isinstance(y, int))) or
        (isinstance(y, bool) and (isinstance(x, int)))):
        return x == y
        
    return str(x) == str(y)

eq=Infix(lambda x,y: safe_equals(x, y))
neq=Infix(lambda x,y: (not safe_equals(x, y)))

def safe_print(text):
    """
    Sometimes printing large strings when running in a Docker container triggers exceptions.
    This function just wraps a print in a try/except block to not crash SimulationVBA when this happens.
    """
    text = safe_str_convert(text)
    try:
        print(text)
    except Exception as e:
        msg = "ERROR: Printing text failed (len text = " + str(len(text)) + ". " + str(e)
        if (len(msg) > 100):
            msg = msg[:100]
        try:
            print(msg)
        except:
            pass

    for handler in log.handlers:
        if type(handler) is FileHandler or type(handler) is CappedFileHandler:
            handler.setFormatter(logging.Formatter("%(message)s"))
            handler.emit(LogRecord(log.name, logging.INFO, "", None, text, None, None, "safe_print"))
            handler.setFormatter(logging.Formatter("%(levelname)-8s %(message)s"))

def fix_python_overlap(var_name):
    builtins = set(["str", "list", "bytes", "pass"])
    if (var_name.lower() in builtins):
        var_name = "MAKE_UNIQUE_" + var_name
    var_name = var_name.replace("$", "__DOLLAR__")
    if ((not var_name.endswith(".Pattern")) and
        (not var_name.endswith(".Global"))):
        var_name = var_name.replace(".", "")
    return var_name

def b64_decode(value):
    """
    Base64 decode a string.
    """

    try:
        tmp_str = ""
        try:
            tmp_str = filter(isascii, str(value).strip())
        except UnicodeDecodeError:
            return None
        tmp_str = tmp_str.replace(" ", "").replace("\x00", "")
        b64_pat = r"^[A-Za-z0-9+/=]+$"
        if (re.match(b64_pat, tmp_str) is not None):
            
            missing_padding = len(tmp_str) % 4
            if missing_padding:
                tmp_str += b'='* (4 - missing_padding)
        
            conv_val = base64.b64decode(tmp_str)
            return conv_val
    
    except Exception as e:
        pass

    return None

class vb_RegExp(object):
    """
    Class to simulate a VBS RegEx object in python.
    """

    def __init__(self):
        self.Pattern = None
        self.Global = False

    def __repr__(self):
        return "<RegExp Object: Pattern = '" + str(self.Pattern) + "', Global = " + str(self.Global) + ">"
        
    def _get_python_pattern(self):
        pat = self.Pattern
        if (pat is None):
            return None
        if (pat.strip() != "."):
            pat1 = pat.replace("$", "\\$").replace("-", "\\-")
            fix_dash_pat = r"(\[.\w+)\\\-(\w+\])"
            pat1 = re.sub(fix_dash_pat, r"\1-\2", pat1)
            fix_dash_pat1 = r"\((\w+)\\\-(\w+)\)"
            pat1 = re.sub(fix_dash_pat1, r"[\1-\2]", pat1)
            pat = pat1
        return pat
        
    def Test(self, string):
        pat = self._get_python_pattern()
        if (pat is None):
            return False
        return (re.match(pat, string) is not None)

    def Replace(self, string, rep):
        pat = self._get_python_pattern()
        if (pat is None):
            return string
        rep = re.sub(r"\$(\d)", r"\\\1", rep)
        r = string
        try:
            r = re.sub(pat, rep, string)
        except Exception as e:
            pass
        return r

def get_num_bytes(i):
    """
    Get the minimum number of bytes needed to represent a given
    int value.
    """
    
    if ((i & 0x00000000FF) == i):
        return 1
    if ((i & 0x000000FFFF) == i):
        return 2
    if ((i & 0x00FFFFFFFF) == i):
        return 4
    return 8

def int_convert(arg, leave_alone=False):
    """
    Convert a VBA expression to an int, handling VBA NULL.
    """

    if (isinstance(arg, int)):
        return arg
    
    if (arg == "NULL"):
        return 0

    if (arg == ""):
        return "NULL"
    
    if (arg == "**MATCH ANY**"):
        return arg

    if (isinstance(arg, float)):
        arg = int(round(arg))

    if (isinstance(arg, str) and (arg.strip().lower().startswith("&h"))):
        hex_str = "0x" + arg.strip()[2:]
        try:
            return int(hex_str, 16)
        except:
            log.error("Cannot convert hex '" + str(arg) + "' to int. Defaulting to 0. " + str(e))
            return 0
            
    arg_str = str(arg)
    if ("." in arg_str):
        arg_str = arg_str[:arg_str.index(".")]
    try:
        return int(arg_str)
    except Exception as e:
        if (not leave_alone):
            log.error("Cannot convert '" + str(arg_str) + "' to int. Defaulting to 0. " + str(e))
            return 0
        log.error("Cannot convert '" + str(arg_str) + "' to int. Leaving unchanged. " + str(e))
        return arg_str

def str_convert(arg):
    """
    Convert a VBA expression to an str, handling VBA NULL.
    """
    if (arg == "NULL"):
        return ''
    if (excel.is_cell_dict(arg)):
        arg = arg["value"]
    try:
        return str(arg)
    except Exception as e:
        if (isinstance(arg, unicode)):
            return ''.join(filter(lambda x:x in string.printable, arg))
        log.error("Cannot convert given argument to str. Defaulting to ''. " + str(e))
        return ''

def strip_nonvb_chars(s):
    """
    Strip invalid VB characters from a string.
    """

    if (isinstance(s, unicode)):
        s = s.encode('ascii','replace')
    
    if (not isinstance(s, str)):
        return s

    if (re.search(r"[^\x09-\x7e]", s) is None):
        return s
    
    r = re.sub(r"[^\x09-\x7e]", "", s)
    
    if (r.count("NULL") > 10):
        r = r.replace("NULL", "")
    return r
    
