"""
SimulationVBA - Strip useles lines from Visual Basic code.

Author: Philippe Lagadec - http://www.decalage.info
License: BSD, see source code or documentation

Project Repository:
https://github.com/decalage2/ViperMonkey
"""

#=== LICENSE ==================================================================






import logging
import sys
import re
try:
    import rure as re2
except:
    import re as re2
from logger import log
import vba_context
from random import randint

debug_strip = False

def is_useless_dim(line):
    """
    See if we can skip this Dim statement and still successfully emulate.
    We only use Byte type information when emulating.
    """

    line = line.strip()
    if (not line.startswith("Dim ")):
        return False
    r = (("Byte" not in line) and
         ("Long" not in line) and
         ("Integer" not in line) and
         (":" not in line) and
         ("=" not in line) and
         (not line.strip().endswith("_")))

    line = line.lower()
    for builtin in vba_context.VBA_LIBRARY.keys():
        if (builtin in line):
            r = False

    return r

aggressive_strip = False
def is_interesting_call(line, external_funcs, local_funcs):

    log_funcs = ["CreateProcessA", "CreateProcessW", ".run", "CreateObject",
                 "Open", "CreateMutex", "CreateRemoteThread", "InternetOpen",
                 ".Open", "GetObject", "Create", ".Create", "Environ",
                 "CreateTextFile", ".CreateTextFile", "Eval", ".Eval", "Run",
                 "SetExpandedStringValue", "WinExec", "URLDownloadToFile", "Print",
                 "Split", "Exec"]
    if (not aggressive_strip):
        log_funcs.extend(local_funcs)
    for func in log_funcs:
        if (func in line):
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("Line '" + line + "' contains interesting call (1) ('" + func + "').")
            return True

    for ext_func_decl in external_funcs:
        if (("Function" in ext_func_decl) and ("Lib" in ext_func_decl)):
            start = ext_func_decl.index("Function") + len("Function")
            end = ext_func_decl.index("Lib")
            ext_func = ext_func_decl[start:end].strip()
            if (ext_func in line):
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("Line '" + line + "' contains interesting call (2).")
                return True
        
    return False

def is_useless_call(line):
    """
    See if the given line contains a useless do-nothing function call.
    """

    useless_funcs = set(["Cos", "Log", "Cos", "Exp", "Sin", "Tan", "DoEvents"])

    if ("=" in line):
        return False

    line = line.replace(" ", "")
    called_func = line
    if ("(" in line):
        called_func = line[:line.index("(")]
    called_func = called_func.strip()
    for func in useless_funcs:
        if (called_func == func):
            return True
    return False

def collapse_macro_if_blocks(vba_code):
    """
    When emulating we only pick a single block from a #if statement. Speed up parsing
    by picking the largest block and strip out the rest.
    """

    if (log.getEffectiveLevel() == logging.DEBUG):
        log.debug("Collapsing macro blocks...")
    curr_blocks = None
    curr_block = None
    r = ""
    for line in vba_code.split("\n"):

        strip_line = line.strip()
        if (curr_blocks is None):

            if (strip_line.startswith("#If")):

                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("Start block " + strip_line)
                curr_blocks = []
                curr_block = []
                r += "' STRIPPED LINE\n"
                continue

            r += line + "\n"
            continue


        if (strip_line.startswith("#Else")):

            curr_blocks.append(curr_block)
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("Else if " + strip_line)
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("Save block " + str(curr_block))

            curr_block = []
            r += "' STRIPPED LINE\n"
            continue

        if (strip_line.startswith("#End")):

            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("End if " + strip_line)

            curr_blocks.append(curr_block)
            
            biggest_block = []
            for block in curr_blocks:
                if (len(block) > len(biggest_block)):
                    biggest_block = block
            for block_line in biggest_block:
                r += block_line + "\n"
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("Pick block " + str(biggest_block))
                
            curr_blocks = None
            curr_block = None
            continue

        curr_block.append(line)

    if (r.strip() != vba_code.strip()):
        r = collapse_macro_if_blocks(r)
        
    return r
    
def hide_weird_calls(vba_code):
    """
    Hide weird calls like 'foo.bar (1,2), cat, dog'. These are hard to parse.
    """

    uni_vba_code = None
    try:
        uni_vba_code = vba_code.decode("utf-8")
    except UnicodeDecodeError:
        return vba_code
    pat = u"\r?\n\s*\w+(?:\.\w+)*\s*\(.{1,30}\)"
    print re2.findall(pat, uni_vba_code)
    if (re2.search(pat, uni_vba_code) is None):
        return vba_code
    return vba_code

def fix_bad_puts(vba_code):
    """
    Change file Put statements like 'Put #foo(1,2,3) ...' to 'Put foo(1,2,3)'.
    """

    if ("Put #" not in vba_code):
        return vba_code

    vba_code = re.sub(r"([^A-Za-z])Put +#([A-Za-z_])", r"\1Put \2", vba_code)

    if ("Close #" not in vba_code):
        return vba_code

    vba_code = re.sub(r"([^A-Za-z])Close +#([A-Za-z_])", r"\1Close \2", vba_code)
    return vba_code
    
def fix_unbalanced_quotes(vba_code):
    """
    Fix lines with missing double quotes.
    """

    uni_vba_code = None
    try:
        uni_vba_code = vba_code.decode("utf-8")
    except UnicodeDecodeError:
        return vba_code

    if debug_strip:
        print "UNBALANCED_QUOTES: 1"
        print vba_code
    if (re2.search(u"\r?\n\s*(?:Set)?\s*(\w+)\s+=\s+\"\r?\n", uni_vba_code) is not None):
        vba_code = re.sub(r"\r?\n\s*(?:Set)?\s*(\w+)\s+=\s+\"\r?\n", r'\n\1 = ""\n', vba_code)
        if debug_strip:
            print "UNBALANCED_QUOTES: 2"
            print vba_code
    if (re2.search(u"(\w+\s+=\s+\")(:[^\"]+)\r?\n", uni_vba_code) is not None):
        vba_code = re.sub(r"(\w+\s+=\s+\")(:[^\"]+)\r?\n", r'\1"\2\n', vba_code)
        if debug_strip:
            print "UNBALANCED_QUOTES: 2"
            print vba_code
    if (re2.search(u"^\"[^=]*([=>])\s*\"\s+[Tt][Hh][Ee][Nn]", uni_vba_code) is not None):
        vba_code = re.sub(r"^\"[^=]*([=>])\s*\"\s+[Tt][Hh][Ee][Nn]", r'\1 "" Then', vba_code)
        if debug_strip:
            print "UNBALANCED_QUOTES: 3"
            print vba_code
        
    vba_code += "\n"
    vba_code = re.sub(r"'('[^'^\"]+\n)", r"\1", vba_code, re.DOTALL)
    if debug_strip:
        print "UNBALANCED_QUOTES: 4"
        print vba_code
    
    vba_code = re.sub(r"(\n[^'^\n]+)'[^'^\"^\n]+'[^'^\"^\n]+\n", r"\1\n", vba_code, re.DOTALL)
    if debug_strip:
        print "UNBALANCED_QUOTES: 5"
        print vba_code
    
    vba_code = re.sub(r"\n\s*([Ee][Xx][Ee][Cc][Uu][Tt][Ee])\"", r'\nExecute "', vba_code)
    if debug_strip:
        print "UNBALANCED_QUOTES: 6"
        print vba_code
    
    r = ""
    lines = vba_code.split("\n")
    pos = -1
    synthetic_line = False
    while (pos < (len(lines) - 1)):

        pos += 1
        if (not synthetic_line):
            line = lines[pos]
        synthetic_line = False
        next_line = ""
        if ((pos + 1) < len(lines)):
            next_line = lines[pos + 1]
            while ((len(next_line.strip()) == 0) and
                   ((pos + 2) < len(lines))):
                r += next_line
                pos += 1
                next_line = lines[pos + 1]
        if ('"' not in line):
            r += line + "\n"
            continue

        if ("'" in line):
            quote_index = None
            in_str = False
            tmp_pos = -1
            for c in line:
                tmp_pos += 1
                if (c == '"'):
                    in_str = not in_str
                    continue
                if ((c == "'") and (not in_str)):
                    quote_index = tmp_pos
                    break
            if (quote_index is not None):
                line = line[:quote_index]
        num_quotes = line.count('"')
        if ((num_quotes % 2) != 0):
            
            if (line.strip().endswith('"') and next_line.strip().startswith('"')):
                tmp_line = line + "\\n" + next_line
                line = tmp_line
                synthetic_line = True
                continue
            
            first_quote = line.index('"')
            line = line[:first_quote] + '"' + line[first_quote:]
        r += line + "\n"

    if debug_strip:
        print "UNBALANCED_QUOTES: 7"
        print r
    return r

MULT_ASSIGN_RE = r"((?:\w+\s*=\s*){3,})(.+)"
def fix_multiple_assignments(line):

    if ("'" in line):

        in_str = False
        quote_pos = None
        curr_pos = -1
        for c in line:
            curr_pos += 1
            if (c == '"'):
                in_str = not in_str
            if ((c == "'") and (not in_str)):
                quote_pos = curr_pos
                break
        if (quote_pos is not None):
            line = line[:quote_pos]
    
    items = re.findall(MULT_ASSIGN_RE, line)
    if (len(items) == 0):
        return line

    in_str = False
    new_line = ""
    for c in line:

        if (c == '"'):
            in_str = not in_str

        if ((c == '=') and (in_str)):
            new_line += 'IN_STR_EQUAL'
        else:
            new_line += c
            
    items = re.findall(MULT_ASSIGN_RE, new_line)
    if (len(items) == 0):
        return line
    items = items[0]
    assigns = items[0].replace(" ", "").split("=")
    val = items[1]

    r = ""
    for var in assigns:
        var = var.strip()
        if (len(var) == 0):
            continue
        r += var + " = " + val + "\n"

    r.replace('IN_STR_EQUAL', '=')
    return r

def fix_skipped_1st_arg1(vba_code):
    """
    Replace calls like foo(, 1, ...) with foo(SKIPPED_ARG, 1, ...).
    """

    uni_vba_code = None
    try:
        uni_vba_code = vba_code.decode("utf-8")
    except UnicodeDecodeError:
        return vba_code
    if (re2.search(unicode(r".*[0-9a-zA-Z_\.]+\(\s*,.*"), uni_vba_code, re.DOTALL) is None):
        return vba_code
    

    tmp_code = vba_code.replace('""', '__ESCAPED_QUOTES__')
    
    strings = {}
    in_str = False
    curr_str = None
    for c in tmp_code:

        if (c == '"'):

            if (not in_str):
                curr_str = ""
                in_str = True

            else:

                str_name = "A_STRING_LITERAL_" + str(randint(0, 100000000))
                while (str_name in strings):
                    str_name = "A_STRING_LITERAL_" + str(randint(0, 100000000))
                curr_str += c
                strings[str_name] = curr_str
                in_str = False
                curr_str = None

        if (in_str):
            curr_str += c

    for str_name in strings.keys():
        tmp_code = tmp_code.replace(strings[str_name], str_name)
        
    vba_code = re.sub(r"([0-9a-zA-Z_\.]+)\(\s*,", r"\1(SKIPPED_ARG,", tmp_code)

    for str_name in strings.keys():
        vba_code = vba_code.replace(str_name, strings[str_name])

    vba_code = vba_code.replace('__ESCAPED_QUOTES__', '""')
        
    return vba_code

def fix_skipped_1st_arg2(vba_code):
    """
    Replace calls like \nfoo, 1, ... with \nfoo SKIPPED_ARG, 1, ... .
    """

    if (re.match(r".*\n\s*([0-9a-zA-Z_\.\(\)]+)\s*,.*", vba_code, re.DOTALL) is None):
        return vba_code
    

    strings = {}
    in_str = False
    curr_str = None
    for c in vba_code:

        if (c == '"'):

            if (not in_str):
                curr_str = ""
                in_str = True

            else:

                str_name = "A_STRING_LITERAL_" + str(randint(0, 100000000))
                while (str_name in strings):
                    str_name = "A_STRING_LITERAL_" + str(randint(0, 100000000))
                curr_str += c
                strings[str_name] = curr_str
                in_str = False
                curr_str = None

        if (in_str):
            curr_str += c

    tmp_code = vba_code
    for str_name in strings.keys():
        tmp_code = tmp_code.replace(strings[str_name], str_name)
        
    in_paren = False
    paren_count = 0
    parens = {}
    curr_paren = None
    for c in tmp_code:

        if (c == '('):

            paren_count += 1
            if (paren_count > 0):
                curr_paren = ""
                in_paren = True

        if (c == ')'): 

            paren_count -= 1
            if (paren_count <= 0):
                
                paren_count = 0
                in_paren = False
                if (curr_paren is not None):
                    paren_name = "A_PAREN_EXPR_" + str(randint(0, 100000000))
                    while (paren_name in parens):
                        str_name = "A_PAREN_EXPR_" + str(randint(0, 100000000))
                    curr_paren += c
                    parens[paren_name] = curr_paren
                    curr_paren = None

        if (in_paren):
            curr_paren += c

    for paren_name in parens.keys():
        tmp_code = tmp_code.replace(parens[paren_name], paren_name)
        
    vba_code = re.sub(r"\n\s*([0-9a-zA-Z_\.]+)\s*,", r"\n\1 SKIPPED_ARG,", tmp_code)

    for paren_name in parens.keys():
        vba_code = vba_code.replace(paren_name, parens[paren_name])
    for str_name in strings.keys():
        vba_code = vba_code.replace(str_name, strings[str_name])
        
    return vba_code

def fix_bad_next_statements(vba_code):
    """
    Change things like "Next x,y" to "Next x\nNext y"
    """

    uni_vba_code = None
    try:
        uni_vba_code = vba_code.decode("utf-8")
    except UnicodeDecodeError:
        return vba_code
    
    pat = "Next +(?:\w+ *, *)+\w+ *\n"
    r = vba_code
    if (re2.search(unicode(pat), uni_vba_code) is not None):
        index_pat = "(?:(\w+) *, *)+(\w+)"
        for bad_next in re.findall(pat, vba_code):
            new_nexts = ""
            for index in re.findall(index_pat, bad_next)[0]:
                new_nexts += "Next " + index + "\n"
            r = r.replace(bad_next, new_nexts)
    return r

def fix_items_ref(vba_code):
    """
    Change Scripting.Dictionary.Items() references to 
    Scripting.Dictionary.Items.
    """

    if (".Items()(" not in vba_code):
        return vba_code
    r = vba_code.replace(".Items()(", ".Items(")
    return r

def fix_stupid_string_concats(vba_code):
    """
    Change garbage string concatentations like 's1 & s2 & + "foo"' to
    's1 & s2 & "foo"'.
    """

    uni_vba_code = None
    try:
        uni_vba_code = vba_code.decode("utf-8")
    except UnicodeDecodeError:
        return vba_code
    if (re2.search(unicode(r"&\s+\+"), uni_vba_code) is None):
        return vba_code

    r = re.sub(r"&\s+\+", "& ", vba_code)
    return r

def fix_bad_exponents(vba_code):
    """
    Change things like '2^2' to '2 ^ 2'. 
    """

    uni_vba_code = None
    try:
        uni_vba_code = vba_code.decode("utf-8")
    except UnicodeDecodeError:
        return vba_code
    
    pat = '([\w\(\)])\^([\w\(\)])'
    r = ""
    if (re2.search(unicode(pat), uni_vba_code) is not None):

        for line in vba_code.split("\n"):

            if ('"' in line):
                r += line + "\n"
                continue

            uni_line = None
            try:
                uni_line = line.decode("utf-8")
            except UnicodeDecodeError:
                r += line + "\n"
                continue
            if (re2.search(unicode(pat), uni_line) is None):
                r += line + "\n"
                continue

            new_line = re.sub(pat, r"\1 ^ \2", line)
            r += new_line + "\n"

    else:
        return vba_code
            
    return r

def fix_bad_var_names(vba_code):
    """
    Change things like a& = b& + 1 to a = b + 1.
    """

    return vba_code
    
    uni_vba_code = None
    try:
        uni_vba_code = vba_code.decode("utf-8")
    except UnicodeDecodeError:
        return vba_code
    
    pat = "(\w)&\s*((?:[\+\-/\*=n,\)\n&]|[Mm]od|[Aa]nd|[Oo]r|[Xx]or|[Ee]qv))"
    if (re2.search(unicode(pat), uni_vba_code) is not None):
        vba_code = re.sub(pat, r"\1 \2", vba_code) + "\n"
    return vba_code

def fix_unhandled_named_params(vba_code):
    """
    Currently things like 'foo(a:=1, b:=2)' are not handled.
    Comment them out.
    """

    uni_vba_code = None
    try:
        uni_vba_code = vba_code.decode("utf-8")
    except UnicodeDecodeError:
        return vba_code
    
    pat = "\n([^\n]*\w+\([^\n]*\w+:=)"
    if (re2.search(unicode(pat), uni_vba_code) is not None):

        line_pat = r"\n[^\n]*:=[^\n]*\n"

        lines = re.findall(line_pat, "\n" + vba_code + "\n")
        for line in lines:

            got_marker = False
            in_str = False
            last_char = ""
            for c in line:

                if (c == '"'):
                    in_str = not in_str
                if (in_str):
                    last_char = c
                    continue

                if ((last_char + c) == ':='):
                    got_marker = True
                    break
                last_char = c

            if (got_marker):

                log.warning("Named parameters are not currently handled. Commenting them out...")
                line = line[:-1]
                new_line = None

                if (line.strip().lower().startswith("with ")):

                    new_line = "\nWith UNHANDLED_NAMED_PARAMS_REPLACEMENT\n"

                elif (line.strip().lower().startswith("if ")):

                    new_line = "\n' UNHANDLED_NAMED_PARAMS_REPLACEMENT in If.\n'" + line.strip() + "\nIf 1=1 Then\n"                
                
                else:
                    new_line = re.sub(pat, r"\n' UNHANDLED NAMED PARAMS \1", line) + "\n"

                vba_code = vba_code.replace(line, new_line)
                
    return vba_code

def fix_unhandled_array_assigns(vba_code):
    """
    Currently things like 'foo(1, 2, 3) = 1' are not handled.
    Comment them out.
    """

    uni_vba_code = None
    try:
        uni_vba_code = vba_code.decode("utf-8")
    except UnicodeDecodeError:
        return vba_code
    
    pat = "\n(\s*\w+\((?:\w+\s*,\s*){2,}\w+\)\s*=)"
    if (re2.search(unicode(pat), uni_vba_code) is not None):
        vba_code = re.sub(pat, r"\n' UNHANDLED ARRAY ASSIGNMENT \1", vba_code) + "\n"
        fix_pat = r"' UNHANDLED ARRAY ASSIGNMENT\s+Mid\("
        vba_code = re.sub(fix_pat, r"Mid(", vba_code)
    return vba_code

def fix_unhandled_event_statements(vba_code):
    """
    Currently things like 'Event ...' are not handled.
    Comment them out.
    """

    uni_vba_code = None
    try:
        uni_vba_code = u"\n" + vba_code.decode("utf-8") + u"\n"
    except UnicodeDecodeError:
        return vba_code
    
    pat = "\n( *(?:[Pp]ublic) *Event[^\n]{10,})"
    if (re2.search(unicode(pat), uni_vba_code) is not None):
        vba_code = "\n" + vba_code + "\n"
        vba_code = re.sub(pat, r"\n' UNHANDLED EVENT STATEMENT \1", vba_code) + "\n"
    return vba_code

def fix_unhandled_raiseevent_statements(vba_code):
    """
    Currently things like 'RaiseEvent ...' are not handled.
    Comment them out.
    """

    uni_vba_code = None
    try:
        uni_vba_code = u"\n" + vba_code.decode("utf-8") + u"\n"
    except UnicodeDecodeError:
        return vba_code
    
    pat = "\n( *RaiseEvent[^\n]{3,})"
    if (re2.search(unicode(pat), uni_vba_code) is not None):
        vba_code = "\n" + vba_code + "\n"
        vba_code = re.sub(pat, r"\n' UNHANDLED RAISEEVENT STATEMENT \1", vba_code) + "\n"
    return vba_code

def hide_string_content(s):
    """
    Replace contents of string literals with '____'.
    """
    if (not isinstance(s, str)):
        return s
    if ('"' not in s):
        return s
    r = ""
    start = 0    
    while ('"' in s[start:]):

        end = s[start:].index('"') + start + 1
        r += s[start:end]
        start = end

        end = len(s)
        if ('"' in s[start:]):
            end = s[start:].index('"') + start            
        r += "_" * (end - start - 1)
        start = end + 1
    r += s[start:]

    return r

def is_in_string(line, s):
    """
    Check to see if s appears in a quoted string in line.
    """

    if ('"' not in line):
        return False

    strs = []
    in_str = False
    curr_str = None
    for c in line:

        if ((c == '"') and (not in_str)):
            curr_str = ""
            in_str = True
            continue

        if ((c == '"') and in_str):
            strs.append(curr_str)
            in_str = False
            continue

        if in_str:
            curr_str += c

    for curr_str in strs:
        if (s in curr_str):
            return True

    return False

def hide_colons_in_ifs(vba_code):

    if ("If " not in vba_code):
        return vba_code
    r = ""
    pos = 0
    vba_code += "\n"
    while ("If " in vba_code[pos:]):

        if_index = vba_code[pos:].index("If ") + pos
        r += vba_code[pos:if_index]

        end_pos = vba_code[if_index:].index("\n") + if_index
        r += vba_code[if_index:end_pos+1].replace(":", "__COLON__")
        pos = end_pos
    r += vba_code[pos:]
    return r

def convert_colons_to_linefeeds(vba_code):
    """
    Convert things like 'a=1:b=2' to 'a=1\n:b=2'
    Also change things like 'a&"ff"' to 'a & "ff"'
    """

    if ((":" not in vba_code) and ("&" not in vba_code)):
        return vba_code

    vba_code = hide_colons_in_ifs(vba_code)
    
    marker_chars = [('"', '"', None), ('[', ']', None), ("'", '\n', None), ('#', '#', None), ("If ", "Then", "End If")]

    pos = 0
    r = ""
    while (pos < len(vba_code)):

        marker_pos1 = len(vba_code)
        use_start_marker = None
        use_end_marker = None
        found_marker = False
        for marker, end_marker, not_marker in marker_chars:

            if (marker in vba_code[pos:]):                    
                
                curr_marker_pos1 = vba_code[pos:].index(marker) + pos
                if (curr_marker_pos1 < marker_pos1):

                    if ((not_marker is not None) and (len(not_marker) < curr_marker_pos1)):
                        prev_text = vba_code[curr_marker_pos1 - (len(not_marker) - len(marker) + 1):curr_marker_pos1 + 2]
                        if (prev_text == not_marker):
                            continue

                    found_marker = True
                    marker_pos1 = curr_marker_pos1
                    use_end_marker = end_marker
                    use_start_marker = marker

        if (found_marker):

            change_chunk = vba_code[pos:marker_pos1+1]
            change_chunk = change_chunk.replace(":", "\n")
            change_chunk = re.sub(r"([\w_])&\"", r"\1 & " + "\"", change_chunk)
            change_chunk = re.sub(r"\"&\"", r"\" & " + "\"", change_chunk)
            change_chunk = re.sub(r"([\w_])&\(", r"\1 & " + "(", change_chunk)
            
            marker_pos2a = len(vba_code)
            marker_pos2b = len(vba_code)
            if (use_end_marker in vba_code[marker_pos1+1:]):
                marker_pos2a = vba_code[marker_pos1+1:].index(use_end_marker) + marker_pos1 + 2

            if ("\n" in vba_code[marker_pos1+1:]):
                marker_pos2b = vba_code[marker_pos1+1:].index("\n") + marker_pos1 + 2

            if (marker_pos2b < marker_pos2a):
                marker_pos2 = marker_pos2b
                use_end_marker = "\n"
            else:
                marker_pos2 = marker_pos2a
                
            leave_chunk = vba_code[marker_pos1+1:marker_pos2]
            r += change_chunk + leave_chunk
            pos = marker_pos2

        if (not found_marker):
            change_chunk = vba_code[pos:]
            change_chunk = change_chunk.replace(":", "\n")
            change_chunk = re.sub(r"([\w_])&\"", r"\1 & " + "\"", change_chunk)
            change_chunk = re.sub(r"\"&\"", r"\" & " + "\"", change_chunk)
            change_chunk = re.sub(r"([\w_])&\(", r"\1 & " + "(", change_chunk)
            
            r += change_chunk
            pos = len(vba_code)

    if (r == ""):
        r = vba_code

    r = r.replace("__COLON__", ":")

    tmp_r = ""
    for line in r.split("\n"):
        if ((not line.strip().startswith("Dim")) and
            (not line.strip().startswith("ReDim"))):
            tmp_r += line + "\n"
            continue
        line = line.replace(" & ", "&")
        tmp_r += line + "\n"
    r = tmp_r
    
    return r

def fix_varptr_calls(vba_code):
    """
    Change calls like VarPtr(foo(0)) to VarPtr(foo) so we can report on the
    whole byte array.
    """

    if ("VarPtr(" not in vba_code):
        return vba_code
    vba_code = re.sub(r"(VarPtr\(\w+)\(0\)\)", r'\1)', vba_code)
    return vba_code

def break_up_whiles(vba_code):
    """
    Break up while statements like 'While(a>b)c = c+1'.
    """

    vba_code_l = vba_code.lower()
    if ("while" not in vba_code_l):
        return vba_code

    pos = 0
    changes = {}
    vba_code += "\n"
    while ("while" in vba_code_l[pos:]):

        start = vba_code_l[pos:].index("while")
        end = vba_code_l[start + pos:].index("\n")
        curr_while = vba_code[start + pos:end + start + pos]

        line_start = vba_code_l[:start + pos].rindex("\n")
        curr_line = vba_code_l[line_start:end + start + pos]
        
        if ((not (curr_while.replace(" ", "").strip().lower()).startswith("while(")) or
            (is_in_string(curr_line, "while"))):
            pos = end + start + pos
            continue


        parens = 0
        curr_pos = -1
        for c in curr_while:
            curr_pos += 1
            if (c == "("):
                parens += 1
            if (c == ")"):
                parens -= 1
                if (parens == 0):
                    break

        if (len(curr_while[curr_pos+1:].strip()) > 0):

            new_while = curr_while[:curr_pos+1] + "\n" + curr_while[curr_pos+1:]
            changes[curr_while] = new_while

        pos = end + start + pos

    r = vba_code
    for curr_while in changes:
        r = r.replace(curr_while, changes[curr_while])

    return r
    
def fix_difficult_code(vba_code):
    """
    Replace characters whose ordinal value is > 128 with dNNN, where NNN
    is the ordinal value.

    Also change things like "a!b!c" to "a.b.c".

    Also break up multiple statements seperated with '::' or ':' onto different lines.

    Also change assignments like "a =+ 1 + 2" to "a = 1 + 2".
    """

    vba_code = vba_code.replace(chr(0x85), "")
    
    if debug_strip:
        print "HERE: 1"
        print vba_code
    vba_code = vba_code.replace("\n" + chr(0x85), "\n")
    vba_code = vba_code.replace("spli.tt.est", "splittest").replace("Mi.d", "Mid")
    vba_code = vba_code.replace("msgbox\"", "msgbox \"")
    vba_code = fix_unhandled_array_assigns(vba_code)
    if debug_strip:
        print "HERE: 2.1"
        print vba_code
    vba_code = fix_unhandled_event_statements(vba_code)
    if debug_strip:
        print "HERE: 2.2"
        print vba_code
    vba_code = fix_unhandled_raiseevent_statements(vba_code)
    if debug_strip:
        print "HERE: 2.3"
        print vba_code
    vba_code = fix_unhandled_named_params(vba_code)
    if debug_strip:
        print "HERE: 2.4"
        print vba_code
    vba_code = fix_bad_var_names(vba_code)
    if debug_strip:
        print "HERE: 2.5"
        print vba_code
    vba_code = fix_bad_exponents(vba_code)
    if debug_strip:
        print "HERE: 2.6"
        print vba_code
    vba_code = fix_stupid_string_concats(vba_code)
    if debug_strip:
        print "HERE: 2.6.1"
        print vba_code
    vba_code = fix_items_ref(vba_code)
    if debug_strip:
        print "HERE: 2.6.2"
        print vba_code
    vba_code = fix_bad_next_statements(vba_code)
    if debug_strip:
        print "HERE: 2.7"
        print vba_code
    vba_code = fix_varptr_calls(vba_code)

    uni_vba_code = u""
    try:
        uni_vba_code = vba_code.decode("utf-8")
    except UnicodeDecodeError:
        log.warning("Converting VB code to unicode failed.")
    if debug_strip:
        print "HERE: 3"
        print vba_code
    array_acc_pat = r'(\w+\([\d\w_\+\*/\-"]+\))(?:\([\d\w_\+\*/\-"]+\)){2,50}'
    if (re2.search(unicode(array_acc_pat), uni_vba_code) is not None):
        vba_code = re.sub(array_acc_pat, r"\1", vba_code)
    
    uni_vba_code = u""
    try:
        uni_vba_code = vba_code.decode("utf-8")
    except UnicodeDecodeError:
        log.warning("Converting VB code to unicode failed.")
    if debug_strip:
        print "HERE: 4"
        print vba_code
    namespace_pat = r"(\w+\.NameSpace\(.+\)\.CopyHere\(.+\)),\s*[^\n]+"
    if (re2.search(unicode(namespace_pat), uni_vba_code) is not None):
        vba_code = re.sub(namespace_pat, r"\1", vba_code)
    uni_vba_code = u""
    try:
        uni_vba_code = vba_code.decode("utf-8")
    except UnicodeDecodeError:
        log.warning("Converting VB code to unicode failed.")
    if debug_strip:
        print "HERE: 5"
        print vba_code
    namespace_pat = r"(CreateObject\(.+\).[Nn]ame[Ss]pace\(.+\)\.CopyHere\s+.+),\s*[^\n]+"
    if (re2.search(unicode(namespace_pat), uni_vba_code) is not None):
        vba_code = re.sub(namespace_pat, r"\1", vba_code)
    uni_vba_code = u""
    try:
        uni_vba_code = vba_code.decode("utf-8")
    except UnicodeDecodeError:
        log.warning("Converting VB code to unicode failed.")
    if debug_strip:
        print "HERE: 6"
        print vba_code
    namespace_pat = r"(\w+\.Run\(.+\)[^,\n]*),\s*[^\n\)]+"
    if (re2.search(unicode(namespace_pat), uni_vba_code) is not None):
        if debug_strip:
            print "HERE: 6.1"
        vba_code = re.sub(namespace_pat, r"\1", vba_code)
    
    uni_vba_code = u""
    try:
        uni_vba_code = vba_code.decode("utf-8")
    except UnicodeDecodeError:
        log.warning("Converting VB code to unicode failed.")
    if debug_strip:
        print "HERE: 8"
        print vba_code
    bad_else_pat = r"\n\s*Else\s*'.*\n"
    if (re2.search(unicode(bad_else_pat), uni_vba_code) is not None):
        bad_exps = re.findall(bad_else_pat, vba_code)
        for bad_exp in bad_exps:
            vba_code = vba_code.replace(bad_exp, "\nElse\n")
        
    if debug_strip:
        print "HERE: 9"
        print vba_code
    if (("!" not in vba_code) and
        (":" not in vba_code) and
        ("ElseIf" not in vba_code) and
        ("&" not in vba_code) and
        ("^" not in vba_code) and
        ("Rem " not in vba_code) and
        ("rem " not in vba_code) and
        ("REM " not in vba_code) and
        ("MultiByteToWideChar" not in vba_code) and
        (re.match(r".*[\x7f-\xff].*", vba_code, re.DOTALL) is None) and
        (re.match(r".*=\+.*", vba_code, re.DOTALL) is None)):
        return vba_code

    if debug_strip:
        print "HERE: 10"
        print vba_code
    if ("StrPtr" in vba_code):
        strptr_pat = r"(StrPtr\s*\(\s*)(\w+)(\s*\))"
        vba_code = re.sub(strptr_pat, r'\1"&\2"\3', vba_code)

    if debug_strip:
        print "HERE: 11"
        print vba_code
    if (":" in vba_code):
        label_pat = r"(\n\s*\w+:)([^\n])"
        vba_code = re.sub(label_pat, r'\1\n\2', vba_code)

        label_pat = r"(\n *\w+): *(?=\n)"
        vba_code = re.sub(label_pat, r'\1__LABEL_COLON__\n', vba_code)

        vba_code = vba_code.replace(" Do__LABEL_COLON__", " Do:").\
                   replace("\nDo__LABEL_COLON__", "\nDo:").\
                   replace(" Else__LABEL_COLON__", " Else:").\
                   replace("\nElse__LABEL_COLON__", "\nElse:")
        
    if debug_strip:
        print "HERE: 12"
        print vba_code
    vba_code = vba_code.replace("#if", "HASH__if")
    vba_code = vba_code.replace("#If", "HASH__if")
    vba_code = vba_code.replace("#else", "HASH__else")    
    vba_code = vba_code.replace("#Else", "HASH__else")
    vba_code = vba_code.replace("#end if", "HASH__endif")
    vba_code = vba_code.replace("#End If", "HASH__endif")

    if debug_strip:
        print "HERE: 13"
        print vba_code
    vba_code = re.sub(r"[Aa]s\s+#", "as__HASH", vba_code)
    vba_code = re.sub(r"[Pp]ut\s+#", "put__HASH", vba_code)
    vba_code = re.sub(r"[Gg]et\s+#", "get__HASH", vba_code)
    vba_code = re.sub(r"[Cc]lose\s+#", "close__HASH", vba_code)

    uni_vba_code = u""
    try:
        uni_vba_code = vba_code.decode("utf-8")
    except UnicodeDecodeError:
        log.warning("Converting VB code to unicode failed.")
    if debug_strip:
        print "HERE: 14"
        print vba_code
    pat = r"(?i)If\s+.{1,100}\s+Then\s*:[^\n]{1,100}\n"
    if (re2.search(unicode(pat), uni_vba_code) is not None):
        for curr_if in re.findall(pat, vba_code):
            new_if = curr_if.replace("Then:", "Then ")
            vba_code = vba_code.replace(curr_if, new_if)
    
    if debug_strip:
        print "HERE: 15"
        print vba_code
    uni_vba_code = u""
    try:
        uni_vba_code = vba_code.decode("utf-8")
    except UnicodeDecodeError:
        log.warning("Converting VB code to unicode failed.")
    pat = r"(?i)\n\s*If\s+.{1,100}\s+Then.{1,100}:.{1,100}(?:\s*Else.{1,100})?\n"
    single_line_ifs = []
    if (re2.search(unicode(pat), uni_vba_code) is not None):
        pos = 0
        for curr_if in re.findall(pat, vba_code):
            if_name = "HIDE_THIS_IF" + "_" * len(str(pos)) + str(pos)
            pos += 1
            vba_code = vba_code.replace(curr_if, "\n" + if_name + "\n")
            single_line_ifs.append((if_name, curr_if))
            
    if debug_strip:
        print "HERE: 16"
        print vba_code
    vba_code = vba_code.replace(":=", "__COLON_EQUAL__")
        
    vba_code = vba_code.replace("\nRem ", "\n' ")
    vba_code = vba_code.replace(" Rem ", " ' ")
    vba_code = vba_code.replace("\nREM ", "\n' ")
    vba_code = vba_code.replace(" REM ", " ' ")
    vba_code = vba_code.replace("\nrem ", "\n' ")
    vba_code = vba_code.replace(" rem ", " ' ")

    if debug_strip:
        print "HERE: 17"
        print vba_code
    vba_code = convert_colons_to_linefeeds(vba_code)

    if debug_strip:
        print "HERE: 17.1"
        print vba_code
    vba_code = break_up_whiles(vba_code)

    uni_vba_code = u""
    try:
        uni_vba_code = u"\n" + vba_code.decode("utf-8") + u"\n"
    except UnicodeDecodeError:
        pass
    elif_pat = "(\r?\n[^\"]*ElseIf.{5,50}Then)"
    if (re2.search(unicode(elif_pat), uni_vba_code) is not None):
        vba_code = re.sub(elif_pat, r"\1\n", vba_code)

    interesting_chars = [r'"', r'#', r"'", r"!", r"+", r"^",
                         r"PAT:[\x7f-\xff]", ";", r"[", r"]", "&"]
    
    in_str = False
    in_comment = False
    in_date = False    
    in_square_bracket = False
    num_square_brackets = 0
    prev_char = ""
    next_char = ""
    r = ""
    pos = -1
    if debug_strip:
        print "HERE: 18"
        print vba_code
    while (pos < (len(vba_code) - 1)):

        pos += 1
        c = vba_code[pos]
        if (pos > 1):
            prev_char = vba_code[pos - 1]
        if (pos < (len(vba_code) - 1)):
            next_char = vba_code[pos + 1]
        got_interesting = False
        curr_interesting_chars = interesting_chars
        if (in_comment):
            curr_interesting_chars.append("\n")
        for interesting_c in curr_interesting_chars:

            index = None
            if (interesting_c.startswith("PAT:")):
                interesting_c = interesting_c[len("PAT:"):]
                index = re.search(interesting_c, c)
                
            else:
                if (c == interesting_c):
                    index = 0

            if (index is None):

                continue

            got_interesting = True
            break

        if (not got_interesting):

            next_pos = len(vba_code)

            curr_interesting_chars = interesting_chars
            if (in_str):
                curr_interesting_chars = [r'"']
            if (in_comment):
                curr_interesting_chars = ["\n"]

            for interesting_c in curr_interesting_chars:

                index = None
                if (interesting_c.startswith("PAT:")):
                    interesting_c = interesting_c[len("PAT:"):]
                    index = re.search(interesting_c, vba_code[pos:])
                    if (index is not None):
                        index = index.start()
                    
                else:
                    if (interesting_c in vba_code[pos:]):
                        index = vba_code[pos:].index(interesting_c)

                if (index is None):

                    continue

                poss_pos = index + pos
                if (poss_pos < next_pos):
                    next_pos = poss_pos

            r += vba_code[pos:next_pos]

            pos = next_pos - 1
            continue
        
        if ((not in_comment) and (c == '"')):
            r += '"'
            in_str = not in_str
            continue

        if ((not in_comment) and (not in_str)):
            if (c == '['):
                num_square_brackets += 1
            if (c == ']'):
                num_square_brackets -= 1
            in_square_bracket = (num_square_brackets > 0)
            
        if ((not in_comment) and (not in_str) and (c == '#')):

            if (not in_date):

                end = pos + 1
                if ((end < len(vba_code)) and (vba_code[end] != ",") and (vba_code[end] != "\n")):

                    end = pos + 60
                    if (end > len(vba_code)):
                        end = len(vba_code)
                    got_hash = False
                    for i in range(pos + 1, end):
                        if (vba_code[i] == "#"):
                            got_hash = True
                            break

                    if (got_hash):
                        in_date = True

            else:
                in_date = False

        if ((not in_str) and (c == "'")):
            in_comment = True
        if (c == "\n"):
            in_comment = False
            r += "\n"

        if (in_str or in_comment or in_date):
            r += c
            continue

        if ((c == "!") and (next_char.isalpha())):
            r += "."
            continue

        if (c == "^"):
            r += " ^ "
            continue
        
        if (c == "&"):
            r += "&"
            continue

        if (c == "+"):
            if (prev_char != "="):
                r += " " + c
            continue
            

        if ((c == ";") and (prev_char == "&")):

            continue
            
        if (ord(c) > 127):
            r += "d" + str(ord(c))

        if (c == ";"):
            r += c

    r = r.replace("HASH__if", "#If")
    r = r.replace("HASH__else", "#Else")
    r = r.replace("HASH__endif", "#End If")

    r = r.replace("as__HASH", "As #")
    r = r.replace("put__HASH", "Put #")
    r = r.replace("get__HASH", "Get #")
    r = r.replace("close__HASH", "Close #")

    for if_info in single_line_ifs:
        r = r.replace(if_info[0], if_info[1])

    r = r.replace("__LABEL_COLON__", ":")

    r = r.replace("__COLON_EQUAL__", ":=")
    
    vba_code = "\n" + vba_code
    known_funcs = ["Randomize"]
    for func in known_funcs:
        r = r.replace("\n" + func + ":", "\n" + func)
    
    if debug_strip:
        print "HERE: 19"
        print r

    return r

def strip_comments(vba_code):
    """
    Strip comment lines from the VBA code.
    """

    vba_code = vba_code.replace("\n;", "\n'")
    if ("'" not in vba_code):
        return vba_code

    r = ""
    for curr_line in vba_code.split("\n"):

        if (not curr_line.strip().startswith("'")):
            r += curr_line + "\n"

    r = r.replace("\n\n\n", "\n")
        
    return r

defined_constants = set()
def find_defined_constants(vba_code):
    """
    Get the names of all the defined constants in the given VB code.
    """

    if ("Const " not in vba_code):
        return

    const_pat = r" Const +([\w_]+) "
    const_names = re.findall(const_pat, vba_code)

    defined_constants.update(const_names)
    
def rename_constants(vba_code):
    """
    Make sure constants have unique names to avoid overlap with function
    names.
    """

    if (len(defined_constants) == 0):
        return vba_code

    if (len(defined_constants) == 0):
        return vba_code

    changed = False
    if ("H" in defined_constants):
        vba_code = vba_code.replace("&H", "__HEX_STR__")
        changed = True
    
    for const_name in defined_constants:

        rep_pat = const_name + r"(\s*[^\(^=^ ^\w^_])"
        vba_code = re.sub(rep_pat, const_name + r"_CONST\1", vba_code)

        rep_pat = r"Const\s+(" + const_name + r")[\s=]"
        vba_code = re.sub(rep_pat, r"Const \1_CONST ", vba_code)

    if changed:
        vba_code = vba_code.replace("__HEX_STR__", "&H")
        
    return vba_code

def fix_vba_code(vba_code):
    """
    Fix up some substrings that SimulationVBA has problems parsing.
    """

    if debug_strip:
        print "FIX_VBA_CODE: 1"
        print vba_code
    vba_code = strip_comments(vba_code)

    if ("\n" in vba_code.strip()):

        last_line = vba_code[vba_code.strip().rindex("\n") + 1:].strip()
        if ("'" in last_line):
            last_line = last_line[:last_line.index("'")]

        if (last_line.endswith(".")):
            vba_code = vba_code.strip()
            vba_code = vba_code[:vba_code.rindex("\n")]
            vba_code += "\n"
    
    if debug_strip:
        print "FIX_VBA_CODE: 2"
        print vba_code
    vba_code = vba_code.replace("End SubPrivate", "End Sub\nPrivate")

    if debug_strip:
        print "FIX_VBA_CODE: 3"
        print vba_code
    vba_code = vba_code.replace("\x00", "")
    
    vba_code = re.sub(r"End\s+Try", "##End ##Try", vba_code)
    if debug_strip:
        print "FIX_VBA_CODE: 4"
        print vba_code
    
    if ("}" in vba_code):
        vba_code = re.sub(r"\r?\n *\} *\r?\n", "\n", vba_code)
    
    if debug_strip:
        print "FIX_VBA_CODE: 5"
        print vba_code
    linputs = re.findall(r"Line\s+Input\s+#\d+\s*,\s*\w+", vba_code, re.DOTALL)
    if (len(linputs) > 0):
        log.warning("VB Line Input constructs are not currently handled. Stripping them from code...")
    for linput in linputs:
        vba_code = vba_code.replace(linput, "")
    

    if debug_strip:
        print "FIX_VBA_CODE: 7"
        print vba_code
    implements = re.findall(r"Implements \w+", vba_code, re.DOTALL)
    if (len(implements) > 0):
        log.warning("VB Implements constructs are not currently handled. Stripping them from code...")
    for imp in implements:
        vba_code = vba_code.replace(imp, "")
        
    if debug_strip:
        print "FIX_VBA_CODE: 9"
        print vba_code
    brackets = re.findall(r"\(\[[^\]]+\]\)", vba_code, re.DOTALL)
    if (len(brackets) > 0):
        log.warning("([a1]) style constructs are not currently handled. Rewriting them...")
    for bracket in brackets:

        start = vba_code.index(bracket)
        in_str1 = False
        pos = start
        while (pos > 0):
            if (vba_code[pos] == "\n"):
                break
            if (vba_code[pos] == '"'):
                in_str1 = True
                break
            pos -= 1
        in_str2 = False
        while (pos < len(vba_code)):
            if (vba_code[pos] == "\n"):
                break
            if (vba_code[pos] == '"'):
                in_str2 = True
                break
            pos += 1
        if (in_str1 and in_str2):
            continue

        vba_code = vba_code.replace(bracket, "(" + bracket[2:-2] + ")")
    
    if debug_strip:
        print "FIX_VBA_CODE: 10"
        print vba_code
    vba_code = re.sub(r"([^\w_])_ *\r?\n", r"\1", vba_code)
    vba_code = "\n" + vba_code
    vba_code = re.sub(r"\n:", "\n", vba_code)

    if debug_strip:
        print "FIX_VBA_CODE: 11"
        print vba_code
    dumb_member_exps = re.findall(r"\n(?:\w+\.)+\n", vba_code)
    for dumb_exp in dumb_member_exps:
        log.warning("Commenting out bad line '" + dumb_exp.replace("\n", "") + "'.")
        safe_exp = "\n'" + dumb_exp[1:]
        vba_code = vba_code.replace(dumb_exp, safe_exp)

    if debug_strip:
        print "FIX_VBA_CODE: 12"
        print vba_code
    space_subs = re.findall(r"\n\s*Sub\s*\w+\s+\w+\s*\(", vba_code)
    for space_sub in space_subs:
        start = space_sub.index("Sub") + len("Sub")
        end = space_sub.rindex("(")
        sub_name = space_sub[start:end]
        new_name = sub_name.replace(" ", "_")
        log.warning("Replacing bad sub name '" + sub_name + "' with '" + new_name + "'.")
        vba_code = vba_code.replace(sub_name, new_name)
    
    if debug_strip:
        print "FIX_VBA_CODE: 13"
        print vba_code
    if (vba_code.count('\x0b') > 20):
        vba_code = vba_code.replace('\x0b', '')
    if (vba_code.count('\x88') > 20):
        vba_code = vba_code.replace('\x88', '')

    if debug_strip:
        print "FIX_VBA_CODE: 14"
        print vba_code
    vba_code = fix_difficult_code(vba_code)
    
    if debug_strip:
        print "FIX_VBA_CODE: 15.0"
        print vba_code
    vba_code = fix_skipped_1st_arg1(vba_code)
    if debug_strip:
        print "FIX_VBA_CODE: 15.1"
        print vba_code
    vba_code = fix_skipped_1st_arg2(vba_code)

    if debug_strip:
        print "FIX_VBA_CODE: 16"
        print vba_code
    vba_code = fix_unbalanced_quotes(vba_code)

    if debug_strip:
        print "FIX_VBA_CODE: 16.1"
        print vba_code
    vba_code = fix_bad_puts(vba_code)
    
    
    if debug_strip:
        print "FIX_VBA_CODE: 17"
        print vba_code
    vba_code = replace_constant_int_inline(vba_code)

    if debug_strip:
        print "FIX_VBA_CODE: 17.5"
        print vba_code
    vba_code = rename_constants(vba_code)

    if debug_strip:
        print "FIX_VBA_CODE: 17.6"
        print vba_code
    uni_vba_code = None
    try:
        uni_vba_code = vba_code.decode("utf-8")
    except UnicodeDecodeError:
        pass
    if (uni_vba_code is not None):
        bad_call_pat = "(\r?\n\s*[\w_]{2,50})\""
        if (re2.search(unicode(bad_call_pat), uni_vba_code)):
            vba_code = re.sub(bad_call_pat, r'\1 "', vba_code)
    
    if debug_strip:
        print "FIX_VBA_CODE: 18"
        print vba_code
    uni_vba_code = None
    try:
        uni_vba_code = vba_code.decode("utf-8")
    except UnicodeDecodeError:
        return vba_code
    got_multassign = (re2.search(u"(?:\w+\s*=\s*){2}", uni_vba_code) is not None)
    if ((" if+" not in vba_code) and
        (" If+" not in vba_code) and
        ("\nif+" not in vba_code) and
        ("\nIf+" not in vba_code) and
        (not got_multassign)):
        return vba_code
    
    r = ""
    if debug_strip:
        print "FIX_VBA_CODE: 19"
        print vba_code
    for line in vba_code.split("\n"):

        line = fix_multiple_assignments(line)
        
        if ("if+" not in line.lower()):
            
            r += line + "\n"
            continue

        in_str = False
        window = "   "
        new_line = ""
        for c in line:

            if (c == '"'):
                in_str = not in_str

            if ((not in_str) and (c == "+") and (window.lower() == " if")):

                new_line += " "

            else:
                new_line += c
                
            window = window[1:] + c

        r += new_line + "\n"

    r = strip_comments(r)
    
    if debug_strip:
        print "FIX_VBA_CODE: 20"
        print r
    return r

def replace_constant_int_inline(vba_code):
    """
    Replace constant integer definitions inline, but leave the definition
    behind in case the regex fails.
    """

    const_pattern = re.compile("(?i)const +([a-zA-Z][a-zA-Z0-9]{0,20})\s?=\s?(\d+)")
    d_const = dict()

    for const in re.findall(const_pattern, vba_code):
        d_const[const[0]] = const[1]
        
    if len(d_const) > 0:
        log.info("Found constant integer definitions, replacing them.")
    for const in d_const:
        this_const = re.compile('(?i)(?<=(?:[(), ]))' + str(const) + '(?=(?:[(), ]))(?!\s*=)')
        vba_code = re.sub(this_const, str(d_const[const]), vba_code)
    return(vba_code)

def strip_line_nums(line):
    """
    Strip line numbers from the start of a line.
    """

    pos = 0
    line = line.strip()
    for c in line:
        if (not c.isdigit()):
            if (c == ':'):
                return line
            break
        pos += 1
    return line[pos:]

def strip_attribute_lines(vba_code):
    """
    Strip all Attribute statements from the code.
    """
    if (vba_code is None):
        return vba_code
    r = ""
    for line in vba_code.split("\n"):
        if (line.strip().startswith("Attribute ")):
            log.warning("Attribute statements not handled. Stripping '" + line.strip() + "'.")
            continue
        r += line + "\n"
    return r

def strip_difficult_tuple_lines(vba_code):
    """
    Strip all calls like "foo.bar.baz (1,2)-(3,4),5" from the code.
    They are awful to parse with PyParsing.
    """
    if (vba_code is None):
        return vba_code
    r = ""
    tuple_pat = r"[\w_]{1,40}(?:\.[\w_]{1,40})+ +\( *[\w_\d\.]+ *(?:, *[\w_\d]+ *)+\ *\) *[-\+\*/]"
    for line in vba_code.split("\n"):
        line = line.strip()
        if (re.search(tuple_pat, line)):
            log.warning("Difficult member access expression with tuple arg not handled. Stripping '" + line.strip() + "'.")
            continue
        r += line + "\n"
    return r
        
external_funcs = []
def strip_useless_code(vba_code, local_funcs):
    """
    Strip statements that have no useful effect from the given VB. The
    stripped statements are commented out.
    """

    log.info("Modifying VB code...")
    vba_code = fix_vba_code(vba_code)

    vba_code = strip_comments(vba_code)

    vba_code = strip_difficult_tuple_lines(vba_code)
    
    exec_pat = r"execute(?:global)?\s*\("
    if (re.search(exec_pat, vba_code, re.IGNORECASE) is not None):
        r = strip_attribute_lines(vba_code)
        r = collapse_macro_if_blocks(r)
        return r
    
    change_callbacks = set()    
    
    assign_re = re2.compile(u"(?:\s*(\w+(?:\([^\)]*\))?(\.\w+)*)\s*=\s*)|(?:Dim\s+(\w+(\.\w+)*))")
    assigns = {}
    line_num = 0
    bool_statements = set(["If", "For", "Do"])
    global external_funcs
    for line in vba_code.split("\n"):

        line = strip_line_nums(line)
        
        line_num += 1
        if (line.strip().startswith("'")):
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("SKIP: Comment. Keep it.")
            continue
        
        if (("Declare" in line) and ("Lib" in line)):
            external_funcs.append(line.strip())
        
        if (("Sub " in line) and ("_Change(" in line)):

            data_name = line.replace("Sub ", "").\
                        replace("Private ", "").\
                        replace(" ", "").\
                        replace("()", "").\
                        replace("_Change", "").strip()
            change_callbacks.add(data_name)
        
        tmp_line = line
        if ("=" in line):
            tmp_line = line[:line.index("=") + 1]
        uni_tmp_line = ""
        try:
            uni_tmp_line = tmp_line.decode("utf-8")
        except UnicodeDecodeError:
            uni_tmp_line = tmp_line.decode("latin-1")
        match = assign_re.findall(uni_tmp_line)
        if (len(match) > 0):

            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("SKIP: Assign line: '" + line + "'. Line # = " + str(line_num))
                
            if (line.strip().startswith("While ")):
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("SKIP: While loop. Keep it.")
                continue

            if (".create" in line.strip().lower()):
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("SKIP: .Create() call. Keep it.")
                continue

            if (":" in line):
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("SKIP: Multi-statement line. Keep it.")
                continue

            if ((line.strip().lower().startswith("if ")) or
                (line.strip().lower().startswith("elseif "))):
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("SKIP: If statement. Keep it.")
                continue
            
            if (line.strip().lower().startswith("function ")):
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("SKIP: Function decl. Keep it.")
                continue

            if (" const " in line.lower()):
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("SKIP: Const decl. Keep it.")
                continue
                
            if (line.strip().endswith("_")):
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("SKIP: Continuation line. Keep it.")
                continue

            if (line.strip().startswith("#")):
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("SKIP: macro line. Keep it.")
                continue

            if (("sub " in line.lower()) or ("function " in line.lower())):
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("SKIP: Function definition. Keep it.")
                continue

            if (("GetObject" in line) or ("Shell" in line)):
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("SKIP: GetObject()/Shell() call. Keep it.")
                continue

            if ("Loop " in line):
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("SKIP: Loop statement. Keep it.")
                continue

            if ("while" in line.lower()):
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("SKIP: While statement. Keep it.")
                continue

            if (line.strip().startswith("Mid")):
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("SKIP: Mid() string update. Keep it.")
                continue

            if (is_interesting_call(line, external_funcs, local_funcs)):
                continue
            
            strip_line = line.strip()            
            skip = False
            for bool_statement in bool_statements:
                if (strip_line.startswith(bool_statement + " ")):
                    skip = True
                    break
            if (skip):
                continue

            if (strip_line.startswith(".") or (strip_line.lower().startswith("let ."))):
                continue

            if ('"' in line):
                eq_index = -1
                if ("=" in line):
                    eq_index = line.index("=")
                qu_index1 = -1
                if ('"' in line):
                    qu_index1 =  line.index('"')
                qu_index2 = -1
                if ('"' in line):
                    qu_index2 =  line.rindex('"')
                if ((qu_index1 < eq_index) and (qu_index2 > eq_index)):
                    continue
            
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("SKIP: Assigned vars = " + str(match) + ". Line # = " + str(line_num))
            for m in match:
                
                expanded_vars = []
                for var in m:
                    if (var is None):
                        continue
                    expanded_vars.append(var)
                    if ("(" in var):
                        array_var = var[:var.index("(")]
                        expanded_vars.append(array_var)
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("SKIP: Assigned vars (1) = " + str(expanded_vars) + ". Line # = " + str(line_num))
                for var in expanded_vars:

                    if ((var is None) or (len(var.strip()) == 0)):
                        continue

                    var = var.encode("ascii","ignore")
                
                    val = ""
                    if ("=" in val):
                        val = line[line.rindex("=") + 1:]
                    if ("." in val):
                        continue

                    if ("CreateObject" in val):
                        continue
                    
                    if (var.lower() in val.lower()):
                        continue

                    if ("." in var):
                        continue
                    
                    if (var not in assigns):
                        assigns[var] = set()
                    assigns[var].add(line_num)

    refs = {}
    for var in assigns.keys():
        refs[var] = ("." in var)
    line_num = 0
    for line in vba_code.split("\n"):

        line_num += 1
        if (line.strip().startswith("'")):
            continue
        
        for var in assigns.keys():
            
            if (line_num in assigns[var]):
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("STRIP: Var '" + str(var) + "' assigned in line '" + line + "'. Don't count as reference. " + " Line # = " + str(line_num))
                continue

            if (var.lower() in line.lower()):

                
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("STRIP: Var '" + str(var) + "' referenced in line '" + line + "'. " + " Line # = " + str(line_num))
                refs[var] = True

    for change_var in change_callbacks:
        for var in assigns.keys():
            refs[var] = ((change_var in var) or (var in change_var) or refs[var])
                
    comment_lines = set()
    keep_lines = set()
    for var in refs.keys():
        if (not refs[var]):
            for num in assigns[var]:
                comment_lines.add(num)
        else:
            for num in assigns[var]:
                keep_lines.add(num)

    tmp = set()
    for l in comment_lines:
        if (l not in keep_lines):
            tmp.add(l)
    comment_lines = tmp
        
    r = ""
    line_num = 0
    if_count = 0
    in_func = False
    for line in vba_code.split("\n"):

        line = strip_line_nums(line)
        
        if (("End Sub" in line) or ("End Function" in line)):
            in_func = False
        elif (("Sub " in line) or ("Function " in line)):
            in_func = True
        
        line_num += 1
        if (line.strip().startswith("If ")):
            if_count += 1

        if ((line_num in comment_lines) and
            (not line.strip().startswith("Function ")) and
            (not line.strip().startswith("Sub ")) and
            (not line.strip().startswith("End Sub")) and
            (not line.strip().startswith("End Function"))):
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("STRIP: Stripping Line (1): " + line)
            r += "' STRIPPED LINE\n"
            continue

        if (is_useless_call(line)):
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("STRIP: Stripping Line (2): " + line)
            r += "' STRIPPED LINE\n"
            continue

        if ((line.strip().lower().startswith("class ")) or
            (line.strip().lower().startswith("end class"))):
            log.warning("Classes not handled. Stripping '" + line.strip() + "'.")
            continue

        if (line.strip().startswith("Attribute ")):
            log.warning("Attribute statements not handled. Stripping '" + line.strip() + "'.")
            continue

        if ('.TypeText Text:="' in line.strip()):
            log.warning("'.TypeText Text:=\"' statements not handled. Stripping '" + line.strip() + "'.")
            continue
            

        tmp_line = line.lower().strip()
        func_ret_pat = r"function\s+\w+\(.*\)(?:\s+as\s+\w+)?"
        if ((tmp_line.startswith("function ")) and
            (not tmp_line.endswith(")")) and
            (re.match(func_ret_pat + r"$", tmp_line) is None)):
            match_obj = re.match(func_ret_pat, tmp_line)
            if (match_obj is not None):
                pos = match_obj.span()[1]
                tmp_line = line[:pos] + "\n"
                if ((pos + 1) < len(line)):
                    tmp_line += line[pos + 1:]
                line = tmp_line
        
        if ((line.lower().endswith("end function")) and
            (not line.strip().startswith("'")) and
            (len(line) > len("End Function"))):
            tmp_line = line[:-len("End Function")]
            r += tmp_line + "\n"
            r += "End Function\n"
            continue
            
        if (line.strip().startswith("Application.Run") and
            (re.match(r'Application\.Run\s+"[^"]+"', line) is not None) and
            (line.count('"') == 2) and
            (line.strip().endswith('"'))):

            line = line.replace('"', '').replace("Application.Run", "")

            fields = line.split(",")
            if ((len(fields) > 1) and (" " not in fields[0].strip())):
                line = "WScript.Shell " + line
        
        r += line + "\n"

    r = collapse_macro_if_blocks(r)

    return r
