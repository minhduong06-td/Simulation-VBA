"""@package read_ole_fields
Read in data values from OLE items like shapes and text boxes.
"""

"""
SimulationVBA is a specialized engine to parse, analyze and interpret Microsoft
VBA macros (Visual Basic for Applications), mainly for malware analysis.

Author: Philippe Lagadec - http://www.decalage.info
License: BSD, see source code or documentation

Project Repository:
https://github.com/decalage2/ViperMonkey
"""

#=== LICENSE ==================================================================


import io
import json
import subprocess
import struct
import logging
import zipfile
import tempfile
import re
import random
import os
import sys
from collections import Counter
import string

import olefile

from logger import log
import filetype
import ooxml_context

_thismodule_dir = os.path.normpath(os.path.abspath(os.path.dirname(__file__)))

def is_garbage_vba(vba, test_all=False, bad_pct=.6):
    """Check to see if the given supposed VBA is actually just a bunch of
    non-ASCII characters.

    @param vba (str) The VBA code to check.
    
    @param test_all (boolean) A flag indicating whether to look at all
    the code (True) or just the first part of the code (False).

    @param bad_pct (float) The max ratio of bad code to all code for
    this to be considered to be bad (i.e. percent bad divided by
    100).

    """

    if filetype.is_pe_file(vba, True):
        return True

    total_len = len(vba)
    if ((total_len > 50000) and (not test_all)):
        total_len = int(len(vba) * .25)
    if (total_len == 0):
        return False
    substr = vba[:total_len]

    if ("\n'" in substr):
        tmp = ""
        for line in substr.split("\n"):
            if (not line.strip().startswith("'")):
                tmp += line + "\n"
        substr = tmp
    
    num_bad = 0.0
    in_string = False
    for c in substr:
        if (c == '"'):
            in_string = not in_string
        if in_string:
            continue
        if (c not in string.printable):
            num_bad += 1

    return ((num_bad/total_len) > bad_pct)

def pull_base64(data):
    """Pull base64 strings from some data.

    @param data (str) The data from which to extract base64 strings.
    
    @return (list) A list of base64 strings found in the input.

    """

    base64_pat_loose = r"[A-Za-z0-9+/=]{40,}"
    r = set(re.findall(base64_pat_loose, data))
    return r

def unzip_data(data):
    """Unzip zipped data in memory.

    @param data (str) The data to unzip.

    @return (tuple) A 2 element tuple where the 1st element is the
    unzipped data and the 2nd element is the name of a temp file used
    in the unzipping process. Someone will need to clean this file
    up.

    """

    zip_magic = chr(0x50) + chr(0x4B) + chr(0x03) + chr(0x04)
    delete_file = False
    fname = None
    if data.startswith(zip_magic):
        f = tempfile.NamedTemporaryFile(delete=False)
        fname = f.name
        f.write(data)
        f.close()
        delete_file = True
    else:
        return (None, None)
        
    try:
        if (not zipfile.is_zipfile(fname)):
            if (delete_file):
                os.remove(fname)
            return (None, None)
    except OSError:
        if (delete_file):
            os.remove(fname)
        return (None, None)
        
    unzipped_data = zipfile.ZipFile(fname, 'r')

    return (unzipped_data, fname)

def _clean_2007_text(s):
    """Replace special 2007 formatting strings (XML escaped, etc.) with
    actual text.

    @param s (str) The string to clean.

    @return (str) The cleaned string.

    """    
    s = s.replace("&amp;", "&")\
         .replace("&gt;", ">")\
         .replace("&lt;", "<")\
         .replace("&apos;", "'")\
         .replace("&quot;", '"')\
         .replace("_x000d_", "\r")
    
    return s

def get_drawing_titles(data):
    """Read custom Drawing element title values from an Office 2007+
    file.
    
    @param data (str) The read in Office 2007+ file (data).

    @return (list) A list of 2 element tuples where the 1st tuple
    element is the name of the drawing element and the 2nd element is
    the title of the drawing element.

    """

    if (not filetype.is_office2007_file(data, True)):
        return []

    unzipped_data, fname = unzip_data(data)
    delete_file = (fname is not None)
    if (unzipped_data is None):
        return []

    zip_subfile = 'word/document.xml'
    if (zip_subfile not in unzipped_data.namelist()):
        if (delete_file):
            unzipped_data.close()
            os.remove(fname)
        return []

    f1 = unzipped_data.open(zip_subfile)
    contents = f1.read()
    f1.close()

    if (delete_file):
        unzipped_data.close()
        os.remove(fname)
    
    pat = r"<wp\:docPr id=\"(\d+)\" name=\"([^\"]*)\" title=\"([^\"]*)\""
    if (re.search(pat, contents) is None):
        return []
    drawings = re.findall(pat, contents)

    r = []
    for drawing_info in drawings:
        drawing_id = drawing_info[0]
        drawing_text = _clean_2007_text(drawing_info[2])
        var_name = "Shapes('" + drawing_id + "')"
        r.append((var_name, drawing_text))

    return r

def get_defaulttargetframe_text(data):
    """Read custom DefaultTargetFrame value from an Office 2007+ file.

    @param data (str) The read in Office 2007+ file (data).

    @return (str) On success return the DefaultTargetFrame value. On
    error return None.

    """

    if (not filetype.is_office2007_file(data, True)):
        return None

    unzipped_data, fname = unzip_data(data)
    delete_file = (fname is not None)
    if (unzipped_data is None):
        return None

    zip_subfile = 'docProps/custom.xml'
    if (zip_subfile not in unzipped_data.namelist()):
        if (delete_file):
            unzipped_data.close()
            os.remove(fname)
        return None

    f1 = unzipped_data.open(zip_subfile)
    contents = f1.read()
    f1.close()

    if (delete_file):
        unzipped_data.close()
        os.remove(fname)
    
    pat = r"<vt:lpwstr>([^<]+)</vt:lpwstr>"
    if (re.search(pat, contents) is None):
        return None
    r = _clean_2007_text(re.findall(pat, contents)[0])
    return r

def get_customxml_text(data):
    """Read custom CustomXMLParts text values from an Office 2007+ file.

    @param data (str) The read in Office 2007+ file (data).

    @return (list) A list of 2 element tuples where the 1st tuple
    element is the name of the custom XML part and the 2nd element is
    the text of the part.

    """

    if (not filetype.is_office2007_file(data, True)):
        return []

    unzipped_data, fname = unzip_data(data)
    delete_file = (fname is not None)
    if (unzipped_data is None):
        return []

    
    r = []
    for nn in range(1, 6):

        zip_subfile = 'customXml/item' + str(nn) + ".xml"
        if (zip_subfile not in unzipped_data.namelist()):
            continue

        f1 = unzipped_data.open(zip_subfile)
        contents = f1.read()
        f1.close()
    
        pat = r"<Item\d+>([^<]+)</Item\d+>"
        if (re.search(pat, contents) is None):
            continue
        txt_val = _clean_2007_text(re.findall(pat, contents)[0])

        var_name = "customxmlparts('activedocument.customxmlparts.count').selectnodes('//items')(" + str(nn) + ").childnodes('2').text"
        r.append((var_name, txt_val))

    if (delete_file):
        unzipped_data.close()
        os.remove(fname)

    return r
    
def get_msftedit_variables_97(data):
    """Looks for variable/text value pairs stored in an embedded rich
    edit control from an Office 97 doc. See
    https://docs.microsoft.com/en-us/windows/win32/controls/about-rich-edit-controls.

    @param data (str) The read in Office 97 file (data).

    @return (list) A list of 2 element tuples where the 1st tuple
    element is the name of the rich edit control variable and the 2nd
    element is the variable value.

    """

    pat = r"'\x01\xff\xff\x03.+?\x5c\x00\x70\x00\x61\x00\x72\x00\x0d\x00\x0a\x00\x7d"
    r = []
    for chunk in re.findall(pat, data, re.DOTALL):

        chunk = chunk.replace("\x00", "")
    

        name_pat = r"'\x01\xff\xff\x03\x92\x03\x04([A-Za-z0-9_]+)"
        names = re.findall(name_pat, chunk)

        if (len(names) != 1):
            name_pat = r"([A-Za-z0-9_]+)"
            tmp = re.findall(name_pat, chunk)
            names = []
            for poss_name in tmp:
                if (len(poss_name) < 30):
                    names.append(poss_name)
        
        data_pat = r"\\fs\d{1,3} (.+)\\par"
        chunk_data = re.findall(data_pat, chunk, re.DOTALL)
        if (len(chunk_data) != 1):
            continue
        chunk_data = chunk_data[0]

        for chunk_name in names:
            r.append((chunk_name, chunk_data))

    return r

def get_msftedit_variables(obj):
    """Looks for variable/text value pairs stored in an embedded rich edit
    control from an Office 97 or 2007+ doc.  See
    https://docs.microsoft.com/en-us/windows/win32/controls/about-rich-edit-controls.

    @param data (str) The read in Office 97 or 2007+ file (data).

    @return (list) A list of 2 element tuples where the 1st tuple
    element is the name of the rich edit control variable and the 2nd
    element is the variable value.

    """

    if obj[0:4] == '\xd0\xcf\x11\xe0':
        data = obj
    else:
        fname = obj
        try:
            f = open(fname, "rb")
            data = f.read()
            f.close()
        except IOError:
            data = obj
        except TypeError:
            data = obj

    if (filetype.is_office97_file(data, True)):
        return get_msftedit_variables_97(data)

    return []

def remove_duplicates(lst):
    """Remove duplicate subsequences from a list. Taken from
    https://stackoverflow.com/questions/49833528/python-identifying-and-deleting-duplicate-sequences-in-list/49835215.

    @param lst (list) The list from which to remove duplicate
    subsequences.
    
    @return (list) The list with duplicate subsequences removed.

    """

    lst = list(lst)
    lst.reverse()
    
    dropped_indices = set()
    counter = Counter(tuple(lst[i:i+2]) for i in range(len(lst) - 1))

    for i in range(len(lst) - 2, -1, -1):
        sub = tuple(lst[i:i+2])
        if counter[sub] > 1:
            dropped_indices |= {i, i + 1}
            counter[sub] -= 1

    r = [x for i, x in enumerate(lst) if i not in dropped_indices]
    r.reverse()
    return r

def entropy(text):
    """
    Compute the entropy of a string. Taken from
    https://rosettacode.org/wiki/Entropy#Uses_Python_2.
    
    @param text (str) The string for which to compute the entropy.
    """
    import math
    log2=lambda x:math.log(x)/math.log(2)
    exr={}
    infoc=0
    for each in text:
        try:
            exr[each]+=1
        except KeyError:
            exr[each]=1
    textlen=len(text)
    for _,v in exr.items():
        freq  =  1.0*v/textlen
        infoc+=freq*log2(freq)
    infoc*=-1
    return infoc


cruft_pats = [r'Microsoft Forms 2.0 Form',
              r'Embedded Object',
              r'CompObj',
              r'VBFrame',
              r'VERSION [\d\.]+\r\nBegin {[\w\-]+} \w+ \r\n\s+Caption\s+=\s+"UserForm1"\r\n\s+ClientHeight\s+=\s+\d+\r\n' + \
              r'\s+ClientLeft\s+=\s+\d+\r\n\s+ClientTop\s+=\s+\d+\r\n\s+ClientWidth\s+=\s+\d+\r\n' + \
              r'\s+StartUpPosition\s+=\s+\d+\s+\'CenterOwner\r\n\s+TypeInfoVer\s+=\s+\d+\r\nEnd\r\n',
              r'DEFAULT',
              r'InkEdit\d+',
              r'MS Sans Serif',
              r'\{\\rtf1\\ansi\\ansicpg\d+\\deff\d+\\deflang\d+\{\\fonttbl\{\\f\d+\\f\w+\\fcharset\d+.+;\}\}',
              r'{\\\*\\generator [\w\d\. ]+;}\\[\d\w]+\\[\d\w]+\\[\d\w]+\\[\d\w]+\\[\d\w]+ ',
              r'\\par',
              r'HelpContextID="\d+"',
              r'VersionCompatible\d+="\d+"',
              r'CMG="[A-Z0-9]+"',
              r'DPB="[A-Z0-9]+"',
              r'GC="[A-Z0-9]+"',
              r'\[Host Extender Info\]',
              r'&H\d+=\{[A-Z0-9]+\-[A-Z0-9]+\-[A-Z0-9]+\-[A-Z0-9]+\-[A-Z0-9]+\};VBE;&H\d+',
              r'&H\d+=\{[A-Z0-9]+\-[A-Z0-9]+\-[A-Z0-9]+\-[A-Z0-9]+\-[A-Z0-9]+\};Word\d.\d;&H\d+',
              r'\[Workspace\]',
              r'http://schemas.openxmlformats.org/\w+/\w+/\w+',
              r'Root Entry',
              r'Data',
              r'WordDocument',
              r'ObjectPool',
]

def _read_chunk(anchor, pat, data):
    """Read in delimited chunks of data based on an anchor at the start
    of the chunk and a pattern for recognizing a chunk.

    @param anchor (str) The anchor string at the start of the chunk to
    identify.

    @param pat (str) The regex pattern for identifying a chunk.

    @param data (str) The data from which to pull chunks.

    @return (list) A list of recognized chunks (str).

    """
    
    if (anchor not in data):
        return None
    data = data[data.index(anchor):]
    if (re.search(pat, data, re.DOTALL) is not None):
        return re.findall(pat, data, re.DOTALL)
    return None

def _get_field_names(vba_code, debug):
    """Get the names of object fields referenced in the given VBA code.

    @param vba_code (str) The VBA code to scan.

    @param debug (boolean) A flag indicating whether to print debug
    information.

    @return (list) A list of the names (str) of the object fields
    referenced in the VBA code.

    """

    object_names = set(re.findall(r"(?:ThisDocument|ActiveDocument|\w+)\.(\w+(?:\.ControlTipText)?)", vba_code))
    object_names.update(re.findall(r"(\w+)\.Caption", vba_code))
    
    page_pat = r"(?:ThisDocument|ActiveDocument|\w+)\.(Pages\(.+\))"
    if (re.search(page_pat, vba_code) is not None):

        for i in range(1, 10):
            object_names.add("Page" + str(i))

    object_names = clean_names(object_names)            
    if debug:
        print "\nget_ole_textbox_values2()"
        print "\nNames from VBA code:"
        print object_names
            
    control_tip_var_names = set()
    for name in object_names:

        if (name.endswith(".ControlTipText")):
            fields = name.split(".")[:-1]
            short_name = fields[-1]
            control_tip_var_names.add(short_name)

    return object_names, control_tip_var_names

def _read_large_chunk(data, debug):
    """
    Pull out a chunk of raw data containing mappings from object names to
    object text values.

    @param data (str) The Office 97 file data from which to pull an
    object name/value chunk.

    @param debug (boolean) A flag indicating whether to print debug
    information.

    @return (str) A chunk of data.
    """

    chunk_pats = [('ID="{',
                   r'ID="\{.{20,}(?:UserForm\d{1,10}=\d{1,10}, \d{1,10}, \d{1,10}, \d{1,10}, ' + \
                   r'\w{1,10}, \d{1,10}, \d{1,10}, \d{1,10}, \d{1,10}, \r\n){1,10}(.+?)Microsoft Forms '),
                  ('\x05\x00\x00\x00\x17\x00',
                   r'\x05\x00\x00\x00\x17\x00(.*)(?:(?:Microsoft Forms 2.0 Form)|(?:ID="{))'),
                  ('\xd7\x8c\xfe\xfb',
                   r'\xd7\x8c\xfe\xfb(.*)(?:(?:Microsoft Forms 2.0 Form)|(?:ID="{))'),
                  ('\x00V\x00B\x00F\x00r\x00a\x00m\x00e\x00',
                   r'\x00V\x00B\x00F\x00r\x00a\x00m\x00e\x00(.*)(?:(?:Microsoft Forms 2.0 (?:Form|Frame))|(?:ID="\{))')]
    for anchor, chunk_pat in chunk_pats:
        chunk = _read_chunk(anchor, chunk_pat, data)
        if (chunk is not None):
            if debug:
                print "\nCHUNK ANCHOR: '" + anchor + "'"
                print "CHUNK PATTERN: '" + chunk_pat + "'"
            break

    if (chunk is None):                
        if debug:
            print "\nNO VALUES"
        return None

    chunk = chunk[0]

    chunk = chunk.replace("\x02$", "").replace("\x01@", "")

    page_name_pat = r"Page(\d+)(?:(?:\-\d+)|[a-zA-Z]+)"
    chunk = re.sub(page_name_pat, r"Page\1", chunk)
    
    if debug:
        print "\nChunk:"
        print chunk

    return chunk

def _read_raw_strs(chunk, stream_names, debug):
    """Pull out all the ASCII strings from a given chunk of data.

    @param chunk (str) The data chunk from which to pull strings.
    
    @param stream_names (list) A list of the names of OLE streams in
    the Office OLE file. OLE stream names will not be counted as
    strings in the chunk.

    @return (list) A list of strings from the chunk.

    @param debug (boolean) A flag indicating whether to print debug
    information.

    """

    ascii_pat = r"(?:[\x09\x20-\x7f]|\x0d\x0a){4,}|(?:(?:[\x09\x20-\x7f]\x00|\x0d\x00\x0a\x00)){4,}"
    vals = re.findall(ascii_pat, chunk)
    vals = vals[:-1]
    tmp_vals = []
    for val in vals:

        val = val.replace("\x00", "")
        
        for cruft_pat in cruft_pats:
            val = re.sub(cruft_pat, "", val)
            
        if (len(val) == 0):
            continue
            
        if ((val.startswith("Taho")) or
            (val.startswith("PROJECT")) or
            (val.startswith("_DELETED_NAME_"))):
            continue

        if (val in stream_names):
            continue
        
        tmp_vals.append(val)

    vals = tmp_vals
    if debug:
        print "\nORIG RAW VALS:"
        print vals

    return vals

def _handle_control_tip_text(control_tip_var_names, vals, debug):
    """Find the text for each named control tip object.

    @param control_tip_var_names (list) The names (str) of the control
    tip objects.

    @param vals (list) Potential control tip text values (str).

    @param debug (boolean) A flag indicating whether to print debug
    information.

    @return (list) A list of 2 element tuples where the 1st element is
    the control tip object name and the 2nd is the control tip text.

    """

    r = []
    if debug:
        print "\nCONTROL TIP PROCESSING:"
    for name in control_tip_var_names:
        pos = -1
        for str_val in vals:
            pos += 1
            if ((str_val.startswith(name)) and ((pos + 1) < len(vals))):

                if (vals[pos + 1].startswith("ControlTipText")):
                    continue
                
                if debug:
                    print (name, vals[pos + 1])
                r.append((name, vals[pos + 1]))

                n = name
                if (len(n) > 2):
                    n = n[:-1]
                    r.append((n, vals[pos + 1]))
                if (len(n) > 2):
                    n = n[:-1]
                    r.append((n, vals[pos + 1]))
                if (len(n) > 2):
                    n = n[:-1]
                    r.append((n, vals[pos + 1]))

    return r

def _get_specific_values(chunk, stream_names, debug):
    """Get possible OLE object text values.

    @param chunk (str) A chunk of OLE data containing OLE object names
    and text values.

    @param stream_names (list) A list of the names of OLE streams in
    the Office OLE file. OLE stream names will not be counted as
    potential object values in the chunk.

    @param debug (boolean) A flag indicating whether to print debug
    information.

    @return (list) Potential OLE object text values (str).

    """

    val_pat = r"(?:[\x02\x10]\x00\x00([\x09\x20-\x7f]{2,}))|" + \
              r"((?:\x00[\x09\x20-\x7f]|\x00\x0d\x00\x0a){2,})|" + \
              r"(?:\x05\x80([\x09\x20-\x7f]{2,}))|" + \
              r"(?:[\x15\x0c\x0b]\x00\x80([\x09\x20-\x7f]{2,}(?:\x01\x00C\x00o\x00m\x00p\x00O\x00b\x00j.+[\x09\x20-\x7f]{5,})?))"
    vals = re.findall(val_pat, chunk.replace("\x19 ", "`\x00"))
    if debug:
        print "\nORIG SPECIFIC VALS:"
        print vals
    
    tmp_vals = []
    rev_vals = list(vals)
    rev_vals.reverse()
    seen = set()
    for val in rev_vals:

        if (len(val[0]) > 0):
            val = val[0]
        elif (len(val[1]) > 0):
            val = val[1]            
        else:
            val = val[2]

        compobj_pat = r"\x01\x00C\x00o\x00m\x00p\x00O\x00b\x00j"
        if (re.search(compobj_pat, val) is not None):
            tmp_val = ""
            ascii_pat = r"[\x20-\x7f]{5,}"
            for s in re.findall(ascii_pat, val):
                tmp_val += s
            val = tmp_val
            
        val = val.replace("\x00", "")
        
        for cruft_pat in cruft_pats:
            val = re.sub(cruft_pat, "", val)
            
        if (len(val) == 0):
            continue
            
        if ((val.startswith("Taho")) or
            (val.startswith("PROJECT")) or
            (val.startswith("_DELETED_NAME_")) or
            ("Normal.ThisDocument" in val)):
            continue

        if (val in seen):
            continue
        seen.add(val)

        if (val in stream_names):
            continue
        
        tmp_vals.append(val)

    tmp_vals.reverse()
    var_vals = tmp_vals

    if debug:
        print "\nORIG VAR_VALS:"
        print var_vals
    
    if (len(var_vals) > 4):
        num_random = 0
        for s in var_vals[:5]:
            if (entropy(s) > 3.0):
                num_random += 1
            elif (s[0].isupper() and s[1:].islower()):
                num_random += 1
        if (num_random >= 3):
            var_vals = var_vals[1:]

    var_vals = remove_duplicates(var_vals)
    
    longest_val = ""
    tmp_vals = []
    for v in var_vals:
        if (v not in tmp_vals):
            tmp_vals.append(v)
        if (len(v) > len(longest_val)):
            longest_val = v
    var_vals = tmp_vals

    return var_vals, longest_val

def _get_specific_names(object_names, chunk, control_tip_var_names, debug):
    """Get possible OLE object names.

    @param object_names (list) A list of the names (str) of the object fields
    referenced in the VBA code.

    @param chunk (str) A chunk of OLE data containing OLE object names
    and text values.

    @param control_tip_var_names (list) The names (str) of the control
    tip objects.

    @param debug (boolean) A flag indicating whether to print debug
    information.

    @return (list) OLE object names (str) that appear in the given
    chunk.

    """

    name_pat1 = r"(?:(?:\x17\x00)|(?:\x00\x80))(\w{2,})"
    name_pat = r"(?:" + name_pat1 + ")|("
    first = True
    for object_name in object_names:
        if (not first):
            name_pat += "|"
        first = False
        if ("." in object_name):
            object_name = object_name[:object_name.index(".")]
        name_pat += object_name
    name_pat += ")"
    names = re.findall(name_pat, chunk)
    if debug:
        print "\nORIG NAMES:"
        print names

    tmp_names = []
    for name in names:
        if (len(name[0]) > 0):
            name = name[0]
        else:
            name = name[1]
        if (name in control_tip_var_names):
            continue
        if (name in tmp_names):
            continue
        tmp_names.append(name)
    var_names = tmp_names

    return var_names

def get_ole_textbox_values2(data, debug, vba_code, stream_names):
    """Read in the text associated with embedded OLE form textbox
    objects (hack!). NOTE: This currently is a really NASTY hack.

    @param data (str) The read in Office 97 file (data).

    @param debug (boolean) A flag indicating whether to print debug
    information.

    @param vba_code (str) The VBA macro code from the Office file.

    @param stream_names (list) A list of the names of OLE streams in
    the Office OLE file.

    @return (list) A list of 2 element tuples where the 1st element is
    the object name and the 2nd is the object text.

    """

    object_names, control_tip_var_names = _get_field_names(vba_code, debug)
    
    chunk = _read_large_chunk(data, debug)
    if (chunk is None):                
        return []
    
    vals = _read_raw_strs(chunk, stream_names, debug)

    r = _handle_control_tip_text(control_tip_var_names, vals, debug)


    var_names = _get_specific_names(object_names, chunk, control_tip_var_names, debug)

    var_vals, longest_val = _get_specific_values(chunk, stream_names, debug)
    
    if (len(var_names) > len(var_vals)):
        var_names = var_names[:len(var_vals)]
        
    if debug:
        print "\nROUND 2:\nNAMES:"
        print var_names
        print "\nVALS:"
        print var_vals
    
    pos = -1
    hack_names = set(["Page1", "Label1"])
    for name in var_names:

        pos += 1
        if ((name in hack_names) and (len(longest_val) > 30)):
            val = longest_val

        else:
            val = var_vals[pos]
            if (val.endswith('o')):
                val = val[:-1]
            elif (val.endswith("oe")):
                val = val[:-2]

        r.append((name, val))

        n = name
        if (len(n) > 2):
            n = n[:-1]
            r.append((n, val))
        if (len(n) > 2):
            n = n[:-1]
            r.append((n, val))
        if (len(n) > 2):
            n = n[:-1]
            r.append((n, val))

    if debug:
        print "\nRESULTS VALUES2:"
        print r
    return r

def get_ole_textbox_values1(data, debug, stream_names):
    """Read in the text associated with embedded OLE form textbox
    objects (hack!). NOTE: This currently is a really NASTY hack. 

    @param data (str) The read in Office 97 file (data).

    @param debug (boolean) A flag indicating whether to print debug
    information.

    @param stream_names (list) A list of the names of OLE streams in
    the Office OLE file.

    @return (list) A list of 2 element tuples where the 1st element is
    the object name and the 2nd is the object text.

    """


    if debug:
        print "\nget_ole_textbox_values1"

    chunk_pat = r'DPB=".*"\x0d\x0aGC=".*"\x0d\x0a(.*;Word8.0;&H00000000)'
    chunk = re.findall(chunk_pat, data, re.DOTALL)

    if (len(chunk) == 0):
        if debug:
            print "\nNO VALUES"
        return []
    chunk = chunk[0]

    ignore_pat = r"\[Host Extender Info\]\x0d\x0a&H\d+={[A-Z0-9\-]+};VBE;&H\d+\x0d\x0a&H\d+={[A-Z0-9\-]+}?"
    chunk = re.sub(ignore_pat, "", chunk)
    if ("\x00\x01\x01\x40\x80\x00\x00\x00\x00\x1b\x48\x80" in chunk):
        start = chunk.index("\x00\x01\x01\x40\x80\x00\x00\x00\x00\x1b\x48\x80")
        chunk = chunk[start+1:]

    page_name_pat = r"Page(\d+)(?:(?:\-\d+)|[a-zA-Z]+)"
    chunk = re.sub(page_name_pat, r"Page\1", chunk)
        
    ascii_pat = r"(?:[\x20-\x7f]|\x0d\x0a){5,}"
    vals = re.findall(ascii_pat, chunk)
    vals = vals[:-1]
    tmp_vals = []
    for val in vals:

        if (val.startswith("Taho")):
            continue
        if (val in stream_names):
            continue
        tmp_vals.append(val)
    vals = tmp_vals
    if debug:
        print "\n---------------"
        print "Values:"
        print chunk
        print vals
        print len(vals)


    name_pat = r"\\MSForms.exd(.*)Microsoft Forms 2.0 Form\x00\x10\x00\x00\x00Embedded Object"
    chunk = re.findall(name_pat, data, re.DOTALL)

    if (len(chunk) == 0):
        if debug:
            print "\nNO NAMES"
        return []
    chunk_orig = chunk[0]

    if ("C\x00o\x00m\x00p\x00O\x00b\x00j" not in chunk_orig):
        if debug:
            print "\nNO NARROWED DOWN CHUNK"
        return []
    
    start = chunk_orig.index("C\x00o\x00m\x00p\x00O\x00b\x00j")
    chunk = chunk_orig[start + len("C\x00o\x00m\x00p\x00O\x00b\x00j"):]
    if debug:
        print "\n---------------"
        print "Names:"
        print chunk

    names = re.findall(ascii_pat, chunk)
    if (len(names) > 0):
        names = names[:-1]
    if (len(names) == 0):
        if ("Document" not in chunk_orig):
            if debug:
                print "\nNO NAMES, NO Document IN CHUNK"
            return []
        start = chunk_orig.index("Document")
        chunk = chunk_orig[start + len("Document"):]
        names = re.findall(ascii_pat, chunk)
        names = names[:-1]
    if debug:
        print names
        print len(names)

    if (len(names) > len(vals)):
        if debug:
            print "\nNOT SAME # NAMES/VALS"
        names = names[len(names) - len(vals):]

    pos = -1
    r = []
    for n in names:
        pos += 1
        r.append((n, vals[pos]))

        if (len(n) > 2):
            n = n[:-1]
            r.append((n, vals[pos]))
        if (len(n) > 2):
            n = n[:-1]
            r.append((n, vals[pos]))
        if (len(n) > 2):
            n = n[:-1]
            r.append((n, vals[pos]))

    if debug:
        print "\n-----------\nResult:"
        print r
    return r

def get_vbaprojectbin(data):
    """Pull the vbaProject.bin file from a 2007+ Office (ZIP) file.

    @param data (str) Already read in 2007+ file contents.

    @return (str) On success return the read in contents of
    vbaProject.bin. On error return None.

    """

    if (not filetype.is_office2007_file(data, True)):
        return None

    unzipped_data, fname = unzip_data(data)
    delete_file = (fname is not None)
    if (unzipped_data is None):
        return None

    subfile_names = ['word/vbaProject.bin', 'xl/vbaProject.bin']
    zip_subfile = None
    for subfile in subfile_names:
        if (subfile in unzipped_data.namelist()):
            zip_subfile = subfile
            break
    if (zip_subfile is None):
        if (delete_file):
            os.remove(fname)
        return None

    f1 = unzipped_data.open(zip_subfile)
    r = f1.read()
    f1.close()

    if (delete_file):
        unzipped_data.close()
        os.remove(fname)
    return r

def strip_name(poss_name):
    """Remove bad characters from a potential OLE object name.

    @param poss_name (str) The potential object name.

    @return (str) The given name with bad characters stripped out.

    """
    
    name = re.sub(r"[^A-Za-z\d_]", r"", poss_name)
    return name.strip()

def is_name(poss_name):
    """Check a given string to see if it could be an OLE object name.

    @param poss_name (str) The string to check.

    @return (boolean) True if the given string could be an object
    name, False if not.

    """
    
    if (poss_name is None):
        return False

    name_pat = r"[a-zA-Z]\w*"
    if (re.match(name_pat, poss_name) is None):
        return False

    bad_chars = re.findall(r"[^A-Za-z0-9_]", poss_name)
    return (len(bad_chars) < 5)
    
def clean_names(names):
    """Strip out bad characters from the given OLE object names.

    @param names (list) A list of object names (str) to clean.

    @return (set) A set of cleaned names.

    """
    
    r = set()    
    for poss_name in names:
        poss_name = poss_name.strip()
        if (is_name(poss_name)):
            r.add(poss_name)
    return r

def _get_stream_names(vba_code):
    """Pull the names of OLE streams from olevba output.

    @param vba_code (str) The olevba output for the Office file being
    analyzed.

    @return (list) The names of the OLE streams pulled from the olevba
    output.

    """
    stream_pat = r'Attribute VB_Name = "([\w_]+)"'
    return re.findall(stream_pat, vba_code)    

def _find_name_in_data(object_names, found_names, strs, debug):
    """Look for a VBA name in the string values pulled from a chunk of an
    Office 97 file.

    @param object_names (list) A list of the names (str) of the object
    fields referenced in the Office file's VBA code. These are the
    names being looked for.

    @param found_names (set) Names that we have already found.

    @param strs (list) All of the ASCII strings found in the current
    file chunk being analyzed.

    @param debug (boolean) A flag indicating whether to print debug
    information.

    @return (tuple) A 3 element tuple where the 1st element is the
    last checked position in the string list, the 2nd element is the
    position in the string list where the name was found, and the 3rd
    element is the name that was found.

    """

    curr_pos = 0
    name_pos = 0
    name = None
    page_pat = r"(Page\d+)(?:[A-Za-z]+[A-Za-z0-9]*)?"
    for field in strs[::-1]:
        poss_name = field.replace("\x00", "").replace("\xff", "").strip()
        if (re.search(page_pat, poss_name) is not None):
            poss_name = re.findall(page_pat, poss_name)[0]
        if ((poss_name in object_names) and (poss_name not in found_names)):

            name = poss_name
            name_pos = curr_pos
            if debug:
                print "\nFound referenced name: " + name
            break
        curr_pos += 1

    curr_pos = len(strs) - curr_pos - 1
    name_pos = len(strs) - name_pos - 1
    if (name is not None):
        curr_pos = -1
        max_val = ""
        for field in strs:
            curr_pos += 1
            if ((field == name) and
                ((curr_pos + 1) < len(strs)) and
                (len(strs[curr_pos + 1]) > len(max_val))):
                max_val = strs[curr_pos + 1]
                name_pos = curr_pos
        
    return (curr_pos, name_pos, name)

def _find_repeated_substrings(s, chunk_size, min_str_size):
    """Find all of the repeated substrings in a given string that are longer
    than a certain length. This assumes that repeated substrings of interest
    show up in a prefix of a given size.

    @param s (str) The string to check for repeated substrings. Only
    a prefix of the string will be checked.

    @param chunk_size (int) The size of the string prefix to check for
    repeated substrings. If bigger than the given string length an
    empty set will be returned.

    @param min_str_size (int) The minimum substring size to
    track. Shorter repeated substrings will not be reported.

    @return (set) A set of repeated substrings.

    """
    
    if (chunk_size > len(s)):
        return set()
    chunk = s[:chunk_size]

    pos = -1
    window_size = 2
    r = set()
    while ((pos + window_size) < len(chunk)):

        pos += 1
        curr_str = chunk[pos:pos + window_size]
        if (s.count(curr_str) > 1):

            tmp_window_size = 3
            old_curr_str = None
            while ((s.count(curr_str) > 1) and
                   ((pos + tmp_window_size) < len(chunk))):
                old_curr_str = curr_str
                curr_str = chunk[pos:pos + tmp_window_size]
                tmp_window_size += 1

            if ((old_curr_str is not None) and
                (len(old_curr_str.strip()) >= min_str_size)):

                r.add(old_curr_str)

                if (len(old_curr_str) > min_str_size*3):
                    for i in range(1, len(old_curr_str) - min_str_size*3):
                        r.add(old_curr_str[:i*-1])

    return r

def _find_most_repeated_substring(strs):
    """Find the most common repeated substring in a given list of strings.

    @param strs (list) The strings to check for the most common
    repeated substring.

    @return (str) The most common repeated substring if any were
    found. If no repeats are found None will be returned.

    """
    
    all_substs = set()
    for s in strs:
        all_substs = all_substs.union(_find_repeated_substrings(s, 300, 4))
        
    if (len(all_substs) == 0):
        return None
        
    max_repeats = -1
    max_subst = ""
    for curr_subst in all_substs:
        curr_repeats = 0
        for s in strs:
            curr_repeats += s.count(curr_subst)
        if (curr_repeats < 5):
            continue
        if (curr_repeats * len(curr_subst) > max_repeats * len(max_subst)):
            max_repeats = curr_repeats
            max_subst = curr_subst

    if (max_subst == ""):
        max_subst = None
    return max_subst

def _find_str_with_most_repeats(strs):
    """Find the string in the given list of strings that contains the most
    instances of some repeated substring. In more detail, this finds
    the most commonly repeated substring in all the given strings and
    then finds the given string that contains the most repeats of the
    most common repeated substring.

    @param strs (list) The strings to check.

    @return (str) If repeated substrings were found return the given
    string that has the most repeats of the most common repeated
    substring. If no repeated substrings were found None is returned.

    """
    
    max_subst = _find_most_repeated_substring(strs)
    if (max_subst is None):
        return (None, None)
    
    max_count = -1
    max_str = None
    for s in strs:
        curr_count = s.count(max_subst)
        if (curr_count > max_count):
            max_count = curr_count
            max_str = s

    return (max_str, max_subst)

def get_ole_text_method_1(vba_code, data, debug=False):
    """Pull OLE object name/value pairs from given OLE data using
    heuristic 1.

    @param vba_code (str) The VBA macro code from the Office file.

    @param data (str) The read in Office 97 file (data).

    @param debug (boolean) A flag indicating whether to print debug
    information.

    @return (list) A list of 2 element tuples where the 1st element is
    the object name and the 2nd is the object text.

    """
    
    debug1 = debug
    
    if debug1:
        print "\n\nSTART get_ole_text_method_1 !!!!"
    data = re.sub(r"[\x20-\x7e]\x00(?:\xe5|\xd5)", "", data)
    data = data.replace("\x02$", "").\
           replace("\x01@", "").\
           replace("0\x00\xe5", "").\
           replace("\xfc", "").\
           replace("\x19 ", "").\
           replace("_epx" + chr(223), "").\
           replace("R\x00o\x00o\x00t\x00 \x00E\x00n\x00t\x00r\x00y", "").\
           replace("Embedded Object", "").\
           replace("mbedded Object", "").\
           replace("bedded Object", "").\
           replace("edded Object", "").\
           replace("dded Object", "").\
           replace("ded Object", "").\
           replace("ed Object", "").\
           replace("d Object", "").\
           replace("jd\x00\x00", "\x00").\
           replace("\x00\x00", "\x00").\
           replace("\x0c%", "")
    if (re.search(r"\x00.%([^\x00])\x00", data) is not None):
        data = re.sub(r"\x00.%([^\x00])\x00", "\x00\\1\x00", data)
    data = data.replace("\r", "__CARRIAGE_RETURN__")
    data = data.replace("\n", "__LINE_FEED__")
    if (re.search(r"\x00([ -~])[^ -~\x00]([ -~])\x00", data) is not None):
        data = re.sub(r"\x00([ -~])[^ -~\x00]([ -~])\x00", "\x00\\1\x00\\2\x00", data)
    if (re.search(r"\x00[ -~]{2}([ -~])\x00", data) is not None):
        data = re.sub(r"\x00[ -~]{2}([ -~])\x00", "\x00\\1\x00", data)
    data = re.sub(r"\x00[^ -~]", "", data)
    if (re.search(r"\x00([ -~])[^ -~\x00]([ -~])\x00", data) is not None):
        data = re.sub(r"\x00([ -~])[^ -~\x00]([ -~])\x00", "\x00\\1\x00\\2\x00", data)
    data = data.replace("__CARRIAGE_RETURN__", "\r")
    data = data.replace("__LINE_FEED__", "\n")
    if debug1:
        print data
        print "\n\n\n"

    ascii_pat = r"(?:[\r\n\x09\x20-\x7f]|\x0d\x0a){4,}|(?:(?:[\r\n\x09\x20-\x7f]\x00|\x0d\x00\x0a\x00)){4,}"
    vals = re.findall(ascii_pat, data)
    tmp_vals = []
    for val in vals:
        
        val = val.replace("\x00", "")
        
        for cruft_pat in cruft_pats:
            val = re.sub(cruft_pat, "", val)
            
        if (len(val) == 0):
            continue
            
        if ((val.startswith("Taho")) or
            (val.startswith("PROJECT")) or
            (val.startswith("_DELETED_NAME_"))):
            continue

        if (val.strip().startswith("<!DOCTYPE html")):
            continue
        
        tmp_vals.append(val)
        if debug1:
            print "+++++++++++++++"
            print val

    max_substs, repeated_subst = _find_str_with_most_repeats(tmp_vals)
    if (max_substs is None):
        if debug1:
            print "DONE!! NO REPEATED SUBSTRINGS!!"
        return None
    if debug1:
        print "\n"
        print "*************"
        print "MAX SUBSTS"
        print max_substs
        print "\n"
        print "*************"
        print "REPEATED SUBST"
        print repeated_subst
    
    if debug1:
        print "LEN MAX STR: " + str(len(max_substs))
        print "MAX REPEATS IN 1 STR: " + str(max_substs.count(repeated_subst))
        print "REPEATED STR: '" + repeated_subst + "'"
    if ((len(max_substs) < 100) or (max_substs.count(repeated_subst) < 20)):
        if debug1:
            print "DONE!! TOO FEW REPEATED SUBSTRINGS!!"
        return None

    aggregate_str = ""
    obj_pat = r'VERSION \d\.\d{1,5}\r\n' + \
              r'Begin \{\w{2,20}\-\w{2,20}\-\w{2,20}\-\w{2,20}\-\w{2,20}\} \w{2,20} \r\n' + \
              r' {1,10}Caption {1,30}= {1,30}"\w{1,20}"\r\n' + \
              r' {1,10}ClientHeight {1,30}= {1,30}\d{1,20}\r\n' + \
              r' {1,10}ClientLeft {1,30}= {1,30}\d{1,20}\r\n' + \
              r' {1,10}ClientTop {1,30}=? {0,30}(?:\d{3})?'
    for val in tmp_vals:

        val = val.replace("\x00", "")
        if (len(val) == 0):
            continue

        pct = (val.count(repeated_subst) * len(repeated_subst)) / float(len(val)) * 100
        if (pct > 30):



            first_half_rep = None
            second_half_rep = None
            matched_agg_str = ""
            for end_pos in range(0, 3):
                if debug1:
                    print "CHECK !!!!!!!!!!!!!"
                    print "chopping off " + str(end_pos)
                curr_agg_str = aggregate_str[:-end_pos]
                for i in range(1, len(repeated_subst) + 1):
                    curr_first_half = repeated_subst[:i]
                    if debug1:
                        print "++++"
                        print "curr 1st half"
                        print curr_first_half
                        print "curr 1st half string end"
                        print curr_agg_str[-len(curr_first_half):]
                    if (curr_agg_str.endswith(curr_first_half) and
                        (len(curr_agg_str) > len(matched_agg_str))):
                        if debug1:
                            print "MATCH!!"
                        matched_agg_str = curr_agg_str
                        first_half_rep = curr_first_half
                        second_half_rep = repeated_subst[i:]

            if (first_half_rep == repeated_subst):
                first_half_rep = None
                second_half_rep = None

            if (matched_agg_str != ""):
                aggregate_str = matched_agg_str
                
            if (first_half_rep is not None):

                if debug1:
                    print "FIRST HALF!!"
                    print first_half_rep
                    print "SECOND HALF!!"
                    print second_half_rep
                
                start_pos = 0
                while (start_pos < len(val)):
                    if (val[start_pos:].startswith(second_half_rep)):
                        break
                    start_pos += 1
                if debug1:
                    print "SKIP 2nd HALF!!"
                    print val[:start_pos]
                val = val[start_pos:]

            else:

                if (repeated_subst in val):
                    start_pos = val.index(repeated_subst)                    
                    while (((start_pos - 1) >= 0) and
                           (re.match("[A-Za-z]", val[start_pos - 1]) is not None)):
                        start_pos -= 1
                    val = val[start_pos:]
                else:
                    val = re.sub(obj_pat, "", val)
                
            aggregate_str += val
        if debug1:
            print "-------"
            print val.strip()
            print pct
    if (len(aggregate_str) == 0):
        aggregate_str = max_substs
        
    object_names = set(re.findall(r"(?:ThisDocument|ActiveDocument|\w+)\.(\w+)", vba_code))
    object_names.update(re.findall(r"(\w+)\.Caption", vba_code))
    object_names.update(re.findall(r"(\w+) *_? *(?:\r?\n)? *\. *_? *(?:\r?\n)? *Content", vba_code))
    
    page_pat = r"((?:Pages|Tabs|InlineShapes|Item).?\(.+\))"
    if (re.search(page_pat, vba_code) is not None):

        for i in range(1, 10):
            object_names.add("Page" + str(i))

    if (".StoryRanges" in vba_code):

        for i in range(1, 10):
            object_names.add("StoryRanges.Items('" + str(i) + "')")
            object_names.add("StoryRanges('" + str(i) + "')")
            object_names.add("StoryRanges.Items(" + str(i) + ")")
            object_names.add("StoryRanges(" + str(i) + ")")
            
    object_names = clean_names(object_names)
    if debug1:
        print "\nFINAL:"
        print aggregate_str
        print object_names
        sys.exit(0)

    
    r = []
    for curr_object in object_names:
        r.append((curr_object, aggregate_str))
    return r

def _get_next_chunk(data, index, form_str, form_str_pat, end_object_marker):
    """Get the next chunk of OLE object name/value information from the
    given OLE data.

    @param data (str) The read in Office 97 file (data).

    @param index (int) The position in the OLE data from which to
    start looking for the next chunk.

    @param form_str (str) A string marking the start of the data for
    an OLE form.

    @param form_str_pat (str) A regex for recognizing strings marking
    the start of an OLE form.

    @param end_object_marker (str) The string marking the end of a
    chunk.

    @return (tuple) A 3 element tuple where the 1st element is the
    next chunk, the 2nd element is the index of the start of the chunk
    and the 3rd element is the index of the end of the chunk.

    """

    search_r = re.search(form_str_pat, data[index:])
    index = search_r.start() + index
    start = index + len(search_r.group(0))
    while ((start < len(data)) and (ord(data[start]) in range(32, 127))):
        start += 1

    if ((form_str in data[start:]) and
        (end_object_marker in data[start:]) and
        (data[start:].index(end_object_marker) < data[start:].index(form_str))):

        end = data[start:].index(end_object_marker) + start

    elif (form_str in data[start:]):

        end = data[start:].index(form_str) + start

    elif (end_object_marker in data[start:]):

        end = data[start:].index(end_object_marker) + start

    else:

        end = index + 2500000
        if (end > len(data)):
            end = len(data) - 1

    chunk = data[index : end]

    return (chunk, index, end)

def _pull_object_names(vba_code):
    """Pull out the names of object fields referenced in the given VBA
    code.

    @param vba_code (str) The VBA macro code from the Office file.

    @return (tuple) A 2 element tuple, where the 1st element is a set
    of object field names and the 2nd element is a set of page
    (PageNN) object field names.

    """

    object_names = set(re.findall(r"(?:ThisDocument|ActiveDocument|\w+)\.(\w+)", vba_code))
    object_names.update(re.findall(r"(\w+)\.Caption", vba_code))
    
    page_pat = r"(?:ThisDocument|ActiveDocument|\w+)\.(Pages\(.+\))"
    page_names = set()
    if (re.search(page_pat, vba_code) is not None):

        for i in range(1, 10):
            object_names.add("Page" + str(i))
            page_names.add("Page" + str(i))

    object_names = clean_names(object_names)

    return (object_names, page_names)

def _guess_name_from_data(strs, field_marker, debug):
    """Use heuristics to guess the object name in the given list of
    strings pulled from a chunk of OLE data.

    @param strs (list) The strings pulled from the OLE chunk.

    @param field_marker (str) A string that marks the start of an
    object text value.

    @param debug (boolean) A flag indicating whether to print debug
    information.

    @return (tuple) A 2 element tuple where the 1st element is the
    position in the given list of strings where the name was found and
    the 2nd element is the name. If a name was not found the 2nd
    element will be None.

    """

    name_pos = None
    name = None
    curr_pos = 0
    for field in strs:
    
        if (field.startswith(field_marker)):
    
            poss_name = None
            if ((curr_pos + 1) < len(strs)):
                poss_name = strs[curr_pos + 1].replace("\x00", "").replace("\xff", "").strip()
            skip_names = set(["contents", "ObjInfo", "CompObj"])
            if ((poss_name is not None) and
                ((not poss_name.startswith("_")) or
                 (not poss_name[1:].isdigit())) and
                (poss_name not in skip_names)):
    
                name = poss_name
                name_pos = curr_pos + 1
    
            break

        curr_pos += 1

    if (name is None):

        name_marker = "OCXNAME"
        for field in strs:
            if (field.replace("\x00", "") == 'OCXPROPS'):
                name_marker = "OCXPROPS"

        curr_pos = 0
        if debug:
            print "\nName Marker: " + name_marker
        for field in strs:

            if debug:
                print "\nField: '" + field.replace("\x00", "") + "'"
            if (field.replace("\x00", "") != name_marker):
                curr_pos += 1
                continue
                

            poss_name = strs[curr_pos + 1].replace("\x00", "")
            if debug:
                print "\nTry: '" + poss_name + "'"
            if (poss_name.startswith("_") and poss_name[1:].isdigit()):

                curr_pos += 1
                continue

            if (poss_name != 'contents'):

                name = poss_name
                break

            name_pos = curr_pos + 1
            poss_name = strs[curr_pos + 2].replace("\x00", "")
            if debug:
                print "\nTry: '" + poss_name + "'"
                            
            if ((not poss_name.startswith("_")) or
                (not poss_name[1:].isdigit())):

                name = poss_name
                name_pos = curr_pos + 2
                break

            if ((curr_pos + 3) < len(strs)):                                    
                poss_name = strs[curr_pos + 3].replace("\x00", "")
                if debug:
                    print "\nTry: '" + poss_name + "'"

                if (poss_name != "CompObj"):
                    name = poss_name
                    name_pos = curr_pos + 3
                    break

            if ((curr_pos + 4) < len(strs)):
                poss_name = strs[curr_pos + 4].replace("\x00", "")
                if debug:
                    print "\nTry: '" + poss_name + "'"

                if (poss_name != "ObjInfo"):
                    name = poss_name
                    name_pos = curr_pos + 4
                    break

            if ((curr_pos + 5) < len(strs)):
                poss_name = strs[curr_pos + 5].replace("\x00", "")
                if debug:
                    print "\nTry: '" + poss_name + "'"

                if (poss_name != "ObjInfo"):
                    name = poss_name
                    name_pos = curr_pos + 5
                    break

            curr_pos += 1

    return (name_pos, name)

def _get_raw_text_for_name(name_pos, strs, chunk, debug):
    """Use heuristics to get the potential text value for the object with
    the name at the given name position.

    @param name_pos (int) The position in the given list of strings
    where the name was found.

    @param strs (list) A list of strings pulled from the OLE chunk
    being analyzed.

    @param chunk (str) The OLE chunk being analyzed.

    @param debug (boolean) A flag indicating whether to print debug
    information.
    
    @return (str) The text associated with the object with the name at
    the given position. This will be an empty string if no associated
    text value is found.

    """

    text = ""
    asc_str = None
    if (name_pos + 1 < len(strs)):
        asc_str = strs[name_pos + 1].replace("\x00", "").strip()
    skip_names = set(["contents", "ObjInfo", "CompObj", None])
    if (("Calibr" not in asc_str) and
        ("OCXNAME" not in asc_str) and
        (asc_str not in skip_names) and
        (not asc_str.startswith("_DELETED_NAME_")) and
        (re.match(r"_\d{10}", asc_str) is None)):
        if debug:
            print "\nValue: 1"
            print strs[name_pos + 1]
                
        if (len(strs[name_pos + 1]) > 3):
            text = strs[name_pos + 1]
            if debug:
                print "\nValue: 2"
                print strs[name_pos + 1]

    val_pat = r"(?:\x00|\xff)[\x20-\x7e]+[^\x00]*\x00+\x02\x18"
    vals = re.findall(val_pat, chunk)
    if (len(vals) > 0):
        empty_pat = r"(?:\x00|\xff)#[^\x00]*\x00+\x02\x18"
        if (len(re.findall(empty_pat, vals[0])) == 0):
            poss_val = re.findall(r"[\x20-\x7e]+", vals[0][1:-2])[0]
            if ((poss_val != text) and (len(poss_val) > 1)):
                text += poss_val.replace("\x00", "")
                if debug:
                    print "\nValue: 3"
                    print poss_val.replace("\x00", "")

    val_pat = r"\x00#\x00\x00\x00[^\x02]+\x02"
    vals = re.findall(val_pat, chunk)
    if (len(vals) > 0):
        tmp_text = re.findall(r"[\x20-\x7e]+", vals[0][2:-2])
        if (len(tmp_text) > 0):
            poss_val = tmp_text[0]
            if (poss_val != text):
                if debug:
                    print "\nValue: 4"
                    print poss_val
                text += poss_val

    val_pat = r"([\x20-\x7e]{5,})\x00\x02\x0c\x00\x34"
    vals = re.findall(val_pat, chunk)
    if (len(vals) > 0):
        for v in vals:
            text += v
            if debug:
                print "\nValue: 5"
                print v

    val_pat = r"([\x20-\x7e]{5,})\x00{2,4}\x02\x0c"
    vals = re.findall(val_pat, chunk)
    if (len(vals) > 0):
        for v in vals:
            text += v
            if debug:
                print "\nValue: 6"
                print v
                
    for pos in range(name_pos + 2, len(strs)):
        curr_str = strs[pos].replace("\x00", "")
        if ((len(curr_str) > 40) and (not curr_str.startswith("Microsoft "))):
            text += curr_str

    return text

def _clean_text_for_name(chunk, name, text, object_names, stream_names, longest_str, orig_strs, debug):
    """Clean up the text value associated with an object with a given
    name.

    @param chunk (str) The OLE chunk being analyzed.

    @param name (str) The name of the object.

    @param text (str) The raw text associated with the object.

    @param object_names (list) The names of objects referenced in the
    VBA code of the Office file being analyzed.

    @param stream_names (list) The names of the OLE streams in the
    Office file being analyzed.

    @param longest_str (str) The longest string associated with an
    object (so far).

    @param orig_strs (list) The ASCII strings pulled from the OLE
    chunk.

    @param debug (boolean) A flag indicating whether to print debug
    information.

    @return (str) The cleaned up text value.

    """

    size_pat = r"\x48\x80\x2c\x03\x01\x02\x00(.{2})"
    tmp = re.findall(size_pat, chunk)
    if (len(tmp) == 0):
        size_pat = r"\x48\x80\x2c(.{2})"
        tmp = re.findall(size_pat, chunk)
    if (len(tmp) == 0):
        size_pat = r"\xf8\x00\x28\x00\x00\x00(.{2})"
        tmp = re.findall(size_pat, chunk)
    if (len(tmp) == 0):
        size_pat = r"\x2c\x00\x00\x00\x1d\x00\x00\x00(.{2})"
        tmp = re.findall(size_pat, chunk)
    if (len(tmp) > 0):
        size_bytes = tmp[0]
        size = ord(size_bytes[1]) * 256 + ord(size_bytes[0])
        if (debug):
            print "SIZE: "
            print size
        if ((len(text) > size) and (not name.startswith("Page"))):
            text = text[:size]

    if ((strip_name(text) in object_names) or
        (strip_name(text) in stream_names)):
        if debug:
            print "\nBAD: Val is name '" + text + "'"

        if ((text.startswith("Page")) and (len(longest_str) > 30)):
            tmp_str = ""
            for field in orig_strs:
                if ((len(field) > 20) and
                    (not field.startswith("Microsoft "))):
                    tmp_field = ""
                    for s in re.findall(r"[\x20-\x7f]{5,}", field):
                        tmp_field += s
                    tmp_str += tmp_field
            text = tmp_str
        else:
            text = ""
        if debug:
            print len(longest_str)
            print "BAD: Set Val to '" + text + "'"

    text = text.replace("\x00", "")
    if (len(re.findall(r"[^\x20-\x7f]", text)) > 2):
        if debug:
            print "\nBAD: Binary in Val. Set to ''"
        text = ""

    if ((text.startswith("Forms.")) and (len(text) < 20)):
        text = ""

    return text

def _find_longest_strs_form_results(long_strs, r):
    """Find various longest strings from a general list of extracted
    strings and text values assigned to object names.

    @param long_strs (list) A list of pretty long strings encountered
    during processing.

    @param r (list) A list of 2 element tuples where the 1st tuple
    element is the name of an object and the 2nd element is the object's
    associated text value.

    @return (tuple) A 3 element tuple where the 1st element is the
    longest string found in the longish string list, the 2nd element
    is the longest text value associated with an object name, and the
    3rd element is the longest text value associated with a PageNN
    object.

    """

    longest_val = ""
    longest_str = ""
    page_val = ""
    for s in long_strs:
        if (len(s) > len(longest_str)):
            longest_str = s
    
    for pair in r:
        name = pair[0]
        val = pair[1]
        if (name.startswith("Page")):
            if (len(val) > len(page_val)):
                page_val = val
        if (name != "Page1"):
            continue
        if (len(val) > len(longest_val)):
            longest_val = val

    return (longest_str, longest_val, page_val)

def _merge_ole_form_results(r, v1_vals, v1_1_vals):
    """Merge the results of various heuristic methods used to find the
    text values of OLE objects.

    @param r (list) Value results as a list of 2 element tuples where
    the 1st element is the name (str) of an object and the 2nd element
    is the text value (str) of the object.

    @param v1_vals (list) Value results as a list of 2 element tuples
    where the 1st element is the name (str) of an object and the 2nd
    element is the text value (str) of the object.

    @param v1_1_vals (list) Value results as a list of 2 element
    tuples where the 1st element is the name (str) of an object and
    the 2nd element is the text value (str) of the object.

    @return (list) The merged results as a list of 2 element tuples
    where the 1st element is the name (str) of an object and the 2nd
    element is the text value (str) of the object.

    """

    tmp = []
    v2_vals = r
    for v1_pair in v1_vals:
        tmp.append(v1_pair)
        for v2_pair in v2_vals:
            if (v1_pair[0] != v2_pair[0]):
                tmp.append(v2_pair)
    r = tmp
    if (len(r) == 0):
        r = v2_vals

    tmp = []
    v2_vals = r
    for v1_pair in v1_1_vals:
        tmp.append(v1_pair)
        for v2_pair in v2_vals:
            if (v1_pair[0] != v2_pair[0]):
                tmp.append(v2_pair)
    r = tmp
    if (len(r) == 0):
        r = v2_vals

    tmp = []
    for old_pair in r:
        name = old_pair[0]
        val = old_pair[1]
        for cruft_pat in cruft_pats:
            val = re.sub(cruft_pat, "", val)
        tmp.append((name, val))
    r = tmp

    return r

def _clean_up_ole_form_results(r, long_strs, v1_vals, v1_1_vals, object_names, debug):
    """Clean up the object name/value results computed in various ways,
    merge the various results, and return the merged and cleaned
    results.
    
    @param r (list) Value results as a list of 2 element tuples where
    the 1st element is the name (str) of an object and the 2nd element
    is the text value (str) of the object.

    @param long_strs (list) A list of pretty long strings encountered
    during processing.

    @param v1_vals (list) Value results as a list of 2 element tuples
    where the 1st element is the name (str) of an object and the 2nd
    element is the text value (str) of the object.

    @param v1_1_vals (list) Value results as a list of 2 element
    tuples where the 1st element is the name (str) of an object and
    the 2nd element is the text value (str) of the object.

    @param object_names (list) A list of the names (str) of the object fields
    referenced in the VBA code.

    @param debug (boolean) A flag indicating whether to print debug
    information.

    """

    last_val = None
    tmp = []
    for dat in r:

        if (dat[0].strip() != last_val):
            tmp.append(dat)
        else:
            if debug:
                print "\nSkip 1: " + str(dat)
        last_val = dat[1].strip()
    r = tmp

    if debug:
        print "\nFirst result:"
        print r
    
    tmp = []
    last_var = None
    last_val = None
    for dat in r:

        if (len(dat[0]) > 50):

            last_val = dat[0]

        if (last_var is not None):
            tmp.append((last_var, last_val))

        last_var = dat[0]
        last_val = dat[1]

    if ((last_var is not None) and (len(last_var) < 50)):
        tmp.append((last_var, last_val))
    r = tmp

    tmp = []
    pos = -1
    last_val = ""
    if debug:
        print "\nLONG STRS!!"
        print long_strs
    for dat in r:

        pos += 1
        curr_var = dat[0]
        curr_val = dat[1]        
        if debug:
            print curr_var
            print pos
            print len(curr_val)
        if ((curr_val is None) or (len(curr_val) == 0)):
            
            replaced = False
            for i in range(pos + 1, len(r)):
                poss_val = long_strs[i]
                if (len(r[i][1]) > len(poss_val)):
                    poss_val = r[i][1]
                if (len(poss_val) > 15):
                    if debug:
                        print "\nREPLACE (1)"
                    curr_val = poss_val
                    replaced = True
                    break

            if ((not replaced) and (len(last_val) > 15)):
                if debug:
                    print "\nREPLACE (2)"
                curr_val = last_val

        tmp.append((curr_var, curr_val))
        last_val = curr_val
    r = tmp

    r = _merge_ole_form_results(r, v1_vals, v1_1_vals)
    
    longest_str, longest_val, page_val = _find_longest_strs_form_results(long_strs, r)

    
    page_names = set()
    if (longest_val != ""):
        tmp_r = []
        updated_page1 = False
        for pair in r:
            name = pair[0]
            if (name == "Page2"):
                tmp_r.append((name, longest_str))
                continue
            if (name != "Page1"):
                tmp_r.append(pair)
                continue
            if (not updated_page1):
                tmp_r.append((name, longest_val))
                updated_page1 = True
        r = tmp_r

    if debug:
        print "\nPAGE VAL!!"
        print page_val
    if (page_val == ""):
        page_val = longest_str
        
    for i in range(1, 5):
        curr_name = "Page" + str(i)
        if ((curr_name not in page_names) and (page_val != "")):
            r.append((curr_name, page_val))

    handled_names = set()
    for mapping in r:
        handled_names.add(mapping[0])
    for curr_name in object_names:
        if ((curr_name not in handled_names) and (longest_str != "")):
            r.append((curr_name, longest_str))

    return r
            
def get_ole_textbox_values(obj, vba_code):
    """Read in the text associated with embedded OLE form textbox
    objects. NOTE: This currently is a NASTY hack.

    @param obj (str) The read in Office file to analyze or the name
    of the Office file to analyze. The file will be read in if a file
    name is given.

    @param vba_code (str) The VBA macro code from the Office file.

    @return (list) The results as a list of 2 element tuples where the
    1st element is the name (str) of an object and the 2nd element is
    the text value (str) of the object.

    """

    if obj[0:4] == '\xd0\xcf\x11\xe0':

        data = obj
    else:

        try:
            f = open(obj, "rb")
            data = f.read()
            f.close()
        except IOError:
            data = obj
        except TypeError:
            data = obj

    if (not filetype.is_office97_file(data, True)):

        data = get_vbaprojectbin(data)
        if (data is None):
            return []

    debug = False
    if debug:
        print "\nExtracting OLE/ActiveX TextBox strings..."
        
    stream_names = _get_stream_names(vba_code)
    if debug:
        print "\nStream Names: " + str(stream_names) + "\n"
        
    data = data.replace("R\x00o\x00o\x00t\x00 \x00E\x00n\x00t\x00r\x00y", "")
    data = data.replace("o" + "\x00" * 40, "\x00" * 40)
    data = re.sub("Tahoma\w{0,5}", "\x00", data)

    r = get_ole_text_method_1(vba_code, data)
    if (r is not None):
        return r
    
    v1_vals = get_ole_textbox_values1(data, debug, stream_names)

    v1_1_vals = get_ole_textbox_values2(data, debug, vba_code, stream_names)

    if debug:
        print "\nget_ole_textbox_values()\n"

    object_names, page_names = _pull_object_names(vba_code)
    if debug:
        print "\nNames from VBA code:"
        print object_names
            
    if (data is None):
        if debug:
            print "\nNO DATA"
            sys.exit(0)
        return []

    data = data.replace("c\x00o\x00n\x00t\x00e\x00n\x00t\x00s", "\x00c\x00o\x00n\x00t\x00e\x00n\x00t\x00s\x00")
    data = re.sub("(_(?:\x00\d){10})", "\x00" + r"\1", data)

    page_name_pat = r"Page(\d+)(?:(?:\-\d+)|[a-zA-Z\.]+[a-zA-Z0-9]*)"
    data = re.sub(page_name_pat, r"Page\1", data)
    
    form_str = "Microsoft Forms 2.0"
    form_str_pat = r"Microsoft Forms 2.0 [A-Za-z]{2,30}(?!Form)"
    field_marker = "Forms."
    if (re.search(form_str_pat, data) is None):
        if debug:
            print "\nNO FORMS"
            sys.exit(0)
        return []

    pat = r"(?:(?:[\x20-\x7e]|\r?\n){3,})|(?:(?:(?:\x00|\xff)(?:[\x20-\x7e]|\r?\n)){3,})"
    index = 0
    r = []
    found_names = set()
    long_strs = []
    end_object_marker = "D\x00o\x00c\x00u\x00m\x00e\x00n\x00t\x00S\x00u\x00m\x00m\x00a\x00r\x00y\x00I\x00n\x00f\x00o\x00r\x00m\x00a\x00t\x00i\x00o\x00n"
    while (re.search(form_str_pat, data[index:]) is not None):


        chunk, index, end = _get_next_chunk(data, index, form_str, form_str_pat, end_object_marker)

        strs = re.findall(pat, chunk)
        if debug:
            print "\n\n-------------- CHUNK ---------------"
            print chunk
            print str(strs).replace("\\x00", "").replace("\\xff", "")

        longest_str = ""
        orig_strs = strs
        for field in strs:
            if ((len(field) > 30) and
                (len(field) > len(longest_str)) and
                (not field.startswith("Microsoft "))):
                longest_str = field
        long_strs.append(longest_str)

        curr_pos, name_pos, name = _find_name_in_data(page_names, found_names, strs, debug)

        if (name is None):

            curr_pos, name_pos, name = _find_name_in_data(object_names, found_names, strs, debug)

        if (name is None):        
            name_pos, name = _guess_name_from_data(strs, field_marker, debug)
            
        if (not is_name(name)):
            index = end
            if debug:
                print "\nNo name found. Moving to next chunk."
            r.append(("no name found", "placeholder"))
            continue

        name = strip_name(name)
        if debug:
            print "\nPossible Name: '" + name + "'"
        
        text = _get_raw_text_for_name(name_pos, strs, chunk, debug)
        if debug:
            print "\nORIG:"
            print name
            print text
            print len(text)

        text = _clean_text_for_name(chunk, name, text, object_names, stream_names, longest_str, orig_strs, debug)
                    
        if ((text != "") or (not name.startswith("Page"))):
            if debug:
                print "\nSET '" + name + "' = '" + text + "'"
            r.append((name, text))

        if (text != ""):
            found_names.add(name)

        index = end

    r = _clean_up_ole_form_results(r, long_strs, v1_vals, v1_1_vals, object_names, debug)
                
    if debug:
        print "\nFINAL RESULTS:" 
        print r
        sys.exit(0)
        
    return r

def read_form_strings(vba):
    """Read in the form strings in order as a lists of tuples like
    (stream name, form string).

    @param vba (str) The VBA code to analyze, generated with
    olevba. Note that olevba includes the form strings in the output.

    @return (list) A list of 2 element tuples where the 1st element is
    the name of the stream holding the form and the 2nd element is the
    form text.

    """

    try:
        r = []
        skip_strings = ["Tahoma", "Tahomaz"]
        for (_, stream_path, form_string) in vba.extract_form_strings():

            if (form_string in skip_strings):
                continue
            if (not all((ord(c) > 31 and ord(c) < 127) for c in form_string)):
                continue

            stream_name = stream_path.replace("Macros/", "")
            if ("/" in stream_name):
                stream_name = stream_name[:stream_name.index("/")]

            r.append((stream_name, form_string))

        return r

    except Exception as e:
        log.error("Cannot read form strings. " + str(e))
        return []
    
def get_shapes_text_values_xml(fname):
    """Read in the text associated with Shape objects in a document saved
    as Flat OPC XML files. NOTE: This currently is a hack.

    @param fname (str) The OPC XML file contents (already read in) or
    the name of the file to analyze. If a file name is given it will
    be read in.

    @return (list) The results as a list of 2 element tuples where the
    1st element is the name (str) of an object and the 2nd element is
    the text value (str) of the object.

    """

    contents = None
    if fname.startswith("<?xml"):
        contents=fname
    else:

        try:
            f = open(fname, "r")
            contents = f.read().strip()
            f.close()
        except IOError:
            contents = fname
        except TypeError:
            contents = fname

    if ((not contents.startswith("<?xml")) or
        ("<w:txbxContent>" not in contents)):
        return []

    log.warning("Looking for Shapes() strings in Flat OPC XML file...")

    blocks = []
    start = contents.index("<w:txbxContent>") + len("<w:txbxContent>")
    end = contents.index("</w:txbxContent>")
    while (start is not None):
        blocks.append(contents[start:end])
        if ("<w:txbxContent>" in contents[end:]):
            start = end + contents[end:].index("<w:txbxContent>") + len("<w:txbxContent>")
            end = end + len("</w:txbxContent>") + contents[end + len("</w:txbxContent>"):].index("</w:txbxContent>")
        else:
            start = None
            end = None
            break
    cmd_strs = []
    for block in blocks:

        pat = r"\<w\:t[^\>]*\>([^\<]+)\</w\:t\>"
        strs = re.findall(pat, block)

        if (len(strs) > 1):

            curr_str = ""
            for s in strs:

                curr_str += s

            strs = [curr_str]

        cmd_strs.append(strs[0])
            
    r = []
    pos = 1
    for shape_text in cmd_strs:

        if (len(shape_text) < 100):
            continue
        
        shape_text = shape_text.replace("&amp;", "&")
        var = "Shapes('" + str(pos) + "').TextFrame.TextRange.Text"
        r.append((var, shape_text))
        
        var = "Shapes('" + str(pos) + "').TextFrame.ContainingRange"
        r.append((var, shape_text))

        var = "Shapes('" + str(pos) + "').AlternativeText"
        r.append((var, shape_text))
        
        pos += 1

    return r

def get_shapes_text_values_direct_2007(data):
    """Read in shapes name/value mappings directly from word/document.xml
    from an unzipped Word 2007+ file.

    @param data (str) The contents of the document.xml file to
    analyze.

    @return (list) The results as a list of 2 element tuples where the
    1st element is the name (str) of an object and the 2nd element is
    the text value (str) of the object.

    """

    
    pat1 = r'<v:shape\s+id="(\w+)".+<w:txbxContent>'
    name = re.findall(pat1, data)
    if (len(name) == 0):
        return []
    name = name[0]

    pat2 = r'<w:t[^<]*>([^<]+)</w:t[^<]*>'
    vals = re.findall(pat2, data)
    if (len(vals) == 0):
        return []

    val = ""
    for v in vals:
        val += v
    val = _clean_2007_text(val)
    
    r = [(name, val)]
    return r

def get_shapes_text_values_direct_2007_1(data):
    """Read in shapes name/value mappings directly from word/document.xml
    from an unzipped Word 2007+ file another way.

    @param data (str) The contents of the document.xml file to
    analyze.

    @return (list) The results as a list of 2 element tuples where the
    1st element is the name (str) of an object and the 2nd element is
    the text value (str) of the object.

    """

    
    pat1 = r'<wp\:docPr +id="(\d+)" +name="[^"]*" +descr="([^"]*)"'
    shape_info = re.findall(pat1, data)
    if (len(shape_info) == 0):
        return []
    shape_info = shape_info[0]
    name = shape_info[0]
    val = _clean_2007_text(shape_info[1])
        
    r = [(name, val)]
    return r

def _parse_activex_chunk(data):
    """Parse out ActiveX text values from 2007+ activeXN.bin file
    contents.

    @param data (str) The contents of the activeXN.bin to analyze
    (already read in).

    @return (str) The ActiveX text value if found, None if not found.

    """

    anchor = None
    pad = 0
    if (b"\x1a\x00\x00\x00\x23" in data):
        anchor = b"\x1a\x00\x00\x00\x23"
        pad = 3
    elif (b"\x05\x00\x00\x00\x01\x00\x00\x80" in data):
        anchor = b"\x05\x00\x00\x00\x01\x00\x00\x80"
        pad = 16
    elif (b"\x30\x01\x00\x00" in data):
        anchor = b"\x30\x01\x00\x00"
    if (anchor is None):
        return None
    start = data.rindex(anchor) + len(anchor) + pad
    pat = r"([\x20-\x7e]+)"
    text = re.findall(pat, data[start:])
    if (len(text) == 0):
        return None
    text = text[0]

    size_pat = r"\x48\x80\x2c\x03\x01\x02\x00(.{2})"
    tmp = re.findall(size_pat, data)
    if (len(tmp) == 0):
        size_pat = r"\x48\x80\x2c(.{2})"
        tmp = re.findall(size_pat, data)
    if (len(tmp) == 0):
        size_pat = r"\x00\x01\x00\x00\x80(.{2})"
        tmp = re.findall(size_pat, data)
    if (len(tmp) > 0):
        size_bytes = tmp[0]
        size = ord(size_bytes[1]) * 256 + ord(size_bytes[0])
        if (len(text) > size):
            text = text[:size]
        

    return text

def _parse_activex_rich_edit(data):
    """Parse out Rich Edit control text values from 2007+ activeXN.bin
    file contents.

    @param data (str) The contents of the activeXN.bin to analyze
    (already read in).

    @return (str) The ActiveX text value if found, None if not found.

    """

    data = data.replace("\x00", "")

    pat = r"\\fs\d{1,4} (.+)\\par"
    val = re.findall(pat, data)
    if (len(val) == 0):
        return None
    return _clean_2007_text(val[0])

def _get_comments_docprops_2007(unzipped_data):
    """
    Read in the comments in a document saved in the 2007+ format.
    Gets comments from docProps/core.xml.
    """

    zip_subfile = 'docProps/core.xml'
    if (zip_subfile not in unzipped_data.namelist()):
        zip_subfile = 'docProps\\core.xml'
        if (zip_subfile not in unzipped_data.namelist()):
            return []

    f1 = unzipped_data.open(zip_subfile)
    data = f1.read()
    f1.close()

    comm_pat = r"<dc:description>(.*)</dc:description>"
    comment_blocks = re.findall(comm_pat, data, re.DOTALL)
    if (len(comment_blocks) == 0):
        return []

    pos = 1
    r = []
    for text in comment_blocks:
        r.append((pos, _clean_2007_text(text)))
        pos += 1

    return r
        
def _get_comments_2007(fname):
    """Read in the comments in a document saved in the 2007+ format.
    Gets comments from word/comments.xml.

    @param fname (str) The name of the Office 2007+ file to analyze.

    @return (list) A list of 2 element tuples where the 1st tuple
    element is the ID of the comment and the 2nd element is the
    comment text.

    """
        
    unzipped_data, fname = unzip_data(fname)
    delete_file = (fname is not None)
    if (unzipped_data is None):
        return []

    zip_subfile = 'word/comments.xml'
    if (zip_subfile not in unzipped_data.namelist()):
        zip_subfile = 'word\\comments.xml'
        if (zip_subfile not in unzipped_data.namelist()):

            r = _get_comments_docprops_2007(unzipped_data)
            unzipped_data.close()
            if (delete_file):
                os.remove(fname)
            return r

    r = []
    f1 = unzipped_data.open(zip_subfile)
    data = f1.read()
    f1.close()


    comm_pat = r"<w:comment.*</w:comment>"
    comment_blocks = re.findall(comm_pat, data)
    if (len(comment_blocks) == 0):
        unzipped_data.close()
        if (delete_file):
            os.remove(fname)
        return []

    r = []
    for block in comment_blocks:

        id_pat = r"<w:comment\s+w:id=\"(\d+)\""
        ids = re.findall(id_pat, block)
        if (len(ids) == 0):
            continue
        curr_id = ids[0]

        text_pat = r"<w:t[^>]*>([^<]+)</w:t>"
        texts = re.findall(text_pat, block)
        if (len(texts) == 0):
            continue

        block_text = ""

        for text in texts:
            block_text += _clean_2007_text(text)

        r.append((curr_id, block_text))
        
    unzipped_data.close()
    if (delete_file):
        os.remove(fname)
    return r

def get_comments(fname):
    """Read the comments from an Office file.

    @param fname (str) The name of the Office file to analyze.

    @return (list) A list of 2 element tuples where the 1st tuple
    element is the ID of the comment and the 2nd element is the
    comment text.

    """

    if (not filetype.is_office2007_file(fname, (len(fname) > 2000))):
        return []

    return _get_comments_2007(fname)

def get_shapes_text_values_2007(fname):
    """Read in the text associated with Shape objects in a document saved
    in the 2007+ format.

    @param fname (str) The name of the Office 2007+ file to analyze.

    @return (list) The results as a list of 2 element tuples where the
    1st element is the name (str) of an object and the 2nd element is
    the text value (str) of the object.

    """
        
    unzipped_data, fname = unzip_data(fname)
    delete_file = (fname is not None)
    if (unzipped_data is None):
        return []

    zip_subfile = 'word/document.xml'
    if (zip_subfile not in unzipped_data.namelist()):
        zip_subfile = 'word\\document.xml'
        if (zip_subfile not in unzipped_data.namelist()):
            if (delete_file):
                os.remove(fname)
            return []

    r = []
    f1 = unzipped_data.open(zip_subfile)
    data = f1.read()
    f1.close()

    r = get_shapes_text_values_direct_2007(data)
    if (len(r) > 0):
        return r
    r = get_shapes_text_values_direct_2007_1(data)
    if (len(r) > 0):
        return r
    
    pat = r'<w\:control[^>]+r\:id="(\w+)"[^>]+w\:name="(\w+)"'
    var_info = re.findall(pat, data)
    id_name_map = {}
    for shape in var_info:
        id_name_map[shape[0]] = shape[1]

    zip_subfile = 'word/_rels/document.xml.rels'
    if (zip_subfile not in unzipped_data.namelist()):
        zip_subfile = 'word\\_rels\\document.xml.rels'
        if (zip_subfile not in unzipped_data.namelist()):
            if (delete_file):
                os.remove(fname)
            return []

    r = []
    f1 = unzipped_data.open(zip_subfile)
    data = f1.read()
    f1.close()

    pat = r'<Relationship[^>]+Id="(\w+)"[^>]+Target="([^"]+)"'
    var_info = re.findall(pat, data)
    id_activex_map = {}
    for shape in var_info:
        if (shape[0] not in id_name_map):
            continue
        id_activex_map[shape[0]] = shape[1].replace(".xml", ".bin")

    for shape in id_activex_map:

        path = "word/" + id_activex_map[shape]
        if (path not in unzipped_data.namelist()):
            path = "word\\" + id_activex_map[shape].replace("/", "\\")
            if (path not in unzipped_data.namelist()):
                continue

        f1 = unzipped_data.open(path)
        data = f1.read()
        f1.close()

        text = _parse_activex_chunk(data)

        if (text is None):
            text = _parse_activex_rich_edit(data)
        if (text is None):
            continue
            
        r.append((id_name_map[shape], _clean_2007_text(text)))
    
    unzipped_data.close()
    if (delete_file):
        os.remove(fname)
    return r

def get_shapes_text_values(fname, stream):
    """Read in the text associated with Shape objects in the
    document. NOTE: This currently is a hack.

    @param fname (str) The name of the Office file to analyze.

    @return (list) The results as a list of 2 element tuples where the
    1st element is the name (str) of an object and the 2nd element is
    the text value (str) of the object.

    """

    r = get_shapes_text_values_2007(fname)
    if (len(r) > 0):
        return r
    
    r = []
    try:
        ole = olefile.OleFileIO(fname, write_mode=False)
        if (not ole.exists(stream)):
            return []
        data = ole.openstream(stream).read()
        
        pat = r"\x0d[\x20-\x7e]{100,}\x0d"
        strs = re.findall(pat, data)
        
        pos = 1
        for shape_text in strs:

            shape_text = shape_text[1:-1]
            var = "Shapes('" + str(pos) + "').TextFrame.TextRange.Text"
            r.append((var, shape_text))
            
            var = "Shapes('" + str(pos) + "').TextFrame.ContainingRange"
            r.append((var, shape_text))

            var = "Shapes('" + str(pos) + "').AlternativeText"
            r.append((var, shape_text))
            
            pos += 1

        pat = r"(?:\x00[\x20-\x7e]){100,}"
        strs = re.findall(pat, data)
        
        pos = 1
        for shape_text in strs:

            shape_text = shape_text[1:-1].replace("\x00", "")
            var = "Shapes('" + str(pos) + "').TextFrame.TextRange.Text"
            r.append((var, shape_text))
            
            var = "Shapes('" + str(pos) + "').TextFrame.ContainingRange"
            r.append((var, shape_text))

            var = "Shapes('" + str(pos) + "').AlternativeText"
            r.append((var, shape_text))
            
            pos += 1
            
    except Exception as e:

        if ("not an OLE2 structured storage file" not in str(e)):
            log.error("Cannot read associated Shapes text. " + str(e))

        if ("not an OLE2 structured storage file" in str(e)):
            r = get_shapes_text_values_xml(fname)

    return r


URL_REGEX = r'(http[s]?://(?:(?:[a-zA-Z0-9_\-]+\.[a-zA-Z0-9_\-\.]+(?::[0-9]+)?)+(?:/[/\?&\~=a-zA-Z0-9_\-\.]+)))'
def pull_urls_from_comments(vba):
    """Pull out URLs that just appear in VBA comments.

    @param vba (VBA_Parser object) The olevba VBA_Parser object for
    reading the Office file being analyzed.

    @return (set) URLs (str) that just appear in VBA comment
    statements.

    """

    macros = ""
    for (_, _, _, vba_code) in vba.extract_macros():
        if (vba_code is None):
            continue
        macros += vba_code + "\n"

    urls = set()
    for line in macros.split("\n"):
        line = line.strip()
        if ((not line.startswith("'")) and (not line.lower().startswith("rem "))):
            continue
        for url in re.findall(URL_REGEX, line):
            urls.add(url.strip())

    return urls

def pull_urls_office97(fname, is_data, vba):
    """Pull URLs directly from an Office97 file.

    @param fname (str) The name of the file from which to scrape
    URLs or the raw file contents.

    @param is_data (boolean) A flag indicating whether fname is a file
    name (False) or the raw file contents (True).

    @param vba (str) The decompressed VBA macro code.

    @return (set) The URLs scraped from the file. This will be empty
    if there are no URLs.

    """

    if (not filetype.is_office97_file(fname, is_data)):
        return []
    
    data = None
    if (not is_data):
        with open(fname, 'rb') as f:
            data = f.read()
    else:
        data = fname

    comment_urls = set()
    if (vba is not None):
        comment_urls = pull_urls_from_comments(vba)
    file_urls = re.findall(URL_REGEX, data)
    r = set()
    for url in file_urls:
        url = url.strip()
        not_comment_url = True
        for comment_url in comment_urls:
            if ((url.startswith(comment_url)) or (comment_url.startswith(url))):
                not_comment_url = False
                break
        if (not_comment_url):
            r.add(url)
        
    return r

def _read_doc_vars_zip(fname):
    """Read doc vars from an Office 2007+ file.

    @param fname (str) The name of the Office file to analyze.

    @return (list) A list of 2 element tuples where the 1st element is
    the document variable name and the 2nd element is the value.

    """

    f = zipfile.ZipFile(fname, 'r')

    if ('word/settings.xml' not in f.namelist()):
        return []

    f1 = f.open('word/settings.xml')
    data = f1.read()
    f1.close()
    f.close()

    pat = r'<w\:docVar w\:name="(\w+)" w:val="([^"]*)"'
    var_info = re.findall(pat, data)

    r = []
    for i in var_info:
        val = i[1]
        val = val.replace("&quot;", '"')
        val = val.replace("&amp;", '&')
        val = val.replace("&lt;", '<')
        val = val.replace("&gt;", '>')
        r.append((i[0], val))
    
    return r
    
def _read_doc_vars_ole(fname):
    """Use a heuristic to try to read in document variable names and
    values from the 1Table OLE stream. Note that this heuristic is
    kind of hacky and is not close to being a general solution for
    reading in document variables, but it serves the need for
    SimulationVBA emulation.

    TODO: Replace this when actual support for reading doc vars is
    added to olefile.

    @param fname (str) The name of the Office file to analyze.

    @return (list) A list of 2 element tuples where the 1st element is
    the document variable name and the 2nd element is the value.

    """

    try:

        ole = olefile.OleFileIO(fname, write_mode=False)
        var_offset, var_size = _get_doc_var_info(ole)
        if ((var_offset is None) or (var_size is None) or (var_size == 0)):
            return []
        data = ole.openstream("1Table").read()[var_offset : (var_offset + var_size + 1)]
        tmp_strs = re.findall("(([^\x00-\x1F\x7F-\xFF]\x00){2,})", data)
        strs = []
        for s in tmp_strs:
            s1 = s[0].replace("\x00", "").strip()
            strs.append(s1)
            
        pos = 0
        r = []
        end = len(strs)
        if (end % 2 != 0):
            end = end + 1
            strs.append("Unknown")
        end = end/2
        while (pos < end):
            r.append((strs[pos], strs[pos + end]))
            pos += 1

        return r
            
    except Exception as e:
        log.error("Cannot read document variables. " + str(e))
        return []

def _read_doc_vars(data, fname):
    """Read document variables from Office 97 or 2007+ files.

    @param data (str) The read in Office file data. Can be None if data
    should be read from a file (fname).

    @param fname (str) The name of the Office file to analyze. Can be
    None if data is given (data).

    @return (list) A list of 2 element tuples where the 1st element is
    the document variable name and the 2nd element is the value.

    """
    if ((fname is None) or (len(fname) < 1)):
        obj = io.BytesIO(data)
    else:
        obj = fname
    r = []
    if olefile.isOleFile(obj):
        r = _read_doc_vars_ole(obj)
    elif zipfile.is_zipfile(obj):
        r = _read_doc_vars_zip(obj)
    return r

def _get_inlineshapes_text_values(data):
    """Read in the text associated with InlineShape objects in the
    document. NOTE: This currently is a hack.

    @param data (str) The read in Office file (data).

    @return (list) The results as a list of 2 element tuples where the
    1st element is the name (str) of an object and the 2nd element is
    the text value (str) of the object.

    """

    r = []
    try:

        pat = r"\x00p\x00i\x00x\x00e\x00l\x00*((?:\x00?[\x20-\x7e])+)\x00\x00\x00"
        strs = re.findall(pat, data)

        pos = 1
        for shape_text in strs:

            shape_text = shape_text.replace("\x00", "")
            var = "InlineShapes('" + str(pos) + "').TextFrame.TextRange.Text"
            r.append((var, shape_text))
            
            var = "InlineShapes('" + str(pos) + "').TextFrame.ContainingRange"
            r.append((var, shape_text))

            var = "InlineShapes('" + str(pos) + "').AlternativeText"
            r.append((var, shape_text))
            var = "InlineShapes('" + str(pos) + "').AlternativeText$"
            r.append((var, shape_text))
            
            pos += 1
            
    except Exception as e:

        log.error("Cannot read associated InlineShapes text. " + str(e))

        if ("not an OLE2 structured storage file" in str(e)):
            r = get_shapes_text_values_xml(data)

    return r

def _read_custom_doc_props(fname):
    """Use a heuristic to try to read in custom document property names
    and values from the DocumentSummaryInformation OLE stream. Note
    that this heuristic is kind of hacky and is not close to being a
    general solution for reading in document properties, but it serves
    the need for SimulationVBA emulation.

    TODO: Replace this when actual support for reading doc properties
    is added to olefile.

    @param fname (str) The name of the Office file to analyze.

    @return (list) A list of 2 element tuples where the 1st element is
    the document property name and the 2nd element is the value.

    """

    try:

        ole = olefile.OleFileIO(fname, write_mode=False)
        data = None
        for stream_name in ole.listdir():
            if ("DocumentSummaryInformation" in stream_name[-1]):
                data = ole.openstream(stream_name).read()
                break
        if (data is None):
            return []
        strs = re.findall("([\w\.\:/]{4,})", data)
        

        skip_names = set(["Title"])
        tmp = []
        for s in strs:
            if (s not in skip_names):
                tmp.append(s)
        strs = tmp

        if (len(strs) == 1):
            strs = ["*", strs[0]]

        pos = 0
        r = []
        for s in strs:
            if ((pos + 1) < len(strs)):
                r.append((s, strs[pos + 1]))
            pos += 1

        return r
            
    except Exception as e:
        if ("not an OLE2 structured storage file" not in str(e)):
            log.error("Cannot read custom doc properties. " + str(e))
        return []

def _get_embedded_object_values(fname):
    """Read in the tag and caption associated with Embedded Objects in
    the document.  NOTE: This currently is a hack.

    @param fname (str) The name of the Office file to analyze.

    @return (list) List of tuples of the form (var name, caption
    value, tag value)

    """

    r = []
    try:

        ole = olefile.OleFileIO(fname, write_mode=False)
        
        ole_dirs = ole.listdir()
        for dir_info in ole_dirs:

            curr_dir = ""
            first = True
            for d in dir_info:
                if (not first):
                    curr_dir += "/"
                first = False
                curr_dir += d
            data = ole.openstream(curr_dir).read()


            pat =  r"Begin \{[A-Z0-9\-]{36}\} (\w{1,50})\s*(?:\r?\n)\s{1,10}" + \
                   r"Caption\s+\=\s+\"(\w+)\"[\w\s\='\n\r]+Tag\s+\=\s+\"(.+)\"[\w\s\='\n\r]+End"
            obj_text = re.findall(pat, data)

            for i in obj_text:
                r.append(i)
        
    except Exception as e:
        if ("not an OLE2 structured storage file" not in str(e)):
            log.error("Cannot read tag/caption from embedded objects. " + str(e))

    return r

def _read_doc_text_libreoffice(data):
    """Read in the document text and tables from a Word file (already
    read in) using LibreOffice.

    @param data (str) The read in Office file (data).

    @return (tuple) Returns a tuple containing the doc text and a list
    of tuples containing dumped tables.

    """
    
    if (not filetype.is_office_file(data, True)):
        log.warning("The file is not an Office file. Not extracting document text with LibreOffice.")
        return None
    
    out_dir = None
    while True:
        out_dir = "/tmp/tmp_word_file_" + str(random.randrange(0, 10000000000))
        try:
            f = open(out_dir, "r")
            f.close()
        except IOError:
            break

    f = open(out_dir, 'wb')
    f.write(data)
    f.close()
    
    output = None
    try:
        output = subprocess.check_output(["timeout", "30", "python3", _thismodule_dir + "/../export_doc_text.py",
                                          "--text", "-f", out_dir])
    except Exception as e:
        log.error("Running export_doc_text.py failed. " + str(e))
        os.remove(out_dir)
        return None

    r = []
    for line in output.split("\n"):
        r.append(line)

    if (len(r) > 0):

        first_line = r[0]
        good_pos = 0
        while ((good_pos < 10) and (good_pos < len(first_line))):
            if (first_line[good_pos] in string.printable):
                break
            good_pos += 1
        first_line = first_line[good_pos:]
                
        pat = r'^\*.*\*\/'
        if (re.match(pat, first_line) is not None):
            first_line = "/" + first_line
        if (first_line.startswith("[]*")):
            first_line = "/*" + first_line
        r = [first_line] + r[1:]

    output = None
    try:
        output = subprocess.check_output(["python3", _thismodule_dir + "/../export_doc_text.py",
                                          "--tables", "-f", out_dir])
    except Exception as e:
        log.error("Running export_doc_text.py failed. " + str(e))
        os.remove(out_dir)
        return None

    r1 = []
    if (len(output.strip()) > 0):
        r1 = json.loads(output)
    
    os.remove(out_dir)
    return (r, r1)

def _read_doc_text_strings(data):
    """Use a heuristic to read in the document text. This is used as a
    fallback if reading the text with libreoffice fails.

    @param data (str) The read in Office file (data).

    @return (tuple) A 2 element tuple where the 1st element is the
    strings grabbed from the raw Word file data and the 2nd element is
    an empty list (no table data).

    """

    str_list = re.findall("[^\x00-\x1F\x7F-\xFF]{4,}", data)
    r = []
    for s in str_list:
        r.append(s)
    
    return (r, [])

def _read_doc_text(fname, data=None):
    """Read in text from the given document.

    @param data (str) The read in Office file (data).

    @return (tuple) Returns a tuple containing the doc text and a list
    of tuples containing dumped tables.

    """

    if (data is None):
        try:
            f = open(fname, 'rb')
            data = f.read()
            f.close()
        except Exception as e:
            log.error("Cannot read document text from " + str(fname) + ". " + str(e))
            return ""

    r = _read_doc_text_libreoffice(data)
    if (r is not None):
        return r

    r = _read_doc_text_strings(data)

    return r

def _get_doc_var_info(ole):
    """Get the byte offset and size of the chunk of data containing the
    document variables. This information is read from the FIB
    (https://msdn.microsoft.com/en-us/library/dd944907(v=office.12).aspx). The
    doc vars appear in the 1Table or 0Table stream.

    @param ole (OLE object) The olevba OLE object for the file being
    analyzed.

    @return (tuple) A 2 element tuple where the 1st element is the
    byte offset os the document variables and the 2nd element is the
    size of the document variable data chunk.

    """

    if (not ole.exists('worddocument')):
        return (None, None)
    data = ole.openstream("worddocument").read()

    fib_offset = 32 + 2 + 28 + 2 + 88 + 2 + (120 * 4)
    tmp = data[fib_offset+3] + data[fib_offset+2] + data[fib_offset+1] + data[fib_offset]
    doc_var_offset = struct.unpack('!I', tmp)[0]

    fib_offset = 32 + 2 + 28 + 2 + 88 + 2 + (120 * 4) + 4
    tmp = data[fib_offset+3] + data[fib_offset+2] + data[fib_offset+1] + data[fib_offset]
    doc_var_size = struct.unpack('!I', tmp)[0]
    
    return (doc_var_offset, doc_var_size)

def _read_payload_default_target_frame(data, vm):
    """Read and save the custom DefaultTargetFrame value from an Office
    file.

    @param data (str) The read in Office file (data).

    @param vm (SimulationVBA object) The SimulationVBA emulation engine
    object that will do the emulation. The read values will be saved
    in the given emulation engine.

    """

    def_targ_frame_val = get_defaulttargetframe_text(data)
    if (def_targ_frame_val is not None):
        vm.globals["DefaultTargetFrame"] = def_targ_frame_val
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Added DefaultTargetFrame = " + str(def_targ_frame_val) + " to globals.")
    
def _read_payload_form_strings(vba, vm):
    """Read in and save the text values of OLE forms as given by the
    output of olevba.

    @param vba (str) The VBA code to analyze, generated with
    olevba. Note that olevba includes the form strings in the output.

    @param vm (SimulationVBA object) The SimulationVBA emulation engine
    object that will do the emulation. The read values will be saved
    in the given emulation engine.

    """


    tmp_form_strings = read_form_strings(vba)
    stream_form_map = {}
    for string_info in tmp_form_strings:
        stream_name = string_info[0]
        if (stream_name not in stream_form_map):
            stream_form_map[stream_name] = []
        curr_form_string = string_info[1]
        stream_form_map[stream_name].append(curr_form_string)

    for stream_name in stream_form_map:
        stream_key = stream_name.lower()
        tmp_name = (stream_name + ".Controls").lower()
        form_strings = stream_form_map[stream_name]
        vm.globals[tmp_name] = form_strings
        vm.globals[stream_key] = stream_name
        if (len(form_strings) > 0):
            best_string = sorted(form_strings, key=lambda x: len(str(x)), reverse=True)[0]
            vm.globals[stream_key + ".*"] = best_string
            vm.globals[stream_key + ".text"] = best_string
            vm.globals[stream_key + ".value"] = best_string
            vm.globals[stream_key + ".caption"] = best_string
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Added VBA form Control values %r = %r to globals." % (tmp_name, form_strings))

def _get_form_var_val(var_name, form_vars):
    """Fix the raw value of the text associated with a given OLE form variable.

    @param var_name (str) The name of the form variable whose value is
    to be fixed.

    @param form_vars (dict) A map from form variable names to raw
    values.

    @return (str) The fixed formm variable value. '' will be returned
    if the form variable is not found in form_vars.

    """

    r = form_vars[var_name] if (var_name in form_vars and form_vars[var_name] is not None) else ''
    r = r.replace('\xb1', '').replace('\x03', '')
    return r
    
def _read_payload_form_vars(vba, vm):
    """Read and save the text values associated with OLE form variables.

    @param vba (str) The VBA code to analyze, generated with
    olevba. Note that olevba includes the form strings in the output.

    @param vm (SimulationVBA object) The SimulationVBA emulation engine
    object that will do the emulation. The read values will be saved
    in the given emulation engine.

    """

    log.info("Reading form variables...")
    try:

        for (_, stream_path, form_variables) in vba.extract_form_strings_extended():
            if form_variables is not None:


                var_name = form_variables['name']
                if (var_name is None):
                    continue

                macro_name = stream_path
                if ("/" in macro_name):
                    start = macro_name.rindex("/") + 1
                    macro_name = macro_name[start:]

                global_var_name = (macro_name + "." + var_name).encode('ascii', 'ignore').replace("\x00", "")
                tag = _get_form_var_val('tag', form_variables)

                caption = _get_form_var_val('caption', form_variables)
                if 'value' in form_variables:
                    val = form_variables['value']
                else:
                    val = caption

                control_tip_text = _get_form_var_val('control_tip_text', form_variables)

                group_name = _get_form_var_val('group_name', form_variables)
                if (len(group_name) > 10):
                    group_name = group_name[3:]
                
                if (val is None):
                    val = caption

                if ((val == '') and (tag == '') and (caption == '')):
                    continue


                name = global_var_name.lower()                        
                vm.globals[name] = val
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("1. Added VBA form variable %r = %r to globals." % \
                              (global_var_name, val))
                vm.globals[name + ".tag"] = tag
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("1. Added VBA form variable %r = %r to globals." % \
                              (global_var_name + ".Tag", tag))
                vm.globals[name + ".caption"] = caption
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("1. Added VBA form variable %r = %r to globals." % \
                              (global_var_name + ".Caption", caption))
                vm.globals[name + ".controltiptext"] = control_tip_text
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("1. Added VBA form variable %r = %r to globals." % \
                              (global_var_name + ".ControlTipText", control_tip_text))
                vm.globals[name + ".text"] = val
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("1. Added VBA form variable %r = %r to globals." % \
                              (global_var_name + ".Text", val))
                vm.globals[name + ".value"] = val
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("1. Added VBA form variable %r = %r to globals." % \
                              (global_var_name + ".Value", val))
                vm.globals[name + ".groupname"] = group_name
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("1. Added VBA form variable %r = %r to globals." % \
                              (global_var_name + ".GroupName", group_name))

                if ("." in name):

                    control_name = name[:name.index(".")] + ".controls"
                    if (control_name not in vm.globals):
                        vm.globals[control_name] = []

                    control_data = {}
                    control_data["value"] = val
                    control_data["tag"] = tag
                    control_data["caption"] = caption
                    control_data["controltiptext"] = control_tip_text
                    control_data["text"] = val
                    control_data["groupname"] = group_name

                    if (log.getEffectiveLevel() == logging.DEBUG):
                        log.debug("Added index VBA form control data " + control_name + \
                                  "(" + str(len(vm.globals[control_name])) + ") = " + str(control_data))
                    vm.globals[control_name].append(control_data)
                        
                short_name = global_var_name.lower()
                if ("." in short_name):
                    short_name = short_name[short_name.rindex(".") + 1:]
                    vm.globals[short_name] = val
                    if (log.getEffectiveLevel() == logging.DEBUG):
                        log.debug("1. Added VBA form variable %r = %r to globals." % \
                                  (short_name, val))
                    vm.globals[short_name + ".tag"] = tag
                    if (log.getEffectiveLevel() == logging.DEBUG):
                        log.debug("1. Added VBA form variable %r = %r to globals." % \
                                  (short_name + ".Tag", tag))
                    vm.globals[short_name + ".caption"] = caption
                    if (log.getEffectiveLevel() == logging.DEBUG):
                        log.debug("1. Added VBA form variable %r = %r to globals." % \
                                  (short_name + ".Caption", caption))
                    vm.globals[short_name + ".controltiptext"] = control_tip_text
                    if (log.getEffectiveLevel() == logging.DEBUG):
                        log.debug("1. Added VBA form variable %r = %r to globals." % \
                                  (short_name + ".ControlTipText", control_tip_text))
                    vm.globals[short_name + ".text"] = val
                    if (log.getEffectiveLevel() == logging.DEBUG):
                        log.debug("1. Added VBA form variable %r = %r to globals." % \
                                  (short_name + ".Text", val))
                        vm.globals[short_name + ".groupname"] = group_name
                    if (log.getEffectiveLevel() == logging.DEBUG):
                        log.debug("1. Added VBA form variable %r = %r to globals." % \
                                  (short_name + ".GroupName", group_name))
                
    except Exception as e:

        log.warning("Cannot read form strings. " + str(e) + ". Trying fallback method.")
        try:
            count = 0
            skip_strings = ["Tahoma", "Tahomaz"]
            for (_, stream_path, form_string) in vba.extract_form_strings():
                if ((len(form_string) > 100) and (entropy(form_string) < 1)):
                    continue
                if (form_string.startswith("\x80")):
                    form_string = form_string[1:]
                if (form_string in skip_strings):
                    continue
                bad_char_count = 0
                for c in form_string:
                    if (not (ord(c) > 31 and ord(c) < 127)):
                        bad_char_count += 1
                if (((bad_char_count + 0.0) / len(form_string)) > .1):
                    continue

                global_var_name = stream_path
                if ("/" in global_var_name):
                    tmp = global_var_name.split("/")
                    if (len(tmp) == 3):
                        global_var_name = tmp[1]
                if ("/" in global_var_name):
                    global_var_name = global_var_name[:global_var_name.rindex("/")]
                global_var_name_orig = global_var_name
                global_var_name += "*" + str(count)
                count += 1
                vm.globals[global_var_name.lower()] = form_string
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("2. Added VBA form variable %r = %r to globals." % (global_var_name.lower(), form_string))
                tmp_name = global_var_name_orig.lower() + ".*"
                if (tmp_name not in vm.globals.keys()):
                    vm.globals[tmp_name] = form_string
                    if (log.getEffectiveLevel() == logging.DEBUG):
                        log.debug("2. Added VBA form variable %r = %r to globals." % (tmp_name, form_string))
                    specific_names = ["textbox1", "label1"]
                    for specific_name in specific_names:
                        tmp_name = specific_name
                        vm.globals[tmp_name] = form_string
                        if (log.getEffectiveLevel() == logging.DEBUG):
                            log.debug("2. Added VBA form variable %r = %r to globals." % (tmp_name, form_string))
                        tmp_name = specific_name + ".caption"
                        vm.globals[tmp_name] = form_string
                        if (log.getEffectiveLevel() == logging.DEBUG):
                            log.debug("2. Added VBA form variable %r = %r to globals." % (tmp_name, form_string))
                        tmp_name = global_var_name_orig.lower() + "." + specific_name + ".caption"
                        vm.globals[tmp_name] = form_string
                        if (log.getEffectiveLevel() == logging.DEBUG):
                            log.debug("2. Added VBA form variable %r = %r to globals." % (tmp_name, form_string))
                        tmp_name = specific_name + ".text"
                        vm.globals[tmp_name] = form_string
                        if (log.getEffectiveLevel() == logging.DEBUG):
                            log.debug("2. Added VBA form variable %r = %r to globals." % (tmp_name, form_string))
                        tmp_name = global_var_name_orig.lower() + "." + specific_name + ".text"
                        vm.globals[tmp_name] = form_string
                        if (log.getEffectiveLevel() == logging.DEBUG):
                            log.debug("2. Added VBA form variable %r = %r to globals." % (tmp_name, form_string))
        except Exception as e:
            log.error("Cannot read form strings. " + str(e) + ". Fallback method failed.")

    
def _read_payload_embedded_obj_text(data, vm):
    """Read in and save the tag and caption associated with Embedded OLE
    Objects in an Office document.

    @param data (str) The read in Office file (data).

    @param vm (SimulationVBA object) The SimulationVBA emulation engine
    object that will do the emulation. The read values will be saved
    in the given emulation engine.

    """

    log.info("Reading embedded object text fields...")
    for (var_name, caption_val, tag_val) in _get_embedded_object_values(data):
        tag_name = var_name.lower() + ".tag"
        vm.doc_vars[tag_name] = tag_val
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Added potential VBA object tag text %r = %r to doc_vars." % \
                      (tag_name, tag_val))
        caption_name = var_name.lower() + ".caption"
        vm.doc_vars[caption_name] = caption_val
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Added potential VBA object caption text %r = %r to doc_vars." % \
                      (caption_name, caption_val))    

def _read_payload_custom_doc_props(data, vm):
    """Read in and save custom document property names and values from
    the DocumentSummaryInformation OLE stream.

    @param data (str) The read in Office file (data).

    @param vm (SimulationVBA object) The SimulationVBA emulation engine
    object that will do the emulation. The read values will be saved
    in the given emulation engine.

    """

    log.info("Reading custom document properties...")
    for (var_name, var_val) in _read_custom_doc_props(data):
        vm.doc_vars[var_name.lower()] = var_val
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Added potential VBA custom doc prop variable %r = %r to doc_vars." % (var_name, var_val))
    
def _read_payload_textbox_text(data, vba_code, vm):
    """Read in and save text hidden in TextBox and RichText objects.

    @param data (str) The read in Office file (data).

    @param vba_code (str) The VBA macro code from the Office file.

    @param vm (SimulationVBA object) The SimulationVBA emulation engine
    object that will do the emulation. The read values will be saved
    in the given emulation engine.

    """

    log.info("Reading TextBox and RichEdit object text fields...")
    object_data = get_ole_textbox_values(data, vba_code)
    tmp_data = get_msftedit_variables(data)
    object_data.extend(tmp_data)
    tmp_data = get_customxml_text(data)
    object_data.extend(tmp_data)
    tmp_data = get_drawing_titles(data)
    object_data.extend(tmp_data)
    for (var_name, var_val) in object_data:
        var_name_variants = [var_name,
                             "ActiveDocument." + var_name,
                             var_name + ".Tag",
                             var_name + ".Text",
                             var_name + ".AlternativeText",
                             var_name + ".Title",
                             var_name + ".Value",
                             var_name + ".Caption",
                             var_name + ".Content",
                             var_name + ".ControlTipText",
                             "me." + var_name,
                             "me." + var_name + ".Tag",
                             "me." + var_name + ".Text",
                             "me." + var_name + ".AlternativeText",
                             "me." + var_name + ".Title",
                             "me." + var_name + ".Value",
                             "me." + var_name + ".Caption",
                             "me." + var_name + ".Content",
                             "me." + var_name + ".ControlTipText"]
        for tmp_var_name in var_name_variants:

            if ((isinstance(var_val, str)) and
                (len(var_val) > 1000)):
                num_1st = float(var_val.count(var_val[0]))
                pct = num_1st/len(var_val) * 100
                if (pct > 95):
                    log.warning("Not assigning " + tmp_var_name + " value '" + var_val[:15] + "...'. " +\
                                "Too many repeated characters.")
                    continue

            tmp_var_val = var_val
            if ((tmp_var_name == 'ActiveDocument.Sections') or
                (tmp_var_name == 'Sections')):
                tmp_var_val = [var_val, var_val]
            if ((tmp_var_name.lower() in vm.doc_vars) and
                (len(str(vm.doc_vars[tmp_var_name.lower()])) > len(str(tmp_var_val)))):
                continue
            vm.doc_vars[tmp_var_name.lower()] = tmp_var_val
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("Added potential VBA OLE form textbox text (1) %r = %r to doc_vars." % (tmp_var_name, tmp_var_val))

        page_pat = r"Page(\d+)"
        if (re.match(page_pat, var_name)):
            page_index = str(int(re.findall(page_pat, var_name)[0]) - 1)
            page_var_name = "Pages('" + page_index + "')"
            tab_var_name = "Tabs('" + page_index + "')"
            var_name_variants = [page_var_name,
                                 "ActiveDocument." + page_var_name,
                                 page_var_name + ".Tag",
                                 page_var_name + ".Text",
                                 page_var_name + ".Caption",
                                 page_var_name + ".ControlTipText",
                                 "me." + page_var_name,
                                 "me." + page_var_name + ".Tag",
                                 "me." + page_var_name + ".Text",
                                 "me." + page_var_name + ".Caption",
                                 "me." + page_var_name + ".ControlTipText",
                                 tab_var_name,
                                 "ActiveDocument." + tab_var_name,
                                 tab_var_name + ".Tag",
                                 tab_var_name + ".Text",
                                 tab_var_name + ".Caption",
                                 tab_var_name + ".ControlTipText",
                                 "me." + tab_var_name,
                                 "me." + tab_var_name + ".Tag",
                                 "me." + tab_var_name + ".Text",
                                 "me." + tab_var_name + ".Caption",
                                 "me." + tab_var_name + ".ControlTipText"]

            if (not got_inline_shapes):
                var_name_variants.extend(["InlineShapes('" + page_index + "').TextFrame.TextRange.Text",
                                          "InlineShapes('" + page_index + "').TextFrame.ContainingRange",
                                          "InlineShapes('" + page_index + "').AlternativeText",
                                          "InlineShapes('" + page_index + "').AlternativeText$",
                                          "InlineShapes.Item('" + page_index + "').TextFrame.TextRange.Text",
                                          "InlineShapes.Item('" + page_index + "').TextFrame.ContainingRange",
                                          "InlineShapes.Item('" + page_index + "').AlternativeText",
                                          "InlineShapes.Item('" + page_index + "').AlternativeText$",
                                          "StoryRanges.Item('" + page_index + "')",
                                          "me.StoryRanges.Item('" + page_index + "')"])
            for tmp_var_name in var_name_variants:
                vm.doc_vars[tmp_var_name.lower()] = var_val
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("Added potential VBA OLE form textbox text (2) %r = %r to doc_vars." % (tmp_var_name, var_val))

                    
got_inline_shapes = False                    
def _read_payload_inline_shape_text(data, vm):
    """Read in and save the text associated with InlineShape objects in
    the document.

    @param data (str) The read in Office file (data).

    @param vm (SimulationVBA object) The SimulationVBA emulation engine
    object that will do the emulation. The read values will be saved
    in the given emulation engine.

    """

    log.info("Reading InlineShapes object text fields...")
    global got_inline_shapes
    got_inline_shapes = False
    for (var_name, var_val) in _get_inlineshapes_text_values(data):
        got_inline_shapes = True
        vm.doc_vars[var_name.lower()] = var_val
        log.info("Added potential VBA InlineShape text %r = %r to doc_vars." % (var_name, var_val))
    
def _read_payload_shape_text(data, vm):
    """Read in and save the text associated with Shape objects in a
    document saved as Flat OPC XML files.

    @param data (str) The read in Office file (data).

    @param vm (SimulationVBA object) The SimulationVBA emulation engine
    object that will do the emulation. The read values will be saved
    in the given emulation engine.

    """

    log.info("Reading Shapes object text fields...")
    got_it = False
    shape_text = get_shapes_text_values(data, 'worddocument')
    pos = 1
    for (var_name, var_val) in shape_text:
        got_it = True
        var_name = var_name.lower()
        vm.doc_vars[var_name] = var_val
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Added potential VBA Shape text %r = %r to doc_vars." % (var_name, var_val))
        vm.doc_vars["thisdocument."+var_name] = var_val
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Added potential VBA Shape text %r = %r to doc_vars." % ("thisdocument."+var_name, var_val))
        vm.doc_vars["thisdocument."+var_name+".caption"] = var_val
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Added potential VBA Shape text %r = %r to doc_vars." % ("thisdocument."+var_name+".caption", var_val))
        vm.doc_vars["activedocument."+var_name] = var_val
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Added potential VBA Shape text %r = %r to doc_vars." % ("activedocument."+var_name, var_val))
        vm.doc_vars["activedocument."+var_name+".caption"] = var_val
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Added potential VBA Shape text %r = %r to doc_vars." % ("activedocument."+var_name+".caption", var_val))
        tmp_name = "shapes('" + var_name + "').textframe.textrange.text"
        vm.doc_vars[tmp_name] = var_val
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Added potential VBA Shape text %r = %r to doc_vars." % (tmp_name, var_val))
        tmp_name = "shapes('" + str(pos) + "').textframe.textrange.text"
        vm.doc_vars[tmp_name] = var_val
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Added potential VBA Shape text %r = %r to doc_vars." % (tmp_name, var_val))
        tmp_name = "me.storyranges('" + str(pos) + "')"
        vm.doc_vars[tmp_name] = var_val
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Added potential VBA StoryRange text %r = %r to doc_vars." % (tmp_name, var_val))
        tmp_name = "ActiveDocument.shapes('" + str(pos) + "').AlternativeText"
        vm.doc_vars[tmp_name] = var_val
        vm.doc_vars[tmp_name.lower()] = var_val
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Added potential VBA Shape text %r = %r to doc_vars." % (tmp_name, var_val))
        pos += 1
    if (not got_it):
        shape_text = get_shapes_text_values(data, '1table')
        for (var_name, var_val) in shape_text:
            vm.doc_vars[var_name.lower()] = var_val
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("Added potential VBA Shape text %r = %r to doc_vars." % (var_name, var_val))
    
def _read_payload_doc_comments(data, vm):
    """Read in and save the comments in an Office document.

    @param data (str) The read in Office file (data).

    @param vm (SimulationVBA object) The SimulationVBA emulation engine
    object that will do the emulation. The read values will be saved
    in the given emulation engine.

    """

    log.info("Reading document comments...")
    comments = get_comments(data)
    if (len(comments) > 0):
        vm.comments = []
        for (_, comment_text) in comments:
            vm.comments.append(comment_text)

def _read_payload_doc_vars(data, orig_filename, vm):
    """Read and save document variables from Office 97 or 2007+ files.

    @param data (str) The read in Office file data. Can be None if data
    should be read from a file (orig_fname).

    @param orig_fname (str) The name of the Office file to analyze. Can be
    None if data is given (data).

    @param vm (SimulationVBA object) The SimulationVBA emulation engine
    object that will do the emulation. The read values will be saved
    in the given emulation engine.

    """

    log.info("Reading document variables...")
    for (var_name, var_val) in _read_doc_vars(data, orig_filename):
        vm.doc_vars[var_name] = var_val
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Added potential VBA doc variable %r = %r to doc_vars." % (var_name, var_val))
        vm.doc_vars[var_name.lower()] = var_val
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Added potential VBA doc variable %r = %r to doc_vars." % (var_name.lower(), var_val))



def _read_payload_ooxml_context(data, orig_filename, vm):
    """Read static OOXML workbook runtime values into the emulation context."""
    try:
        ooxml_context.read_ooxml_context(data, orig_filename, vm)
    except Exception as e:
        log.warning("Cannot read OOXML workbook runtime context. " + str(e))

def read_payload_hiding_places(data, orig_filename, vm, vba_code, vba):
    """
    Read in text values from all of the various places in Office
    97/2000+ that text values can be hidden. This reads values from
    things like ActiveX captions, embedded image alternate text,
    document variables, form variables, etc.

    @param (data) The contents (bytes) of the Office file being
    analyzed.

    @param orig_filename (str) The name of the Office file being
    analyzed.

    @param vm (SimulationVBA object) The SimulationVBA emulation engine
    object that will do the emulation. The read values will be saved
    in the given emulation engine.

    @param vba_code (str) The VB code that will be emulated.

    @param vba (VBA_Parser object) The olevba VBA_Parser object for
    reading the Office file being analyzed.
    """

    _read_payload_ooxml_context(data, orig_filename, vm)

    _read_payload_doc_vars(data, orig_filename, vm)

    _read_payload_doc_comments(data, vm)
                
    _read_payload_shape_text(data, vm)

    _read_payload_inline_shape_text(data, vm)
                    
    _read_payload_textbox_text(data, vba_code, vm)
                            
    _read_payload_custom_doc_props(data, vm)

    _read_payload_embedded_obj_text(data, vm)
                
    log.info("Reading document text and tables...")
    vm.doc_text, vm.doc_tables = _read_doc_text('', data=data)

    _read_payload_form_vars(vba, vm)

    _read_payload_form_strings(vba, vm)

    _read_payload_default_target_frame(data, vm)

    
if __name__ == '__main__':
    print get_shapes_text_values(sys.argv[1], "worddocument")
    print get_shapes_text_values(sys.argv[1], '1table')
