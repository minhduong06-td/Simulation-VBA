#!/usr/bin/env pypy


from __future__ import print_function

import pyparsing
pyparsing.ParserElement.enablePackrat(cache_size_limit=100000)

import shutil
import logging
import json
import random
import optparse
import sys
import os
import traceback
import colorlog
import re
from datetime import datetime
from datetime import timedelta
import zipfile
import io

import prettytable
from oletools.thirdparty.xglob import xglob
from oletools.olevba import VBA_Parser, filter_vba, FileOpenError
import olefile
    
from core.meta import get_metadata_exif

_thismodule_dir = os.path.normpath(os.path.abspath(os.path.dirname(__file__)))
if _thismodule_dir not in sys.path:
    sys.path.insert(0, _thismodule_dir)

import core
import core.excel as excel
import core.read_ole_fields as read_ole_fields
import core.deobfuscation as deobfuscation
from core.utils import safe_print
from core.utils import safe_str_convert

from core.logger import log
from core.logger import CappedFileHandler
from logging import FileHandler




__version__ = '1.0.3'



    
def get_vb_contents_from_hta(vba_code):
    """Pull out Visual Basic code from .hta file contents.

    @param vba_code (str) The HTA file contents from which to extract
    the VBScript code.

    @return (str) If the given data is HTA that contains VBScript
    script elements, the VBScript in the HTA is returned. If the given
    data is not VBScript HTA, the original data is returned.

    """

    return deobfuscation.extract_vb_from_hta(vba_code)


def deobfuscate_simulate_text(data, entry_points=None):
    return deobfuscation.simulate_deobfuscation(data, entry_points=entry_points)


def _process_deob_simulate_input(filename, data, entry_points=None):
    if data is None:
        with open(filename, 'rb') as input_file:
            data = input_file.read()
    deobfuscated, actions = deobfuscate_simulate_text(data, entry_points=entry_points)
    safe_print(deobfuscated)
    safe_print('')
    safe_print('Recorded Stubbed Actions:')
    if not actions:
        safe_print('(none)')
    else:
        for action, params, description in actions:
            safe_print('%s\t%s\t%s' % (action, params, description))
    return deobfuscated, actions
    
def parse_stream(subfilename,
                 stream_path=None,
                 vba_filename=None,
                 vba_code=None,
                 strip_useless=False,
                 local_funcs=None):
    """Parse the macros from a single OLE stream.

    @param subfilename (str) The name of the file containing the    
    macros.

    @param stream_path (??) ??

    @param vba_filename (??) ??

    @param vba_code (str) The macro code to parse.

    @param local_funcs (list) A list of the names of already declared
    local VBA functions.

    @return (Module object) A parsed module object.

    """
    
    if (local_funcs is None):
        local_funcs = []
    
    core.vba_object.limits_exceeded(throw_error=True)
    
    if (stream_path is None):
        subfilename, stream_path, vba_filename, vba_code = subfilename

    if (repr(stream_path).strip() == "'xlm_macro'"):
        log.warning("Skipping XLM macro stream...")
        return "empty"
        
    vba_code = core.vba_collapse_long_lines(vba_code)
        
    vba_code = filter_vba(vba_code)

    vba_code = get_vb_contents_from_hta(vba_code)

    if (read_ole_fields.is_garbage_vba(vba_code)):
        raise ValueError("VBA looks corrupted. Not analyzing.")

    if (vba_code.strip().startswith("<?xml")):
        log.warning("Skipping XML stream.")
        return "empty"
    
    if (strip_useless):
        vba_code = core.strip_lines.strip_useless_code(vba_code, local_funcs)
    safe_print('-'*79)
    safe_print('VBA MACRO %s ' % vba_filename)
    safe_print('in file: %s - OLE stream: %s' % (subfilename, repr(stream_path)))
    safe_print('- '*39)
    
    m = None
    if vba_code.strip() == '':
        safe_print('(empty macro)')
        m = "empty"
    else:
        safe_print('-'*79)
        safe_print('VBA CODE (with long lines collapsed):')
        safe_print(vba_code)
        safe_print('-'*79)
        safe_print('PARSING VBA CODE:')
        try:
            m = core.module.parseString(vba_code + "\n", parseAll=True)[0]
            pyparsing.ParserElement.resetCache()
            m.code = vba_code
        except pyparsing.ParseException as err:
            safe_print(err.line)
            safe_print(" "*(err.column-1) + "^")
            safe_print(err)
            log.error("Parse Error. Processing Aborted.")
            return None

    core.vba_object.limits_exceeded(throw_error=True)
        
    return m

def get_all_local_funcs(vba):
    """Get the names of all locally defined functions. Also get the names
    of all defined constants. The constant names are saved in
    core.strip_lines.defined_constants.

    @params vba (VBA_Parser object) The olevba VBA_Parser object for
    reading the Office file being analyzed.

    """

    pat = r"(?:Sub |Function )([^\(]+)"
    r = []
    for (_, _, _, vba_code) in vba.extract_macros():
        if (vba_code is None):
            continue

        for line in vba_code.split("\n"):
            names = re.findall(pat, line)
            r.extend(names)

        core.strip_lines.find_defined_constants(vba_code)

    return r
            
def parse_streams(vba, strip_useless=False):
    """Parse all the VBA streams and return list of parsed module
    objects.

    @params vba (VBA_Parser object) The olevba VBA_Parser object for
    reading the Office file being analyzed.

    @param strip_useless (boolean) Flag turning on/off modification of
    VB code prior to parsing.

    @return (list) A list of parsed Module objects.

    """

    local_funcs = get_all_local_funcs(vba)
    
    r = []
    for (subfilename, stream_path, vba_filename, vba_code) in vba.extract_macros():
        m = parse_stream(subfilename, stream_path, vba_filename, vba_code, strip_useless, local_funcs)
        if (m is None):
            return None
        r.append(m)
    return r


def read_excel_sheets(fname):
    """Read all the sheets of a given Excel file as CSV and return them
    as a ExcelBook object.

    @param fname (str) The name of the Excel file to read.

    @return (core.excel.ExceBook object) On success return the Excel
    sheets as an ExcelBook object. Returns None on error.

    """

    try:
        f = open(fname, 'rb')
        data = f.read()
        f.close()
        return excel.load_excel_libreoffice(data)
    except Exception as e:
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Reading Excel sheets failed. " + str(e))
        return None
    
def pull_urls_office97(fname):
    """Pull URLs directly from an Office97 file.

    @param fname (str) The name of the file from which to scrape
    URLs.

    @return (set) The URLs scraped from the file. This will be empty
    if there are no URLs.

    """
    return read_ole_fields.pull_urls_office97(fname, False, None)
    

# pylint: disable=too-many-arguments
def process_file(container,
                 filename,
                 data,
                 strip_useless=False,
                 entry_points=None,
                 time_limit=None,
                 verbose=False,
                 display_int_iocs=False,
                 set_log=False,
                 tee_log=False,
                 tee_bytes=0,
                 artifact_dir=None,
                 out_file_name=None,
                 do_jit=False):
    """Process an Office file with VBA macros, a VBScript file, or
    VBScript HTA file with simulation_vba. This is the main programatic
    interface for simulation_vba.

    @param container (str) Path and filename of container if the file is within
    a zip archive, None otherwise.

    @param filename (str) str, path and filename of file on disk, or
    within the container.

    @param data (bytes) content of the file if it is in a container,
    None if it is a file on disk.
 
    @param strip_useless (boolean) Flag turning on/off modification of
    VB code prior to parsing.

    @param entry_points (list) A list of the names (str) of the VB functions
    from which to start emulation.
    
    @param time_limit (int) The emulation time limit, in minutes. If
    None there is not time limit.

    @param verbose (boolean) Flag turning debug logging on/off.

    @param display_int_iocs (boolean) Flag turning on/off the
    reporting of intermediate IOCs (base64 strings and URLs) found
    during the emulation process.

    @param set_log (boolean) A flag??

    @param tee_log (boolean) A flag turning on/off saving all of
    SimulationVBA's output in a text log file. The log file will be
    FNAME.log, where FNAME is the name of the file being analyzed.

    @param tee_bytes (int) If tee_log is true, this gives the number
    of bytes at which to cap the saved log file.

    @param artifact_dir (str) The directory in which to save artifacts
    dropped by the sample under analysis. If None the artifact
    directory will be FNAME_artifacts/ where FNAME is the name of the
    file being analyzed.

    @param out_file_name (str) The name of the file in which to store
    the SimulationVBA analysis results as JSON. If None no JSON results
    will be saved.

    @param do_jit (str) A flag turning on/off doing VB -> Python
    transpiling of loops to speed up loop emulation.

    @return (list) A list of actions if actions found, an empty list
    if no actions found, and None if there was an error.

    """
    
    if verbose:
        colorlog.basicConfig(level=logging.DEBUG, format='%(log_color)s%(levelname)-8s %(message)s')
    elif set_log:
        colorlog.basicConfig(level=logging.INFO, format='%(log_color)s%(levelname)-8s %(message)s')

    if tee_bytes > 0:
        tee_log = True

    if tee_log:

        tee_filename = "./" + filename
        if ("/" in filename):
            tee_filename = "./" + filename[filename.rindex("/") + 1:]

        if tee_bytes > 0:
            capped_handler = CappedFileHandler(tee_filename + ".log", sizecap=tee_bytes)
            capped_handler.setFormatter(logging.Formatter("%(levelname)-8s %(message)s"))
            log.addHandler(capped_handler)
        else:
            file_handler = FileHandler(tee_filename + ".log", mode="w")
            file_handler.setFormatter(logging.Formatter("%(levelname)-8s %(message)s"))
            log.addHandler(file_handler)

    if (isinstance(data, Exception)):
        log.error("Cannot open file '" + str(filename) + "'.")
        return None
    
    if not data:
        if container:
            display_filename = '%s in %s' % (filename, container)
        else:
            display_filename = filename
        safe_print('='*79)
        safe_print('FILE: ' + str(display_filename))
        try:
            input_file = open(filename,'rb')
            data = input_file.read()
            input_file.close()
        except IOError as e:
            log.error("Cannot open file '" + str(filename) + "'. " + str(e))
            return None
    r = _process_file(filename,
                      data,
                      strip_useless=strip_useless,
                      entry_points=entry_points,
                      time_limit=time_limit,
                      display_int_iocs=display_int_iocs,
                      artifact_dir=artifact_dir,
                      out_file_name=out_file_name,
                      do_jit=do_jit)

    colorlog.basicConfig(level=logging.ERROR, format='%(log_color)s%(levelname)-8s %(message)s')

    return r

def _remove_duplicate_iocs(iocs):
    """Remove IOC strings that are substrings of other IOC strings.

    @param iocs (list) List of IOCs (str).

    @return (set) The original IOC list with duplicate-ish IOC strings
    stripped out.

    """

    r = set()
    skip = set()
    log.info("Found " + str(len(iocs)) + " possible IOCs. Stripping duplicates...")
    for ioc1 in iocs:
        
        if (read_ole_fields.is_garbage_vba(ioc1, test_all=True, bad_pct=.25)):
            skip.add(ioc1)
            continue

        keep_curr = True
        for ioc2 in iocs:
            if (ioc2 in skip):
                continue
            if ((ioc1 != ioc2) and (ioc1 in ioc2)):
                keep_curr = False
                break
            if ((ioc1 != ioc2) and (ioc2 in ioc1)):
                skip.add(ioc2)
        if (keep_curr):
            r.add(ioc1)

    return r

def _get_vba_parser(data):
    """Get an olevba VBA_Parser object for reading an Office file. This
    handles regular Office files and HTA files with VBScript script
    elements.

    @param data (str) The file contents for which to generate a
    VBA_Parser.

    @return (VBA_Parser object) On success, the olevba VBA_Parser
    object for the given file contents. On error, None.

    """
    
    vba = None
    try:
        vba = VBA_Parser('', data, relaxed=True)
    except Exception as e:

        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Creating VBA_PArser() Failed. Trying as HTA. " + str(e))
        
        extracted_data = get_vb_contents_from_hta(data)

        vba = VBA_Parser('', extracted_data, relaxed=True)

    return vba

def pull_embedded_pe_files(data, out_dir):
    """Directly pull out any PE files embedded in the given data. The PE
    files will be saved in a directory and will be named things like
    embedded*.exe.

    @param data (str) The contents of the file being analyzed.

    @param out_dir (str) The directory in which to save extracted PE
    files.

    """

    if core.filetype.is_office2007_file(data, is_data=True):

        data_io = io.BytesIO(data)
        with zipfile.ZipFile(data_io, "r") as f:
            for name in f.namelist():
                curr_data = f.read(name)
                pull_embedded_pe_files(curr_data, out_dir)
        return
    
    pe_pat = r"MZ.{70,80}This program (?:(?:cannot be run in DOS mode\.)|(?:must be run under Win32))"
    if (re.search(pe_pat, data) is None):
        return


    pe_starts = []
    for match in re.finditer(pe_pat, data):
        pe_starts.append(match.span()[0])
    pe_starts.append(len(data))

    if not os.path.isdir(out_dir):
        os.makedirs(out_dir)
    
    pos = 0
    out_index = 0
    while (pos < len(pe_starts) - 1):
        curr_data = data[pe_starts[pos]:pe_starts[pos+1]]
        curr_name = out_dir + "/embedded_pe" + str(out_index) + ".bin"
        while os.path.isfile(curr_name):
            out_index += 1
            curr_name = out_dir + "/embedded_pe" + str(out_index) + ".bin"
        f = open(curr_name, "wb")
        f.write(curr_data)
        f.close()
        pos += 1
        out_index += 1

def _report_analysis_results(vm, data, display_int_iocs, orig_filename, out_file_name):
    """Report analysis results (screen and file) to the user. Results will
    be printed to stdout and saved in an output file as JSON if needed.

    @param vm (SimulationVBA object) The SimulationVBA emulation engine
    object that did the emulation.

    @param data (str) The read in Office file (data).

    @param display_int_iocs (boolean) Flag turning on/off the
    reporting of intermediate IOCs (base64 strings and URLs) found
    during the emulation process.

    @param orig_filename (str) path and filename of file on disk, or
    within the container.

    @param out_file_name (str) The name of the file in which to store
    the SimulationVBA analysis results as JSON. If None no JSON results
    will be saved.

    @return (tuple) A 3 element tuple where the 1st element is a list
    of reported actions all converted to strings, the 2nd element is a
    list of unique intermediate IOCs, and the 3rd element is a list of
    shell code bytes injected by the VB (empty list if no shell code).

    """

    safe_print('\nRecorded Actions:')
    safe_print(vm.dump_actions())
    safe_print('')
    full_iocs = core.vba_context.intermediate_iocs
    raw_b64_iocs = read_ole_fields.pull_base64(data)
    for ioc in raw_b64_iocs:
        if (core.vba_context.num_b64_iocs > 200):
            log.warning("Found too many potential base64 IOCs. Skipping the rest.")
            break
        full_iocs.add(ioc)
        core.vba_context.num_b64_iocs += 1

    tmp_iocs = []
    if (len(full_iocs) > 0):
        tmp_iocs = _remove_duplicate_iocs(full_iocs)
        if (display_int_iocs):
            safe_print('Intermediate IOCs:')
            safe_print('')
            for ioc in tmp_iocs:
                safe_print("+---------------------------------------------------------+")
                safe_print(ioc)
            safe_print("+---------------------------------------------------------+")
            safe_print('')

    shellcode_bytes = core.vba_context.get_shellcode_data()
    if (len(shellcode_bytes) > 0):
        safe_print("+---------------------------------------------------------+")
        safe_print("Shell Code Bytes: " + str(shellcode_bytes))
        safe_print("+---------------------------------------------------------+")
        safe_print('')

    pull_embedded_pe_files(data, core.vba_context.out_dir)
                
    safe_print('VBA Builtins Called: ' + str(vm.external_funcs))
    safe_print('')
    safe_print('Finished analyzing ' + str(orig_filename) + " .\n")

    if out_file_name:

        actions_data = []
        for action in vm.actions:
            actions_data.append({
                "action": str(action[0]),
                "parameters": str(action[1]),
                "description": str(action[2])
            })

        out_data = {
            "file_name": orig_filename,
            "potential_iocs": list(tmp_iocs),
            "shellcode" : shellcode_bytes,
            "vba_builtins": vm.external_funcs,
            "actions": actions_data
        }

        try:
            with open(out_file_name, 'w') as out_file:
                out_file.write("\n" + json.dumps(out_data, indent=4))
        except Exception as exc:
            log.error("Failed to output results to output file. " + str(exc))

    str_actions = []
    for action in vm.actions:
        str_actions.append((safe_str_convert(action[0]),
                            safe_str_convert(action[1]),
                            safe_str_convert(action[2])))    

    return (str_actions, tmp_iocs, shellcode_bytes)
        
def _process_file (filename,
                   data,
                   strip_useless=False,
                   entry_points=None,
                   time_limit=None,
                   display_int_iocs=False,
                   artifact_dir=None,
                   out_file_name=None,
                   do_jit=False):
    """Process a single file.

    @param container (str) Path and filename of container if the file is within
    a zip archive, None otherwise.

    @param filename (str) path and filename of file on disk, or within
    the container.

    @param data (bytes) content of the file if it is in a container,
    None if it is a file on disk.

    @param strip_useless (boolean) Flag turning on/off modification of
    VB code prior to parsing.

    @param entry_points (list) A list of the names (str) of the VB functions
    from which to start emulation.

    @param time_limit (int) The emulation time limit, in minutes. If
    None there is not time limit.

    @param display_int_iocs (boolean) Flag turning on/off the
    reporting of intermediate IOCs (base64 strings and URLs) found
    during the emulation process.

    @param artifact_dir (str) The directory in which to save artifacts
    dropped by the sample under analysis. If None the artifact

    @param out_file_name (str) The name of the file in which to store
    the SimulationVBA analysis results as JSON. If None no JSON results
    will be saved.

    @param do_jit (str) A flag turning on/off doing VB -> Python
    transpiling of loops to speed up loop emulation.

    @return (list) A list of actions if actions found, an empty list
    if no actions found, and None if there was an error.

    """

    sys.setrecursionlimit(13000)

    if (time_limit is not None):
        core.vba_object.max_emulation_time = datetime.now() + timedelta(minutes=time_limit)

    log.info("Starting emulation...")
    vm = core.SimulationVBA(filename, data, do_jit=do_jit)
    orig_filename = filename
    if (entry_points is not None):
        for entry_point in entry_points:
            vm.entry_points.append(entry_point)
    try:
        if (isinstance(data, Exception)):
            data = None
        vba = None
        try:
            vba = _get_vba_parser(data)
        except FileOpenError as e:

            if ("Failed to open file  is not a supported file type, cannot extract VBA Macros." not in str(e)):

                raise e

            data = data.replace("\x00", "")
            vba = _get_vba_parser(data)

        if (vba.detect_vba_macros() or display_int_iocs):

            try:
                log.info("Reading document metadata...")
                ole = olefile.OleFileIO(data)
                vm.set_metadata(ole.get_metadata())
            except Exception as e:
                log.warning("Reading in metadata failed. Trying fallback. " + str(e))
                vm.set_metadata(get_metadata_exif(orig_filename))

            vm.loaded_excel = excel.load_excel(data)

            if (artifact_dir is None):
                artifact_dir = "./"
                if ((filename is not None) and ("/" in filename)):
                    artifact_dir = filename[:filename.rindex("/")]
            only_filename = filename
            if ((filename is not None) and ("/" in filename)):
                only_filename = filename[filename.rindex("/")+1:]
            
            out_dir = None
            if (only_filename is not None):
                out_dir = artifact_dir + "/" + only_filename + "_artifacts/"
                if os.path.exists(out_dir):
                    shutil.rmtree(out_dir)
            else:
                out_dir = "/tmp/tmp_file_" + str(random.randrange(0, 10000000000))
            log.info("Saving dropped analysis artifacts in " + out_dir)
            core.vba_context.out_dir = out_dir
            del filename
                
            log.info("Parsing VB...")
            comp_modules = parse_streams(vba, strip_useless)
            if (comp_modules is None):
                return None
            got_code = False
            for m in comp_modules:
                if (m != "empty"):
                    vm.add_compiled_module(m)
                    got_code = True
            if ((not got_code) and (not display_int_iocs)):
                log.info("No VBA or VBScript found. Exiting.")
                return ([], [], [], [])

            vba_code = ""
            for (_, _, _, macro_code) in vba.extract_macros():
                if (macro_code is not None):
                    vba_code += macro_code

            if (read_ole_fields.is_garbage_vba(vba_code)):
                raise ValueError("VBA looks corrupted. Not analyzing.")

            read_ole_fields.read_payload_hiding_places(data, orig_filename, vm, vba_code, vba)
            
            safe_print("")
            safe_print('-'*79)
            safe_print('TRACING VBA CODE (entrypoint = Auto*):')
            if (entry_points is not None):
                log.info("Starting emulation from function(s) " + str(entry_points))
            pyparsing.ParserElement.resetCache()
            vm.vba = vba
            vm.trace()


            str_actions, tmp_iocs, shellcode_bytes = _report_analysis_results(vm, data, display_int_iocs, orig_filename, out_file_name)
            
            return (str_actions, vm.external_funcs, tmp_iocs, shellcode_bytes)

        else:
            safe_print('Finished analyzing ' + str(orig_filename) + " .\n")
            safe_print('No VBA macros found.')
            safe_print('')
            return ([], [], [], [])

    except Exception as e:

        if (("SystemExit" not in str(e)) and (". Aborting analysis." not in str(e))):
            traceback.print_exc()
        log.error(str(e))

        if isinstance(e, MemoryError):
            log.error("Exiting SimulationVBA with error code 137 (out of memory)")
            sys.exit(137)

        return None

def process_file_scanexpr (container, filename, data):
    """Process a single file.

    @param container (str) Path and filename of container if the file is within
    a zip archive, None otherwise.

    @param filename (str) path and filename of file on disk, or within
    the container.

    @param data (bytes) Content of the file if it is in a container,
    None if it is a file on disk.

    """
    if container:
        display_filename = '%s in %s' % (filename, container)
    else:
        display_filename = filename
    safe_print('='*79)
    safe_print('FILE: ' + str(display_filename))
    all_code = ''
    try:
        import oletools
        oletools.olevba.enable_logging()
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('opening %r' % filename)
        vba = VBA_Parser(filename, data, relaxed=True)
        if vba.detect_vba_macros():

            vm = core.SimulationVBA(filename, data)
            ole = olefile.OleFileIO(filename)
            try:
                vm.set_metadata(ole.get_metadata())
            except Exception as e:
                log.warning("Reading in metadata failed. Trying fallback. " + str(e))
                vm.set_metadata(get_metadata_exif(filename))
            
            for (subfilename, stream_path, vba_filename, vba_code) in vba.extract_macros():
                vba_code = filter_vba(vba_code)
                safe_print('-'*79)
                safe_print('VBA MACRO %s ' % vba_filename)
                safe_print('in file: %s - OLE stream: %s' % (subfilename, repr(stream_path)))
                safe_print('- '*39)
                if vba_code.strip() == '':
                    safe_print('(empty macro)')
                else:
                    safe_print(vba_code)
                    vba_code = core.vba_collapse_long_lines(vba_code)
                    all_code += '\n' + vba_code
            safe_print('-'*79)
            safe_print('EVALUATED VBA EXPRESSIONS:')
            t = prettytable.PrettyTable(('Obfuscated expression', 'Evaluated value'))
            t.align = 'l'
            t.max_width['Obfuscated expression'] = 36
            t.max_width['Evaluated value'] = 36
            for expression, expr_eval in core.scan_expressions(all_code):
                t.add_row((repr(expression), repr(expr_eval)))
                safe_print(t)

        else:
            safe_print('No VBA macros found.')
    except Exception as e:
        log.error("Caught exception. " + str(e))
        if (log.getEffectiveLevel() == logging.DEBUG):
            traceback.print_exc()

    safe_print('')

def print_version():
    """Print SimulationVBA version information.

    """

    safe_print("Version Information:\n")
    safe_print("SimulationVBA:\t\t" + str(__version__))
    safe_print("Python:\t\t\t" + str(sys.version_info))
    safe_print("pyparsing:\t\t" + str(pyparsing.__version__))
    safe_print("olefile:\t\t" + str(olefile.__version__))
    import oletools.olevba
    safe_print("olevba:\t\t\t" + str(oletools.olevba.__version__))

def main():
    """Main function, called when simulation_vba is run from the command
    line.

    """

    sys.setrecursionlimit(13000)
    
    safe_print(''' _    ___                 __  ___            __             
| |  / (_)___  ___  _____/  |/  /___  ____  / /_____  __  __
| | / / / __ \/ _ \/ ___/ /|_/ / __ \/ __ \/ //_/ _ \/ / / /
| |/ / / /_/ /  __/ /  / /  / / /_/ / / / / ,< /  __/ /_/ / 
|___/_/ .___/\___/_/  /_/  /_/\____/_/ /_/_/|_|\___/\__, /  
     /_/                                           /____/   ''')
    safe_print('vmonkey %s - https://github.com/decalage2/ViperMonkey' % __version__)
    safe_print('THIS IS WORK IN PROGRESS - Check updates regularly!')
    safe_print('Please report any issue at https://github.com/decalage2/ViperMonkey/issues')
    safe_print('')

    DEFAULT_LOG_LEVEL = "info"
    LOG_LEVELS = {
        'debug':    logging.DEBUG,
        'info':     logging.INFO,
        'warning':  logging.WARNING,
        'error':    logging.ERROR,
        'critical': logging.CRITICAL
        }

    usage = 'usage: %prog [options] <filename> [filename2 ...]'
    parser = optparse.OptionParser(usage=usage)
    parser.add_option("-r", action="store_true", dest="recursive",
                      help='find files recursively in subdirectories.')
    parser.add_option("-z", "--zip", dest='zip_password', type='str', default=None,
                      help='if the file is a zip archive, open first file from it, using the '
                           'provided password (requires Python 2.6+)')
    parser.add_option("-f", "--zipfname", dest='zip_fname', type='str', default='*',
                      help='if the file is a zip archive, file(s) to be opened within the zip. '
                           'Wildcards * and ? are supported. (default:*)')
    parser.add_option("-e", action="store_true", dest="scan_expressions",
                      help='Extract and evaluate/deobfuscate constant expressions')
    parser.add_option('-l', '--loglevel', dest="loglevel", action="store", default=DEFAULT_LOG_LEVEL,
                      help="logging level debug/info/warning/error/critical (default=%default)")
    parser.add_option("-s", '--strip', action="store_true", dest="strip_useless_code",
                      help='Strip useless VB code from macros prior to parsing.')
    parser.add_option("-j", '--jit', action="store_true", dest="do_jit",
                      help='Speed up emulation by JIT compilation of VB loops to Python.')
    parser.add_option('-i', '--init', dest="entry_points", action="store", default=None,
                      help="Emulate starting at the given function name(s). Use comma seperated "
                           "list for multiple entries.")
    parser.add_option('-t', '--time-limit', dest="time_limit", action="store", default=None,
                      type='int', help="Time limit (in minutes) for emulation.")
    parser.add_option("-c", '--iocs', action="store_true", dest="display_int_iocs",
                      help='Display potential IOCs stored in intermediate VBA variables '
                           'assigned during emulation (URLs and base64).')
    parser.add_option("-v", '--version', action="store_true", dest="print_version",
                      help='Print version information of packages used by SimulationVBA.')
    parser.add_option("-o", "--out-file", action="store", default=None, type="str",
                      help="JSON output file containing resulting IOCs, builtins, and actions")
    parser.add_option("-p", "--tee-log", action="store_true", default=False,
                      help="output also to a file in addition to standard out")
    parser.add_option("-b", "--tee-bytes", action="store", default=0, type="int",
                      help="number of bytes to limit the tee'd log to")
    parser.add_option("--deob", action="store", default=None, type="str",
                      help="Deobfuscation mode. Use '--deob simulate' for safe HTA/plain-text deobfuscation.")

    (options, args) = parser.parse_args()

    if (options.print_version):
        print_version()
        sys.exit(0)
    
    if len(args) == 0:
        parser.print_help()
        sys.exit(0)

    colorlog.basicConfig(level=LOG_LEVELS[options.loglevel], format='%(log_color)s%(levelname)-8s %(message)s')

    json_results = []

    for container, filename, data in xglob.iter_files(args,
                                                      recursive=options.recursive,
                                                      zip_password=options.zip_password,
                                                      zip_fname=options.zip_fname):

        if container and filename.endswith('/'):
            continue
        if options.deob is not None:
            if options.deob.lower() != "simulate":
                log.error("Unsupported --deob mode: " + str(options.deob))
                sys.exit(2)
            entry_points = None
            if (options.entry_points is not None):
                entry_points = options.entry_points.split(",")
            _process_deob_simulate_input(filename, data, entry_points=entry_points)
        elif options.scan_expressions:
            process_file_scanexpr(container, filename, data)
        else:
            entry_points = None
            if (options.entry_points is not None):
                entry_points = options.entry_points.split(",")
            process_file(container,
                         filename,
                         data,
                         strip_useless=options.strip_useless_code,
                         entry_points=entry_points,
                         time_limit=options.time_limit,
                         display_int_iocs=options.display_int_iocs,
                         tee_log=options.tee_log,
                         tee_bytes=options.tee_bytes,
                         out_file_name=options.out_file,
                         do_jit=options.do_jit)

            if (options.out_file):
                with open(options.out_file, 'r') as json_file:
                    try:
                        json_results.append(json.loads(json_file.read()))
                    except ValueError:
                        pass

    if (options.out_file):
        with open(options.out_file, 'w') as json_file:
            if (len(json_results) > 1):
                json_file.write(json.dumps(json_results, indent=2))
            else:
                json_file.write(json.dumps(json_results[0], indent=2))

        log.info("Saved results JSON to output file " + options.out_file)


if __name__ == '__main__':
    main()
