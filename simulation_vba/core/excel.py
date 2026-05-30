
# pylint: disable=pointless-string-statement
__version__ = '0.03'

from logger import log
import logging
import json
import os
import filetype
import random
import re
import subprocess
try:
    import xlrd2 as xlrd
except ImportError:
    log.warning("xlrd2 Python package not installed. Falling back to xlrd.")
    import xlrd

import utils
    
_thismodule_dir = os.path.normpath(os.path.abspath(os.path.dirname(__file__)))
    
debug = False

def _read_sheet_from_csv(filename):
    f = None
    try:
        f = open(filename, 'r')
    except Exception as e:
        log.error("Cannot open CSV file. " + str(e))
        return None

    data = f.read()
    f.close()
    in_str = False
    tmp = ""
    for c in data:
        if (c == '"'):
            in_str = not in_str
        if (in_str and (c == ',')):
            tmp += "#A_COMMA!!#"
        elif (in_str and (c == '\n')):
            tmp += "#A_NEWLINE!!#"
        else:
            tmp += c
    data = tmp
    
    row = 0
    r = {}
    for line in data.split("\n"):

        line = line.strip()
        cells = line.split(",")
        col = 0
        for cell in cells:

            cell = cell.replace("#A_COMMA!!#", ",").replace("#A_NEWLINE!!#", "\n")
            
            dat = str(cell)
            if (dat.startswith('"')):
                dat = dat[1:]
            if (dat.endswith('"')):
                dat = dat[:-1]

            dat = dat.replace('""', '"')
            
            r[(row, col)] = dat

            col += 1
        row += 1

    r = make_book(r)
    return r

def _fix_sheet_name(sheet_name):
    pat = r"(0x[0-9a-f]{2})"
    r = utils.safe_str_convert(sheet_name)
    hex_strs = re.findall(pat, r)
    if (len(hex_strs) == 0):
        return sheet_name

    for hex_val in hex_strs:
        try:
            chr_val = int(hex_val, 16)
            r = r.replace(hex_val, chr(chr_val))
        except Exception as e:
            log.error("Fixing sheet named failed. " + str(e))
    return r

def load_excel_libreoffice(data):
    if (not filetype.is_office_file(data, True)):
        log.warning("The file is not an Office file. Not extracting sheets with LibreOffice.")
        return None
    
    out_dir = "/tmp/tmp_excel_file_" + str(random.randrange(0, 10000000000))
    f = open(out_dir, 'wb')
    f.write(data)
    f.close()
    
    output = None
    try:
        output = subprocess.check_output(["timeout", "30", "python3", _thismodule_dir + "/../export_all_excel_sheets.py", out_dir])
    except Exception as e:
        log.error("Running export_all_excel_sheets.py failed. " + str(e))
        os.remove(out_dir)
        return None

    try:
        sheet_files = json.loads(output.replace("'", '"'))
    except Exception as e:
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Loading sheet file names failed. " + str(e))
        os.remove(out_dir)
        return None

    if (len(sheet_files) <= 1):
        os.remove(out_dir)
        return None

    active_sheet_name = _fix_sheet_name(sheet_files[0])
    
    sheet_map = {}
    for sheet_file in sheet_files[1:]:

        tmp_workbook = _read_sheet_from_csv(sheet_file)

        cell_data = tmp_workbook.sheet_by_name("Sheet1").cells
        
        start = sheet_file.index("--") + 2
        end = sheet_file.rindex(".")
        sheet_name = _fix_sheet_name(sheet_file[start : end])

        start = sheet_file.index("-") + 1
        end = sheet_file[start:].index("-") + start
        sheet_index = int(sheet_file[start : end])
        
        tmp_sheet = ExcelSheet(cell_data, sheet_name)

        sheet_map[sheet_index] = tmp_sheet

    result_book = ExcelBook(None)
    sorted_indices = list(sheet_map.keys())
    sorted_indices.sort()
    for index in sorted_indices:
        result_book.sheets.append(sheet_map[index])

    if (active_sheet_name != "NO_ACTIVE_SHEET"):
        result_book.active_sheet_name = active_sheet_name
        
    for sheet_file in sheet_files[1:]:
        os.remove(sheet_file)

    if os.path.isfile(out_dir):
        os.remove(out_dir)
        
    return result_book
        
def load_excel_xlrd(data):
    if (not filetype.is_office97_file(data, True)):
        log.warning("File is not an Excel 97 file. Not reading with xlrd2.")
        return None

    try:
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Trying to load with xlrd...")
        r = xlrd.open_workbook(file_contents=data)
        return r
    except Exception as e:
        log.error("Reading in file as Excel with xlrd failed. " + str(e))
        return None

def load_excel(data):
    wb = load_excel_libreoffice(data)
    if (wb is not None):

        if (len(wb.sheet_names()) > 0):
            return wb

    wb = load_excel_xlrd(data)
    if (wb is not None):
        return wb

    return None

def is_cell_dict(x):
    return (isinstance(x, dict) and ("value" in x))

def _get_alphanum_cell_index(row, col):
    dividend = col
    column_name = ""
    modulo = 0
    while (dividend > 0):
        modulo = (dividend - 1) % 26
        column_name = chr(65 + modulo) + column_name
        dividend = int((dividend - modulo) / 26)

    return column_name + str(row)
    
def get_largest_sheet(workbook):
    if (hasattr(workbook, "__largest_sheet__")):
        return workbook.__largest_sheet__
    
    cells = []
    big_sheet = None
    for sheet_index in range(0, len(workbook.sheet_names())):
        
        sheet = None
        try:
            sheet = workbook.sheet_by_index(sheet_index)
        # pylint: disable=bare-except
        except:
            return None

        curr_cells = pull_cells_sheet(sheet, strip_empty=True)
        if (curr_cells is None):
            curr_cells = []
                    
        if (len(curr_cells) > len(cells)):
            cells = curr_cells
            big_sheet = sheet

    workbook.__largest_sheet__ = big_sheet
    return big_sheet

def get_num_rows(sheet):
    if (hasattr(sheet, "num_rows")):
        return sheet.num_rows()

    if (hasattr(sheet, "nrows")):
        return sheet.nrows

    return 0

def get_num_cols(sheet):
    if (hasattr(sheet, "num_cols")):
        return sheet.num_cols()

    if (hasattr(sheet, "ncols")):
        return sheet.ncols

    return 0

def _pull_cells_sheet_xlrd(sheet, strip_empty):
    if (not hasattr(sheet, "nrows") or
        not hasattr(sheet, "ncols")):
        return None
    max_row = sheet.nrows
    max_col = sheet.ncols

    curr_cells = []
    for curr_row in range(0, max_row + 1):
        for curr_col in range(0, max_col + 1):
            try:
                curr_cell_xlrd = sheet.cell(curr_row, curr_col)
                curr_val = curr_cell_xlrd.value
                if (strip_empty and (len(str(curr_val).strip()) == 0)):
                    continue
                curr_cell = { "value" : curr_val,
                              "row" : curr_row + 1,
                              "col" : curr_col + 1,
                              "index" : _get_alphanum_cell_index(curr_row, curr_col) }
                curr_cells.append(curr_cell)
            # pylint: disable=bare-except
            except:
                pass

    return curr_cells
            
def _pull_cells_sheet_internal(sheet, strip_empty):
    if (not hasattr(sheet, "cells")):
        return None
        

    max_row = -1
    max_col = -1
    for cell_index in sheet.cells.keys():
        curr_row = cell_index[0]
        curr_col = cell_index[1]
        if (curr_row > max_row):
            max_row = curr_row
        if (curr_col > max_col):
            max_col = curr_col

    curr_cells = []
    for curr_row in range(0, max_row + 1):
        for curr_col in range(0, max_col + 1):
            try:
                curr_val = sheet.cell(curr_row, curr_col)
                if (strip_empty and (len(str(curr_val).strip()) == 0)):
                    continue
                curr_cell = { "value" : curr_val,
                              "row" : curr_row + 1,
                              "col" : curr_col + 1,
                              "index" : _get_alphanum_cell_index(curr_row, curr_col) }
                curr_cells.append(curr_cell)
            except KeyError:
                pass

    return curr_cells

def pull_cells_sheet(sheet, strip_empty=False):
    curr_cells = _pull_cells_sheet_xlrd(sheet, strip_empty)
    if (curr_cells is None):
        curr_cells = _pull_cells_sheet_internal(sheet, strip_empty)
    return curr_cells
    
def pull_cells_workbook(workbook):
    all_cells = []
    for sheet_index in range(0, len(workbook.sheet_names())):
            
        sheet = None
        try:
            sheet = workbook.sheet_by_index(sheet_index)
        # pylint: disable=bare-except
        except:
            continue

        curr_cells = pull_cells_sheet(sheet)
        if (curr_cells is None):
            continue
        all_cells.extend(curr_cells)

    return all_cells

class ExcelSheet(object):
    def __init__(self, cells, name="Sheet1"):
        self.gloss = None
        self.cells = cells
        self.name = name.replace("0x20", " ")
        self.__num_rows = None
        self.__num_cols = None

    def __repr__(self):
        if (self.gloss is not None):
            return self.gloss
        log.info("Converting Excel sheet to str ...")
        r = ""
        if debug:
            r += "Sheet: '" + self.name + "'\n\n"
            for cell in self.cells.keys():
                r += str(cell) + "\t=\t'" + str(self.cells[cell]) + "'\n"
        else:
            r += "Sheet: '" + self.name + "'\n"
            r += str(self.cells)
        self.gloss = r
        return self.gloss

    def num_rows(self):
        if (self.__num_rows is not None):
            return self.__num_rows
        max_row = -1
        for cell in self.cells.keys():
            curr_row = cell[0]
            if (curr_row > max_row):
                max_row = curr_row
        self.__num_rows = max_row
        return self.__num_rows

    def num_cols(self):
        if (self.__num_cols is not None):
            return self.__num_cols
        max_col = -1
        for cell in self.cells.keys():
            curr_col = cell[1]
            if (curr_col > max_col):
                max_col = curr_col
        self.__num_cols = max_col
        return self.__num_cols
    
    def cell(self, row, col):
        if ((row, col) in self.cells):
            return self.cells[(row, col)]
        raise KeyError("Cell (" + str(row) + ", " + str(col) + ") not found.")

    def cell_value(self, row, col):
        return self.cell(row, col)

    def cell_dict(self, row, col):
        curr_cell = { "value" : self.cell(row, col),
                      "row" : row + 1,
                      "col" : col + 1,
                      "index" : _get_alphanum_cell_index(row, col) }
        return curr_cell
    
class ExcelBook(object):
    def __init__(self, cells=None, name="Sheet1"):
        self.sheets = []
        if (cells is None):
            return

        self.sheets.append(ExcelSheet(cells, name))

    def __repr__(self):
        """String version of workbook.

        """
        log.info("Converting Excel workbook to str ...")
        r = ""
        for sheet in self.sheets:
            r += str(sheet) + "\n"
        return r
        
    def sheet_names(self):
        r = []
        for sheet in self.sheets:
            r.append(sheet.name)
        return r

    def sheet_by_index(self, index):
        if (index < 0):
            raise ValueError("Sheet index " + str(index) + " is < 0")
        if (index >= len(self.sheets)):
            raise ValueError("Sheet index " + str(index) + " is > num sheets (" + str(len(self.sheets)) + ")")
        return self.sheets[index]

    def sheet_by_name(self, name):
        for sheet in self.sheets:
            if (sheet.name == name):
                return sheet
        raise ValueError("Sheet name '" + str(name) + "' not found.")

def make_book(cell_data):
    return ExcelBook(cell_data)
