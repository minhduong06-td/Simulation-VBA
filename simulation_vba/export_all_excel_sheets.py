#!/usr/bin/env python3
import sys
import os
import signal
import psutil
import subprocess
import time
import codecs
import string

from unotools import Socket, connect
from unotools.component.calc import Calc
from unotools.unohelper import convert_path_to_url
from unotools import ConnectionError

sys.stdout = codecs.getwriter("utf-8")(sys.stdout.detach())

HOST = "127.0.0.1"
PORT = 2002

def strip_unprintable(the_str):
    r = the_str
    if ((isinstance(r, str)) or (not isinstance(r, bytes))):
        r = ''.join(filter(lambda x:x in string.printable, r))
        
    else:
        tmp_r = ""
        for char_code in filter(lambda x:chr(x) in string.printable, r):
            tmp_r += chr(char_code)
        r = tmp_r

    return r

def to_str(s):
    if (isinstance(s, bytes)):
        try:
            return s.decode()
        except UnicodeDecodeError:
            return strip_unprintable(s)
    return s

def is_excel_file(maldoc):
    typ = subprocess.check_output(["file", maldoc])
    if ((b"Excel" in typ) or (b"Microsoft OOXML" in typ)):
        return True
    typ = subprocess.check_output(["exiftool", maldoc])
    return (b"vnd.ms-excel" in typ)

def wait_for_uno_api():
    tries = 0
    while tries < 3:
        try:
            connect(Socket(HOST, PORT))
            return
        except ConnectionError:
            time.sleep(5)
            tries += 1

    raise Exception("libreoffice UNO API failed to start")

def get_office_proc():
    for proc in psutil.process_iter():
        try:
            pinfo = proc.as_dict(attrs=['pid', 'name', 'username'])
        except psutil.NoSuchProcess:
            pass
        else:
            if (pinfo["name"].startswith("soffice")):
                return pinfo
    return None

def is_office_running():
    return True if get_office_proc() else False

def run_soffice():
    if not is_office_running():

        cmd = "/usr/lib/libreoffice/program/soffice.bin --headless --invisible " + \
              "--nocrashreport --nodefault --nofirststartwizard --nologo " + \
              "--norestore " + \
              '--accept="socket,host=127.0.0.1,port=2002,tcpNoDelay=1;urp;StarOffice.ComponentContext"'
        subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, shell=True)
        wait_for_uno_api()

def get_component(fname, context):
    url = convert_path_to_url(fname)
    component = Calc(context, url)
    return component

def fix_file_name(fname):
    r = ""
    for c in fname:
        if ((ord(c) < 48) or (ord(c) > 122)):
            r += hex(ord(c))
            continue
        r += c

    return r

def convert_csv(fname):
    if (not is_excel_file(fname)):

        return []

    run_soffice()
    
    
    context = connect(Socket(HOST, PORT))

    component = get_component(fname, context)

    r = []
    controller = component.getCurrentController()
    active_sheet = controller.ActiveSheet
    active_sheet_name = "NO_ACTIVE_SHEET"
    if (active_sheet is not None):
        active_sheet_name = fix_file_name(active_sheet.getName())
    r.append(active_sheet_name)
        
    sheets = component.getSheets()
    enumeration = sheets.createEnumeration()
    pos = 0
    if sheets.getCount() > 0:
        while enumeration.hasMoreElements():

            sheet = enumeration.nextElement()
            name = sheet.getName()
            if (name.count(" ") > 10):
                name = name.replace(" ", "")
            name = fix_file_name(name)
            controller.setActiveSheet(sheet)

            short_name = fname
            if (os.path.sep in short_name):
                short_name = short_name[short_name.rindex(os.path.sep) + 1:]
            short_name = fix_file_name(short_name)
            outfilename =  "/tmp/sheet_%s-%s--%s.csv" % (short_name, str(pos), name.replace(' ', '_SPACE_'))
            pos += 1
            r.append(outfilename)
            url = convert_path_to_url(outfilename)

            component.store_to_url(url,'FilterName','Text - txt - csv (StarCalc)')

    component.close(True)

    os.kill(get_office_proc()["pid"], signal.SIGTERM)

    return r

if __name__ == '__main__':
    r = to_str(str(convert_csv(sys.argv[1])))
    print(r)
