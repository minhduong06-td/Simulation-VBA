#!/usr/bin/env python3


import psutil
import subprocess
import time
import argparse
import json
import os
import signal

from unotools import Socket, connect
from unotools.component.writer import Writer
from unotools.unohelper import convert_path_to_url
from unotools import ConnectionError

HOST = "127.0.0.1"
PORT = 2002

def is_word_file(fname):
    typ = subprocess.check_output(["file", fname])
    return ((b"Microsoft Office Word" in typ) or
            (b"Word 2007+" in typ) or
            (b"Microsoft OOXML" in typ))

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

def get_document(fname, connection):
    url = convert_path_to_url(fname)
    document = Writer(connection, url)
    return document

def get_text(document):
    return "\x0c" + str(document.getText().getString())

def get_tables(document):
    data_array_list = []

    text_tables = document.getTextTables()
    table_count = 0
    while table_count < text_tables.getCount():
        data_array_list.append(text_tables.getByIndex(table_count).getDataArray())
        table_count += 1

    return data_array_list


if __name__ == '__main__':
    arg_parser = argparse.ArgumentParser(description="export text from various properties in a Word "
                                         "document via the LibreOffice API")
    arg_parser.add_argument("--tables", action="store_true",
                            help="export a list of 2D lists containing the cell contents"
                            "of each text table in the document")
    arg_parser.add_argument("--text", action="store_true",
                            help="export a string containing the document text")
    arg_parser.add_argument("-f", "--file", action="store", required=True,
                            help="path to the word doc")
    args = arg_parser.parse_args()

    if (not is_word_file(args.file)):

        exit()

    run_soffice()

    connection = connect(Socket(HOST, PORT))

    document = get_document(args.file, connection)

    if args.text:
        print((get_text(document)))
    elif args.tables:
        print((json.dumps(get_tables(document))))

    document.close(True)
    os.kill(get_office_proc()["pid"], signal.SIGTERM)
