#!/usr/bin/env python

import logging
import subprocess

from logger import log

class FakeMeta(object):
    pass

def get_metadata_exif(filename):

    output = None
    try:
        output = subprocess.check_output(["exiftool", filename])
    except Exception as e:
        log.error("Cannot read metadata with exiftool. " + str(e))
        return {}

    if (log.getEffectiveLevel() == logging.DEBUG):
        log.debug("exiftool output: '" + str(output) + "'")
    if (":" not in output):
        log.warning("Cannot read metadata with exiftool.")
        return {}
    
    lines = output.split("\n")
    r = FakeMeta()
    for line in lines:
        line = line.strip()
        if ((len(line) == 0) or (":" not in line)):
            continue        
        field = line[:line.index(":")].strip().lower()
        val = line[line.index(":") + 1:].strip().replace("...", "\r\n")
        setattr(r, field, val)

    return r
