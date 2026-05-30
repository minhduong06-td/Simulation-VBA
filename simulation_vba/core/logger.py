#!/usr/bin/env python
"""
simulation_vba logging helper

SimulationVBA is a specialized engine to parse, analyze and interpret Microsoft
VBA macros (Visual Basic for Applications), mainly for malware analysis.

Author: Philippe Lagadec - http://www.decalage.info
License: BSD, see source code or documentation

Project Repository:
https://github.com/decalage2/ViperMonkey
"""

# === LICENSE ==================================================================




__version__ = '0.08'



import logging


class CappedFileHandler(logging.FileHandler):

    def __init__(self, filename, sizecap, mode='w', encoding=None, delay=False):
        self.size_cap = sizecap
        self.current_size = 0
        self.cap_exceeded = False
        super(CappedFileHandler, self).__init__(filename, mode, encoding, delay)

    def emit(self, record):
        if not self.cap_exceeded:
            new_size = self.current_size + len(self.formatter.format(record))
            if new_size <= self.size_cap:
                self.current_size = new_size
                super(CappedFileHandler, self).emit(record)
            else:
                self.cap_exceeded = True

class DuplicateFilter(logging.Filter):

    def filter(self, record):
        current_log = (record.module, record.levelno, record.msg)
        if current_log != getattr(self, "last_log", None):
            self.last_log = current_log
            return True
        return False

def get_logger(name, level=logging.NOTSET):
    """
    Create a suitable logger object for this module.
    The goal is not to change settings of the root logger, to avoid getting
    other modules' logs on the screen.
    If a logger exists with same name, reuse it. (Else it would have duplicate
    handlers and messages would be doubled.)
    """
    if name in logging.Logger.manager.loggerDict:
        logger = logging.getLogger(name)
        logger.setLevel(level)
        logger.addFilter(DuplicateFilter()) 
        return logger
    logger = logging.getLogger(name)
    logger.addHandler(logging.NullHandler())
    logger.setLevel(level)
    logger.addFilter(DuplicateFilter()) 
    return logger


log = get_logger('VMonkey')

