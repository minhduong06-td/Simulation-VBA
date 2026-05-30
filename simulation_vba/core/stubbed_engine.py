


__version__ = '0.02'

import unidecode
import string

import logging
from logger import log

class StubbedEngine(object):
    """
    Stubbed out Vipermonkey analysis engine that just supports tracking
    actions.
    """

    def __init__(self):
        self.actions = []

    def report_action(self, action, params=None, description=None):
        """
        Callback function for each evaluated statement to report macro actions
        """

        try:
            if (isinstance(action, str)):
                action = unidecode.unidecode(action.decode('unicode-escape'))
        except UnicodeDecodeError:
            action = ''.join(filter(lambda x:x in string.printable, action))
        if (isinstance(params, str)):
            try:
                decoded = params.replace("\\", "#ESCAPED_SLASH#").decode('unicode-escape').replace("#ESCAPED_SLASH#", "\\")
                params = unidecode.unidecode(decoded)
            except Exception as e:
                log.warn("Unicode decode of action params failed. " + str(e))
                params = ''.join(filter(lambda x:x in string.printable, params))
        try:
            if (isinstance(description, str)):
                description = unidecode.unidecode(description.decode('unicode-escape'))
        except UnicodeDecodeError as e:
            log.warn("Unicode decode of action description failed. " + str(e))
            description = ''.join(filter(lambda x:x in string.printable, description))
        self.actions.append((action, params, description))
        log.info("ACTION: %s - params %r - %s" % (action, params, description))
