


__version__ = '0.02'

class from_unicode_str(str):
    """
    Marker class to mark strings created by StrConv() with the
    vbaFromUnicode option. VipeMonkey currently assumes that unless
    specifically noted, all strings are unicode. This class is used to
    mark strings that are pure ascii, not unicode.
    """
    pass
