


magic_nums = {
    "office97" : "D0 CF 11 E0 A1 B1 1A E1",
    "office2007" : "50 4B 3 4",
}

pe_magic_num = "4D 5A"

def get_1st_8_bytes(fname, is_data):

    info = None
    is_data = (is_data or (len(fname) > 200))
    if (not is_data):
        try:
            tmp = open(fname, 'rb')
            tmp.close()
        except:
            is_data = True
    if (not is_data):
        with open(fname, 'rb') as f:
            info = f.read(8)
    else:
        info = fname[:9]

    curr_magic = ""
    for b in info:
        curr_magic += hex(ord(b)).replace("0x", "").upper() + " "
        
    return curr_magic

def is_pe_file(fname, is_data):
    """
    Check to see if the given file is a PE executable.

    return - True if it is a PE file, False if not.
    """

    curr_magic = get_1st_8_bytes(fname, is_data)

    return (curr_magic.startswith(pe_magic_num))

def is_office_file(fname, is_data):
    """
    Check to see if the given file is a MS Office file format.

    return - True if it is an Office file, False if not.
    """

    curr_magic = get_1st_8_bytes(fname, is_data)

    for typ in magic_nums.keys():
        magic = magic_nums[typ]
        if (curr_magic.startswith(magic)):
            return True
    return False

def is_office97_file(fname, is_data):

    curr_magic = get_1st_8_bytes(fname, is_data)

    return (curr_magic.startswith(magic_nums["office97"]))

def is_office2007_file(fname, is_data):

    curr_magic = get_1st_8_bytes(fname, is_data)

    return (curr_magic.startswith(magic_nums["office2007"]))
