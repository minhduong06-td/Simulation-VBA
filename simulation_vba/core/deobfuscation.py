import re
from functools import reduce

try:
    import regex
    REGEX = True
except ImportError:
    REGEX = False

from operator import xor

try:
    from simulation_vba.core.vba_lines import vba_collapse_long_lines
except Exception:
    try:
        from vba_lines import vba_collapse_long_lines
    except Exception:
        def vba_collapse_long_lines(code):
            lines = _to_text(code).splitlines(True)
            collapsed = []
            pending = ''
            for line in lines:
                stripped = line.rstrip('\r\n')
                if stripped.rstrip().endswith('_'):
                    pending += stripped.rstrip()[:-1]
                    continue
                collapsed.append(pending + line)
                pending = ''
            if pending:
                collapsed.append(pending)
            return ''.join(collapsed)

try:
    from simulation_vba.core.logger import log
except Exception:
    try:
        from logger import log
    except Exception:
        import logging
        log = logging.getLogger(__name__)



if REGEX:
    CHR = regex.compile(r'Chr\((?P<op>\d+)(\s+Xor\s+(?P<op>\d+))*\)', regex.IGNORECASE)
    STRING = regex.compile('(".*?"|\'.*?.\')')

    CONCAT_RUN = regex.compile(
        r'(?P<entry>{chr}|{string})(\s*[&+]\s*(?P<entry>{chr}|{string}))*'.format(
            chr=CHR.pattern, string=STRING.pattern))

    VAR_RUN = regex.compile(r'''
        (?P<var>[A-Za-z][A-Za-z0-9]*)\s*?=\s*(?P<entry>.*?)[\r\n]     # variable = *
        (\s*?(?P=var)\s*=\s*(?P=var)\s+&\s+(?P<entry>.*?)[\r\n])+     # variable = variable & *
    ''', regex.VERBOSE)

    def _replace_code(code, replacements):
        new_code = ''
        index = 0
        for start, end, code_string in sorted(replacements):
            new_code += code[index:start] + code_string
            index = end
        new_code += code[index:]
        return new_code


    def _replace_var_runs(code):
        code_replacements = []
        for match in VAR_RUN.finditer(code):
            code_string = '{var} = {value}{newline}'.format(
                var=match.group('var'),
                value=' & '.join(match.captures('entry')),
                newline=match.group(0)[-1]
            )
            code_replacements.append((match.start(), match.end(), code_string))
        return _replace_code(code, code_replacements)



    def _replace_concat_runs(code):
        code_replacements = []
        for match in CONCAT_RUN.finditer(code):
            code_string = ''
            for entry in match.captures('entry'):
                sub_match = CHR.match(entry)
                if sub_match:
                    character = chr(reduce(xor, map(int, sub_match.captures('op'))))
                    if character == '"':
                        character = '""'
                    code_string += character
                else:
                    code_string += entry.strip('\'"')
            code_replacements.append((match.start(), match.end(), '"{}"'.format(code_string)))
        return _replace_code(code, code_replacements)


def deobfuscate(code):
    code = vba_collapse_long_lines(code)
    if REGEX:
        code = _replace_var_runs(code)
        code = _replace_concat_runs(code)
    else:
        log.debug('regex package is not installed - using limited deobfuscation')
    code = _replace_simple_concat_runs(code)
    return code


def _to_text(data):
    if isinstance(data, bytes):
        for encoding in ('utf-8', 'latin-1'):
            try:
                return data.decode(encoding)
            except UnicodeDecodeError:
                pass
        return data.decode('latin-1', 'replace')
    return data


def extract_vb_from_hta(code):
    code = _to_text(code)
    if (re.search(r"&#\d{1,3};", code) is not None):
        for i in range(0, 256):
            code = code.replace("&#" + str(i) + ";", chr(i))

    hta_regexes = [
        r"<\s*script\b[^>]*(?:language|type)\s*=\s*[\"']?[^>]*vbscript[^>]*>(.*?)</\s*script\s*>",
        r"<\s*script\s+\%\d{1,10}\s*>(.*?)</\s*script\s*>",
        r"<\s*script\b[^>]*(?:language|type)\s*=\s*[\"']?[^>]*vbscript[^>]*>(.*)$",
    ]
    blocks = []
    for pattern in hta_regexes:
        blocks = re.findall(pattern, code.strip(), re.IGNORECASE | re.DOTALL)
        if blocks:
            break
    if not blocks:
        return code

    cleaned = []
    for block in blocks:
        block = block.strip()
        lower_block = block.lower()
        if "</script>" in lower_block:
            block = block[:lower_block.index("</script>")]
        upper_block = block.upper()
        if "<![CDATA[" in upper_block:
            block = block[upper_block.index("<![CDATA[") + len("<![CDATA["):]
            if "]]>" in block[-10:]:
                block = block[:block.rindex("]]>")]
        comment_match = re.findall(r"<!\-\-(.+?)/?/?\-\->", block, re.DOTALL)
        if comment_match:
            block = comment_match[0].strip()
        if block.endswith("//"):
            block = block[:-2]
        cleaned.append(block)
    return "\n".join(cleaned) + "\n"


_PROC_START = re.compile(
    r"^\s*(?:(?:public|private|global|friend|static)\s+)*"
    r"(sub|function|property\s+(?:get|let|set))\s+([A-Za-z_][A-Za-z0-9_]*)\b",
    re.IGNORECASE)

_PROC_END = re.compile(r"^\s*end\s+(sub|function|property)\b", re.IGNORECASE)

_AUTO_ENTRY_POINTS = set([
    'autoopen', 'document_open', 'autoclose', 'document_close', 'auto_open',
    'autoexec', 'autoexit', 'document_beforeclose', 'workbook_open',
    'workbook_activate', 'auto_close', 'workbook_close',
    'workbook_deactivate', 'documentopen', 'app_documentopen', 'main'
])


def _split_vb_args(text):
    args = []
    current = ''
    in_string = False
    parens = 0
    i = 0
    while i < len(text):
        char = text[i]
        if char == '"':
            current += char
            if in_string and i + 1 < len(text) and text[i + 1] == '"':
                current += text[i + 1]
                i += 2
                continue
            in_string = not in_string
        elif (char == '(') and not in_string:
            parens += 1
            current += char
        elif (char == ')') and not in_string:
            parens -= 1
            current += char
        elif (char == ',') and not in_string and parens == 0:
            args.append(current.strip())
            current = ''
        else:
            current += char
        i += 1
    if current.strip():
        args.append(current.strip())
    return args


def _split_vb_concat(text):
    parts = []
    current = ''
    in_string = False
    i = 0
    while i < len(text):
        char = text[i]
        if char == '"':
            current += char
            if in_string and i + 1 < len(text) and text[i + 1] == '"':
                current += text[i + 1]
                i += 2
                continue
            in_string = not in_string
        elif (char in ('&', '+')) and not in_string:
            if current.strip():
                parts.append(current.strip())
            current = ''
        else:
            current += char
        i += 1
    if current.strip():
        parts.append(current.strip())
    return parts


def _strip_inline_comment(line):
    in_string = False
    i = 0
    while i < len(line):
        char = line[i]
        if char == '"':
            if in_string and i + 1 < len(line) and line[i + 1] == '"':
                i += 2
                continue
            in_string = not in_string
        elif char == "'" and not in_string:
            return line[:i]
        i += 1
    return line


def _vb_string_literal(value):
    value = value.strip()
    if len(value) >= 2 and value[0] == '"' and value[-1] == '"':
        return value[1:-1].replace('""', '"')
    return None


def _eval_vb_string_expr(expr, variables):
    expr = _strip_inline_comment(expr).strip()
    if not expr:
        return ''
    parts = _split_vb_concat(expr)
    if len(parts) > 1:
        return ''.join(_eval_vb_string_expr(part, variables) for part in parts)
    literal = _vb_string_literal(expr)
    if literal is not None:
        return literal
    chr_match = re.match(r"Chr[BW]?\$?\s*\(\s*(\d+)(?:\s+Xor\s+(\d+))*\s*\)\s*$", expr, re.IGNORECASE)
    if chr_match:
        nums = [int(num) for num in re.findall(r"\d+", expr)]
        return chr(reduce(xor, nums))
    lower = expr.lower()
    if lower in variables:
        return variables[lower]
    return expr


_SIMPLE_CONCAT_RUN = re.compile(
    r'(?:Chr[BW]?\$?\s*\(\s*\d+(?:\s+Xor\s+\d+)*\s*\)|"(?:""|[^"])*"|\'[^\']*\')'
    r'(?:\s*[&+]\s*(?:Chr[BW]?\$?\s*\(\s*\d+(?:\s+Xor\s+\d+)*\s*\)|"(?:""|[^"])*"|\'[^\']*\'))+',
    re.IGNORECASE)


def _replace_simple_concat_runs(code):
    def replace_match(match):
        value = _eval_vb_string_expr(match.group(0), {})
        return '"' + value.replace('"', '""') + '"'
    return _SIMPLE_CONCAT_RUN.sub(replace_match, code)


def _extract_call_argument(line, func_name):
    pattern = re.compile(r"\b" + re.escape(func_name) + r"\b\s*(?:\((.*)\)|(.*))", re.IGNORECASE)
    match = pattern.search(line)
    if not match:
        return None
    raw = match.group(1) if match.group(1) is not None else match.group(2)
    if raw is None:
        return ''
    raw = raw.strip()
    if raw.lower().startswith('call '):
        raw = raw[5:].strip()
    args = _split_vb_args(raw)
    return args[0] if args else raw


def _collect_procedures(code):
    procedures = {}
    ordered_lines = code.splitlines(True)
    i = 0
    while i < len(ordered_lines):
        line = ordered_lines[i]
        start = _PROC_START.match(line)
        if not start:
            i += 1
            continue
        proc_type = start.group(1).lower()
        proc_name = start.group(2)
        header = line
        body = []
        i += 1
        while i < len(ordered_lines):
            curr = ordered_lines[i]
            end = _PROC_END.match(curr)
            if end:
                footer = curr
                break
            body.append(curr)
            i += 1
        else:
            footer = ''
        procedures[proc_name.lower()] = {
            'name': proc_name,
            'type': proc_type,
            'header': header,
            'body': body,
            'footer': footer,
        }
        i += 1
    return procedures


def _looks_like_proc_call(line, proc_name):
    line = _strip_inline_comment(line).strip()
    if not line:
        return False
    pattern = re.compile(r"^(?:call\s+)?" + re.escape(proc_name) + r"\b(?:\s*\(\s*\))?\s*$", re.IGNORECASE)
    return pattern.match(line) is not None


def _record_action(actions, action, params, description):
    actions.append((action, params, description))


def _simulate_lines(lines, procedures, actions, variables, visited):
    for raw_line in lines:
        line = _strip_inline_comment(raw_line).strip()
        if not line:
            continue

        assign = re.match(r"^(?:set\s+)?([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.+)$", line, re.IGNORECASE)
        if assign:
            variables[assign.group(1).lower()] = _eval_vb_string_expr(assign.group(2), variables)

        create_matches = re.findall(r"\bCreateObject\s*\(\s*([^)]+)\)", line, re.IGNORECASE)
        for raw_obj in create_matches:
            obj_name = _eval_vb_string_expr(raw_obj, variables)
            _record_action(actions, 'Create Object', obj_name, 'CreateObject stubbed')

        shell_arg = _extract_call_argument(line, 'Shell')
        if shell_arg is not None:
            _record_action(actions, 'Execute Command', _eval_vb_string_expr(shell_arg, variables), 'Shell function stubbed')

        run_match = re.search(r"\.run\b", line, re.IGNORECASE)
        if run_match:
            run_arg = line[run_match.end():].strip()
            if run_arg.startswith('(') and run_arg.endswith(')'):
                run_arg = run_arg[1:-1]
            args = _split_vb_args(run_arg)
            command = args[0] if args else run_arg
            _record_action(actions, 'Execute Command', _eval_vb_string_expr(command, variables), 'WScript.Shell.Run() stubbed')

        lower = line.lower()
        dangerous_tokens = [
            ('File Delete', ['kill ', 'rmdir ', '.deletefile', '.deletefolder']),
            ('File Write', ['createtextfile', 'opentextfile', 'savetofile', 'writefile', 'print #', 'write #']),
            ('Network', ['urldownloadtofile', 'xmlhttp', 'winhttprequest', 'internetopen', 'internetconnect']),
            ('Registry', ['regwrite', 'regdelete', 'savesetting', 'deletesetting']),
            ('Subprocess', ['createprocess', 'shellexecute']),
        ]
        for action, tokens in dangerous_tokens:
            if any(token in lower for token in tokens):
                _record_action(actions, action, line, action + ' stubbed')

        for proc_name, proc in list(procedures.items()):
            if proc_name in visited:
                continue
            if _looks_like_proc_call(line, proc['name']):
                visited.add(proc_name)
                _simulate_lines(proc['body'], procedures, actions, variables, visited)


def simulate_deobfuscation(code, entry_points=None):
    vb_code = extract_vb_from_hta(code)
    deobfuscated = deobfuscate(vb_code)
    procedures = _collect_procedures(deobfuscated)
    actions = []
    variables = {}

    if entry_points is None:
        selected = [name for name in procedures if name in _AUTO_ENTRY_POINTS]
        if not selected:
            selected = [
                name for name, proc in list(procedures.items())
                if proc['type'] == 'sub' and '(' in proc['header'] and ')' in proc['header']
            ]
    else:
        selected = [name.lower() for name in entry_points]

    visited = set()
    for name in selected:
        proc = procedures.get(name.lower())
        if proc is None or name.lower() in visited:
            continue
        visited.add(name.lower())
        _simulate_lines(proc['body'], procedures, actions, variables, visited)

    return deobfuscated, actions
