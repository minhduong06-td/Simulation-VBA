import collections
import inspect
import pyparsing
pyparsing.ParserElement.enablePackrat(cache_size_limit=10000000)

from simulation_vba.core import deobfuscation
from simulation_vba.core.modules import *
from simulation_vba.core.modules import Module as _Module
from simulation_vba.core.vba_lines import vba_collapse_long_lines

from simulation_vba.core.vba_library import *


def _get_keywords(line, num=2):
    return line.lower().split(None, num)


class CustomVBALibraryFunc(VbaLibraryFunc):
    def __init__(self, callback):
        self._callback = callback
    def eval(self, context, params=None):
        return self._callback(context, params=params)


class SmartDict(dict):
    def __contains__(self, key):
        return super(SmartDict, self).__contains__(key.lower())

    def __getitem__(self, key):
        return super(SmartDict, self).__getitem__(key.lower())

    def __setitem__(self, key, value):
        if inspect.isclass(value) and issubclass(value, VbaLibraryFunc):
            value = value()
        elif callable(value):
            value = CustomVBALibraryFunc(value)
        super(SmartDict, self).__setitem__(key.lower(), value)
orig_Context = Context


class Context(orig_Context):
    def __init__(self, report_action=None, **kwargs):
        kwargs['engine'] = self
        super(Context, self).__init__(**kwargs)
        self._report_action = report_action
        self.actions = collections.defaultdict(list)

        if not isinstance(self.globals, SmartDict):
            self.globals = SmartDict(self.globals)
        if not isinstance(self.locals, SmartDict):
            self.locals = SmartDict(self.locals)

    def __contains__(self, item):
        try:
            _ = self[item]
            return True
        except KeyError:
            return False

    def __delitem__(self, key):
        key = key.lower()
        if key in self.locals:
            del self.locals[key]
        elif key in self.globals:
            del self.globals[key]
        if key in self.types:
            del self.types[key]

    def __getitem__(self, item):
        return self.get(item)

    def __setitem__(self, key, value):
        self.set(key, value)

    def report_action(self, action, params=None, description=None, strip_null_bytes=False):
        self.actions[description].append((action, params))
        if self._report_action:
            self._report_action(action, params=params, description=description)


class CodeBlock(object):
    def __init__(self, pp_spec, lines, parse_all=True, deobfuscate=False):
        self._pp_spec = pp_spec
        if isinstance(lines, (bytes, str)):
            if deobfuscate:
                lines = deobfuscation.deobfuscate(lines)
            else:
                lines = vba_collapse_long_lines(lines)
            self.lines = lines.splitlines(True)
        else:
            if deobfuscate:
                lines = deobfuscation.deobfuscate('\n'.join(lines)).splitlines(True)
            self.lines = lines
        self._obj = None
        self._parse_attempted = False
        self._parse_all = parse_all
        self._code_blocks = None

    def __str__(self):
        return ''.join(self.lines)

    def __getattr__(self, item):
        return getattr(self.obj, item, None)

    @property
    def __class__(self):
        return self.obj.__class__

    @property
    def obj(self):
        if not self._obj:
            if self._parse_attempted:
                return None
            try:
                self._parse_attempted = True
                self._obj = self._pp_spec.parseString(self.lines[0], parseAll=self._parse_all)[0]
            except ParseException as err:
                log.warn('*** PARSING ERROR (3) ***\n{}\n{}\n{}'.format(
                    err.line, " " * (err.column - 1) + "^", err))
                return None
        return self._obj

    def _take_until(self, line_gen, end):
        for line in line_gen:
            yield line
            if line.lower().split(None, len(end))[:len(end)] == end:
                return

    def _generate_code_block(self, line_gen, line, line_keywords):
        if line_keywords[0] == 'for':
            log.debug('FOR LOOP')
            lines = [line] + list(self._take_until(line_gen, ['next']))
            return CodeBlock(for_start, lines, parse_all=False)
        else:
            return CodeBlock(vba_line + Optional(EOS).suppress(), line)

    def _iter_code_blocks(self):
        line_gen = iter(self.lines[1:-1])
        for line in line_gen:
            log.debug('Parsing line: {}'.format(line.rstrip()))
            line_keywords = line.lower().split(None, 2)
            if not line_keywords or line_keywords[0].startswith("'"):
                continue
            if line_keywords[0] in ('public', 'private'):
                line_keywords = line_keywords[1:]

            yield self._generate_code_block(line_gen, line, line_keywords)

    @property
    def code_blocks(self):
        if self._code_blocks is None:
            code_blocks = []
            for code_block in self._iter_code_blocks():
                code_blocks.append(code_block)
                yield code_block
            self._code_blocks = code_blocks
        else:
            for code_block in self._code_blocks:
                yield code_block

    @property
    def type(self):
        return type(self.obj)

    def eval(self, context=None, params=None):
        context = context or Context()
        if not self.obj:
            log.error('Unable to evaluate "{}" due to parse error.'.format(self))
            return None
        if hasattr(self.obj, 'statements') and not self.obj.statements:
            self.obj.statements = list(self.code_blocks)
        if hasattr(self.obj, 'eval'):
            return self.obj.eval(context=context, params=params)
        else:
            return self.obj

    def load_context(self, context):
        for code_block in self.code_blocks:
            code_block.eval(context)


class Module(CodeBlock):
    _ENTRY_POINTS = ['autoopen', 'document_open', 'autoclose',
                     'document_close', 'auto_open', 'autoexec',
                     'autoexit', 'document_beforeclose', 'workbook_open',
                     'workbook_activate', 'auto_close', 'workbook_close']

    def __init__(self, lines, deobfuscate=False):
        super(Module, self).__init__(None, lines, deobfuscate=deobfuscate)
        self.lines = [''] + self.lines + ['']

    def _generate_code_block(self, line_gen, line, line_keywords):
        if line_keywords[0] == 'attribute':
            return CodeBlock(header_statements_line, line)
        elif line_keywords[0] in ('option', 'dim', 'declare'):
            log.debug('DECLARATION LINE')
            return CodeBlock(declaration_statements_line, line)
        elif line_keywords[0] == 'sub':
            log.debug('SUB')
            lines = [line] + list(self._take_until(line_gen, ['end', 'sub']))
            return CodeBlock(procedures.sub_start_line, lines)
        elif line_keywords[0] == 'function':
            log.debug('FUNCTION')
            lines = [line] + list(self._take_until(line_gen, ['end', 'function']))
            return CodeBlock(procedures.function_start_line, lines)
        else:
            return super(Module, self)._generate_code_block(line_gen, line, line_keywords)

    @property
    def functions(self):
        return list(self.obj.functions.values())

    @property
    def subs(self):
        return list(self.obj.subs.values())

    @property
    def procedures(self):
        return self.functions + self.subs

    @property
    def entry_points(self):
        for name, sub in self.obj.subs.items():
            if name.lower() in self._ENTRY_POINTS:
                yield sub
        for name, function in self.obj.functions.items():
            if name.lower() in self._ENTRY_POINTS:
                yield function

    def eval(self, context=None, params=None):
        context = context or Context()
        self.load_context(context)
        ret = None
        for code_block in self.code_blocks:
            if not isinstance(code_block, (Function, Sub)):
                ret = code_block.eval(context, params)
        return ret

    def load_context(self, context):
        context = context or Context()
        for name, _sub in list(self.obj.subs.items()):
            log.debug('(3) storing sub "%s" in globals' % name)
            context.globals[name.lower()] = _sub
        for name, _function in list(self.obj.functions.items()):
            log.debug('(3) storing function "%s" in globals' % name)
            context.globals[name.lower()] = _function
        for name, _function in list(self.obj.external_functions.items()):
            log.debug('(3) storing external function "%s" in globals' % name)
            context.globals[name.lower()] = _function
        for name, _var in list(self.obj.global_vars.items()):
            log.debug('(3) storing global var "%s" in globals' % name)
            if isinstance(name, str):
                context.globals[name.lower()] = _var
            if isinstance(name, list):
                context.globals[name[0].lower()] = _var
                context.types[name[0].lower()] = name[1]

    @property
    def obj(self):
        if not self._obj:
            self._obj = _Module(str(self), 0, list(self.code_blocks))
        return self._obj

def eval(vba_code, context=None, deobfuscate=False):
    context = context or Context()
    module = Module(vba_code, deobfuscate=deobfuscate)
    return module.eval(context)


def deobfuscate_simulate(vba_code, entry_points=None):
    return deobfuscation.simulate_deobfuscation(vba_code, entry_points=entry_points)
