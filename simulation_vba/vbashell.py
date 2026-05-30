#!/usr/bin/env python

from __future__ import print_function
__version__ = '0.04'
import logging, optparse, sys, os
import colorlog
_thismodule_dir = os.path.normpath(os.path.abspath(os.path.dirname(__file__)))
if not _thismodule_dir in sys.path:
    sys.path.insert(0, _thismodule_dir)
import simulation_vba

vm = simulation_vba.SimulationVBA()
def parse(filename=None):
    if filename is None:
        print('Enter VBA code, end by a line containing only ".":')
        code = ''
        line = None
        while True:
            line = raw_input()
            if line == '.':
                break
            code += line + '\n'
    else:
        print('Parsing file %r' % filename)
        code = open(filename).read()
    vm.add_module(code)

def eval_expression(e):
    print('Evaluating %s' % e)
    value = vm.eval(e)
    print('Returned value: %s' % value)
    print('Recorded Actions:')
    print(vm.dump_actions())


def main():
    print ('vbashell %s - https://github.com/decalage2/ViperMonkey' % __version__)
    print ('THIS IS WORK IN PROGRESS - Check updates regularly!')
    print ('Please report any issue at https://github.com/decalage2/ViperMonkey/issues')
    print ('')

    DEFAULT_LOG_LEVEL = "info"
    LOG_LEVELS = {
        'debug':    logging.DEBUG,
        'info':     logging.INFO,
        'warning':  logging.WARNING,
        'error':    logging.ERROR,
        'critical': logging.CRITICAL
        }

    usage = 'usage: %prog [options] <filename> [filename2 ...]'
    parser = optparse.OptionParser(usage=usage)
    parser.add_option('-p', '--parse', dest='parse_file',
         help='VBA text file to be parsed')
    parser.add_option('-e', '--eval', dest='eval_expr',
        help='VBA expression to be evaluated')
    parser.add_option('-l', '--loglevel', dest="loglevel", action="store", default=DEFAULT_LOG_LEVEL,
                            help="logging level debug/info/warning/error/critical (default=%default)")

    (options, args) = parser.parse_args()
    colorlog.basicConfig(level=LOG_LEVELS[options.loglevel], format='%(log_color)s%(levelname)-8s %(message)s')

    if options.parse_file:
        parse(options.parse_file)

    if options.eval_expr:
        eval_expression(options.eval_expr)

    while True:
        try:
            print("VBA> ", end='')
            cmd = raw_input()

            if cmd.startswith('exit'):
                break

            if cmd.startswith('parse'):
                parse()

            if cmd.startswith('trace'):
                args = cmd.split()
                print('Tracing %s' % args[1])
                vm.trace(entrypoint=args[1])
                print('Recorded Actions:')
                print(vm.dump_actions())

            if cmd.startswith('eval'):
                expr = cmd[5:]
                eval_expression(expr)
        except Exception:
            simulation_vba.log.exception('ERROR')

if __name__ == '__main__':
    main()

