#!/usr/bin/env python


import operator

__version__ = '0.03'


import logging
import sys
from collections import Iterable

from vba_object import *

from logger import log
from vba_object import coerce_to_num
from vba_object import to_python

def debug_repr(op, args):
    r = "("
    first = True
    for arg in args:
        if (not first):
            r += " " + op + " "
        first = False
        r += str(arg)
    r += ")"
    return r


class Sum(VBA_Object):
    """
    VBA Sum using the operator +
    """

    def __init__(self, original_str, location, tokens):
        super(Sum, self).__init__(original_str, location, tokens)
        self.arg = tokens[0][::2]

    def eval(self, context, params=None):

        evaluated_args = eval_args(self.arg, context)
        if ((isinstance(evaluated_args, Iterable)) and ("**MATCH ANY**" in evaluated_args)):
            return "**MATCH ANY**"
        
        try:
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("Compute sum (1) " + str(self.arg))
            r = reduce(lambda x, y: x + y, coerce_args(evaluated_args, preferred_type="int"))
            return r
        except (TypeError, ValueError):
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug('Impossible to sum arguments of different types. Try converting strings to common type.')
            try:
                r = reduce(lambda x, y: int(x) + int(y), evaluated_args)
                return r
            except (TypeError, ValueError):
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("Compute sum (2) " + str(self.arg))
                r = reduce(lambda x, y: str(x) + str(y), coerce_args_to_str(evaluated_args))
                return r
        except RuntimeError as e:
            log.error("overflow trying eval sum: %r" % self.arg)
            raise e

    def to_python(self, context, params=None, indent=0):
        r = ""
        first = True
        for arg in self.arg:
            if (not first):
                r += " + "
            first = False
            r += to_python(arg, context, params=params)
        return "(" + r + ")"
        
    def __repr__(self):
        return debug_repr("+", self.arg)
        return ' + '.join(map(repr, self.arg))


class Eqv(VBA_Object):
    """
    VBA Eqv operator.
    """

    def __init__(self, original_str, location, tokens):
        super(Eqv, self).__init__(original_str, location, tokens)
        self.arg = tokens[0][::2]

    def eval(self, context, params=None):

        evaluated_args = eval_args(self.arg, context)
        if ((isinstance(evaluated_args, Iterable)) and ("**MATCH ANY**" in evaluated_args)):
            return "**MATCH ANY**"

        try:
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("Compute eqv " + str(self.arg))
            return reduce(lambda a, b: (a & b) | ~(a | b), coerce_args(evaluated_args, preferred_type="int"))
        except (TypeError, ValueError):
            log.error('Impossible to Eqv arguments of different types.')
            return 0
        except RuntimeError as e:
            log.error("overflow trying eval Eqv: %r" % self.arg)
            raise e

    def __repr__(self):
        return ' Eqv '.join(map(repr, self.arg))
    

class Xor(VBA_Object):
    """
    VBA Xor operator.
    """

    def __init__(self, original_str, location, tokens):
        super(Xor, self).__init__(original_str, location, tokens)
        self.arg = tokens[0][::2]

    def eval(self, context, params=None):

        evaluated_args = eval_args(self.arg, context)
        if ((isinstance(evaluated_args, Iterable)) and ("**MATCH ANY**" in evaluated_args)):
            return "**MATCH ANY**"

        try:
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("Compute xor " + str(self.arg))
            return reduce(lambda x, y: x ^ y, coerce_args(evaluated_args, preferred_type="int"))
        except (TypeError, ValueError):
            try:
                return reduce(lambda x, y: int(x) ^ int(y), evaluated_args)
            except Exception as e:
                log.error('Impossible to xor arguments of different types. Arg list = ' + str(self.arg) + ". " + str(e))
                return 0
        except RuntimeError as e:
            log.error("overflow trying eval xor: %r" % self.arg)
            raise e

    def __repr__(self):
        return ' ^ '.join(map(repr, self.arg))

    def to_python(self, context, params=None, indent=0):
        r = ""
        first = True
        for arg in self.arg:
            if (not first):
                r += " ^ "
            first = False
            r += "coerce_to_int(" + to_python(arg, context, params=params) + ")"
        return "(" + r + ")"
    

class And(VBA_Object):
    """
    VBA And operator.
    """

    def __init__(self, original_str, location, tokens):
        super(And, self).__init__(original_str, location, tokens)
        self.arg = tokens[0][::2]

    def to_python(self, context, params=None, indent=0):
        r = ""
        first = True
        for arg in self.arg:
            if (not first):
                r += " & "
            first = False
            r += "coerce_to_int(" + to_python(arg, context, params=params) + ")"
        return "(" + r + ")"
        
    def eval(self, context, params=None):

        evaluated_args = eval_args(self.arg, context)
        if ((isinstance(evaluated_args, Iterable)) and ("**MATCH ANY**" in evaluated_args)):
            return "**MATCH ANY**"

        try:
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("Compute and " + str(self.arg))
            return reduce(lambda x, y: x & y, coerce_args(evaluated_args, preferred_type="int"))
        except (TypeError, ValueError):
            try:
                return reduce(lambda x, y: int(x) & int(y), evaluated_args)
            except:
                log.error('Impossible to and arguments of different types.')
                return 0
        except RuntimeError as e:
            log.error("overflow trying eval and: %r" % self.arg)
            raise e

    def __repr__(self):
        return ' & '.join(map(repr, self.arg))


class Or(VBA_Object):
    """
    VBA Or operator.
    """

    def __init__(self, original_str, location, tokens):
        super(Or, self).__init__(original_str, location, tokens)
        self.arg = tokens[0][::2]

    def eval(self, context, params=None):

        evaluated_args = eval_args(self.arg, context)
        if ((isinstance(evaluated_args, Iterable)) and ("**MATCH ANY**" in evaluated_args)):
            return "**MATCH ANY**"

        try:
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("Compute or " + str(self.arg))
            return reduce(lambda x, y: x | y, coerce_args(evaluated_args, preferred_type="int"))
        except (TypeError, ValueError):
            try:
                return reduce(lambda x, y: int(x) | int(y), evaluated_args)
            except:
                log.error('Impossible to or arguments of different types.')
                return 0
        except RuntimeError as e:
            log.error("overflow trying eval or: %r" % self.arg)
            raise e

    def to_python(self, context, params=None, indent=0):
        r = ""
        first = True
        for arg in self.arg:
            if (not first):
                r += " | "
            first = False
            r += "coerce_to_int(" + to_python(arg, context, params=params) + ")"
        return "(" + r + ")"
        
    def __repr__(self):
        return ' | '.join(map(repr, self.arg))


class Not(VBA_Object):
    """
    VBA binary Not operator.
    """

    def __init__(self, original_str, location, tokens):
        super(Not, self).__init__(original_str, location, tokens)
        self.arg = tokens[0][1]
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed %r as binary Not' % self)

    def eval(self, context, params=None):

        evaluated_args = eval_args(self.arg, context)
        if ((isinstance(evaluated_args, Iterable)) and ("**MATCH ANY**" in evaluated_args)):
            return "**MATCH ANY**"

        try:
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("Compute not " + str(self.arg))
            val = self.arg
            if (isinstance(val, VBA_Object)):
                val = val.eval(context)
            return (~ int(val))
        except Exception as e:
            log.error("Cannot compute Not " + str(self.arg) + ". " + str(e))
            return "NULL"

    def to_python(self, context, params=None, indent=0):
        r = "~ (" + "coerce_to_int(" + to_python(self.arg, context) + "))"
        return r
        
    def __repr__(self):
        return "Not " + str(self.arg)


class Neg(VBA_Object):
    """
    VBA binary Not operator.
    """

    def __init__(self, original_str, location, tokens):
        super(Neg, self).__init__(original_str, location, tokens)
        self.arg = tokens[0][1]
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('parsed %r as unary negation' % self)

    def to_python(self, context, params=None, indent=0):
        r = "- (" + "coerce_to_num(" + to_python(self.arg, context) + "))"
        return r
            
    def eval(self, context, params=None):

        evaluated_args = eval_args(self.arg, context)
        if ((isinstance(evaluated_args, Iterable)) and ("**MATCH ANY**" in evaluated_args)):
            return "**MATCH ANY**"

        try:
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("Compute negate " + str(self.arg))
            val = self.arg
            if (isinstance(val, VBA_Object)):
                val = val.eval(context)
            return (- int(val))
        except Exception as e:
            log.error("Cannot compute negation of " + str(self.arg) + ". " + str(e))
            return "NULL"

    def __repr__(self):
        return "-" + str(self.arg)
    

class Subtraction(VBA_Object):
    """
    VBA Subtraction using the binary operator -
    """

    def __init__(self, original_str, location, tokens):
        super(Subtraction, self).__init__(original_str, location, tokens)
        self.arg = tokens[0][::2]

    def eval(self, context, params=None):

        evaluated_args = eval_args(self.arg, context)
        if ((isinstance(evaluated_args, Iterable)) and ("**MATCH ANY**" in evaluated_args)):
            return "**MATCH ANY**"

        try:
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("Compute subract " + str(self.arg))
            return reduce(lambda x, y: x - y, coerce_args(evaluated_args, preferred_type="int"))
        except (TypeError, ValueError):
            try:
                return reduce(lambda x, y: coerce_to_int(x) - coerce_to_int(y), evaluated_args)
            except Exception as e:

                l1 = []
                orig = evaluated_args
                for v in orig:
                    if (isinstance(v, int)):
                        l1.append(v)
                        continue
                    if (isinstance(v, str) and (len(v) == 1)):
                        l1.append(ord(v))
                        continue

                if (len(orig) != len(l1)):                
                    log.error('Impossible to subtract arguments of different types. ' + str(e))
                    return 0

                return reduce(lambda x, y: int(x) - int(y), l1)

    def __repr__(self):
        return debug_repr("-", self.arg)
        return ' - '.join(map(repr, self.arg))


class Multiplication(VBA_Object):
    """
    VBA Multiplication using the binary operator *
    """

    def __init__(self, original_str, location, tokens):
        super(Multiplication, self).__init__(original_str, location, tokens)
        self.arg = tokens[0][::2]

    def eval(self, context, params=None):

        evaluated_args = eval_args(self.arg, context)
        if ((isinstance(evaluated_args, Iterable)) and ("**MATCH ANY**" in evaluated_args)):
            return "**MATCH ANY**"

        try:
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("Compute mult " + str(self.arg))
            return reduce(lambda x, y: x * y, coerce_args(evaluated_args, preferred_type="int"))
        except (TypeError, ValueError):
            try:
                return reduce(lambda x, y: int(x) * int(y), evaluated_args)
            except Exception as e:
                log.error('Impossible to multiply arguments of different types. ' + str(e))
                return 0

    def __repr__(self):
        return debug_repr("*", self.arg)
        return ' * '.join(map(repr, self.arg))


class Power(VBA_Object):
    """
    VBA exponentiation using the binary operator ^
    """

    def __init__(self, original_str, location, tokens):
        super(Power, self).__init__(original_str, location, tokens)
        self.arg = tokens[0][::2]

    def eval(self, context, params=None):

        evaluated_args = eval_args(self.arg, context)
        if ((isinstance(evaluated_args, Iterable)) and ("**MATCH ANY**" in evaluated_args)):
            return "**MATCH ANY**"

        try:
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("Compute pow " + str(self.arg))
            return reduce(lambda x, y: pow(x, y), coerce_args(evaluated_args, preferred_type="int"))
        except (TypeError, ValueError):
            try:
                return reduce(lambda x, y: pow(int(x), int(y)), evaluated_args)
            except Exception as e:
                log.error('Impossible to do exponentiation with arguments of different types. ' + str(e))
                return 0

    def __repr__(self):
        return debug_repr("^", self.arg)
        return ' ^ '.join(map(repr, self.arg))

    def to_python(self, context, params=None, indent=0):
        r = reduce(lambda x, y: "pow(coerce_to_num(" + to_python(x, context) + "), coerce_to_num(" + to_python(y, context) + "))", self.arg)
        return r
    

class Division(VBA_Object):
    """
    VBA Division using the binary operator /
    """

    def __init__(self, original_str, location, tokens):
        super(Division, self).__init__(original_str, location, tokens)
        self.arg = tokens[0][::2]

    def eval(self, context, params=None):

        evaluated_args = eval_args(self.arg, context)
        if ((isinstance(evaluated_args, Iterable)) and ("**MATCH ANY**" in evaluated_args)):
            return "**MATCH ANY**"

        try:
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("Compute div " + str(self.arg))
            return reduce(lambda x, y: x / y, coerce_args(evaluated_args, preferred_type="int"))
        except (TypeError, ValueError):
            try:
                return reduce(lambda x, y: int(x) / int(y), evaluated_args)
            except Exception as e:
                if (str(e).strip() != "division by zero"):
                    log.error('Impossible to divide arguments of different types. ' + str(e))
                return 0
        except ZeroDivisionError:
            context.set_error("Division by 0 error. Returning 'NULL'.")
            return 'NULL'

    def __repr__(self):
        return debug_repr("/", self.arg)
        return ' / '.join(map(repr, self.arg))


class MultiOp(VBA_Object):
    """
    Defines multiple operators that work within the same level of order or operations.
    """
    operator_map = {}

    def __init__(self, original_str, location, tokens):
        super(MultiOp, self).__init__(original_str, location, tokens)
        self.arg = tokens[0][::2]
        self.operators = tokens[0][1::2]

    def to_python(self, context, params=None, indent=0):

        set_flag = False
        if (not context.in_bitwise_expression):
            context.in_bitwise_expression = True
            set_flag = True
            
        if (self.operators[0] == "+"):
            ret = [to_python(self.arg[0], context, params=params)]
        else:
            ret = ["coerce_to_num(" + to_python(self.arg[0], context, params=params)  + ")"]
        for operator, arg in zip(self.operators, self.arg[1:]):
            if (operator == "+"):
                ret.append(' {} {!s}'.format("|plus|", to_python(arg, context, params=params)))
            else:
                ret.append(' {} {!s}'.format(operator, "coerce_to_num(" + to_python(arg, context, params=params) + ")"))

        if set_flag:
            context.in_bitwise_expression = False

        return '({})'.format(''.join(ret))
        
    def eval(self, context, params=None):

        set_flag = False
        if (not context.in_bitwise_expression):
            context.in_bitwise_expression = True
            set_flag = True
        
        evaluated_args = eval_args(self.arg, context)
        if ((isinstance(evaluated_args, Iterable)) and ("**MATCH ANY**" in evaluated_args)):
            if set_flag:
                context.in_bitwise_expression = False
            return "**MATCH ANY**"

        try:
            args = coerce_args(evaluated_args)
            ret = args[0]
            for operator, arg in zip(self.operators, args[1:]):
                try:
                    ret = self.operator_map[operator](ret, arg)
                except OverflowError:
                    log.error("overflow trying eval: %r" % str(self))
            if set_flag:
                context.in_bitwise_expression = False
            return ret
        except (TypeError, ValueError):
            try:
                args = map(coerce_to_num, evaluated_args)
                ret = args[0]
                for operator, arg in zip(self.operators, args[1:]):
                    ret = self.operator_map[operator](ret, arg)
                if set_flag:
                    context.in_bitwise_expression = False
                return ret
            except ZeroDivisionError:
                context.set_error("Division by 0 error. Returning 'NULL'.")
                if set_flag:
                    context.in_bitwise_expression = False
                return 'NULL'
            except Exception as e:
                log.error('Impossible to operate on arguments of different types. ' + str(e))
                if set_flag:
                    context.in_bitwise_expression = False
                return 0
        except ZeroDivisionError:
            context.set_error("Division by 0 error. Returning 'NULL'.")
            if set_flag:
                context.in_bitwise_expression = False
            return 'NULL'

    def __repr__(self):
        ret = [str(self.arg[0])]
        for operator, arg in zip(self.operators, self.arg[1:]):
            ret.append(' {} {!s}'.format(operator, arg))
        return '({})'.format(''.join(ret))


class MultiDiv(MultiOp):
    """
    VBA Multiplication/Division (used for performance)
    """
    operator_map = {'*': operator.mul, '/': operator.truediv}


class AddSub(MultiOp):
    """
    VBA Addition/Subtraction (used for performance)
    """
    operator_map = {'+': operator.add, '-': operator.sub}



class FloorDivision(VBA_Object):
    """
    VBA Floor Division using the binary operator \
    """

    def __init__(self, original_str, location, tokens):
        super(FloorDivision, self).__init__(original_str, location, tokens)
        self.arg = tokens[0][::2]

    def eval(self, context, params=None):

        evaluated_args = eval_args(self.arg, context)
        if ((isinstance(evaluated_args, Iterable)) and ("**MATCH ANY**" in evaluated_args)):
            return "**MATCH ANY**"

        try:
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("Compute floor div " + str(self.arg))
            return reduce(lambda x, y: x // y, coerce_args(evaluated_args, preferred_type="int"))
        except (TypeError, ValueError):
            try:
                return reduce(lambda x, y: int(x) // int(y), evaluated_args)
            except Exception as e:
                if (str(e).strip() != "division by zero"):
                    log.error('Impossible to divide arguments of different types. ' + str(e))
                return 0
        except ZeroDivisionError as e:
            context.set_error(str(e))
            
    def __repr__(self):
        return debug_repr("//", self.arg)
        return ' \\ '.join(map(repr, self.arg))

    def to_python(self, context, params=None, indent=0):
        r = ""
        first = True
        for arg in self.arg:
            if (not first):
                r += " // "
            first = False
            r += "coerce_to_num(" + to_python(arg, context, params=params) + ")"
        return "(" + r + ")"
    

class Concatenation(VBA_Object):
    """
    VBA String concatenation using the operator &
    """

    def __init__(self, original_str, location, tokens):
        super(Concatenation, self).__init__(original_str, location, tokens)
        self.arg = tokens[0][::2]
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('Concatenation: self.arg=%s' % repr(self.arg))

    def eval(self, context, params=None):

        _preview = [str(a)[:80] for a in self.arg]
        log.info("Concatenation.eval called: %d operands, preview=%r", len(self.arg), _preview)

        evaluated_args = eval_args(self.arg, context)
        if ((isinstance(evaluated_args, Iterable)) and ("**MATCH ANY**" in evaluated_args)):
            return "**MATCH ANY**"

        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('Concatenation before eval: %r' % params)

        coerce_args_to_str_safe = coerce_args_to_str
        eval_params = coerce_args_to_str_safe(evaluated_args)
        for idx, (raw, conv) in enumerate(zip(evaluated_args, eval_params)):
            preview = (str(conv)[:40].replace("\n", "\\n")
                       if isinstance(conv, str) else repr(conv))
            conv_len = len(conv) if isinstance(conv, str) else 0
            log.info(
                "Concat operand %d: type=%s raw_len=%d conv_len=%d len_mod_4=%d preview=%r",
                idx + 1, type(raw).__name__,
                len(str(raw)) if not isinstance(raw, str) else len(raw),
                conv_len,
                conv_len % 4,
                preview,
            )

        try:
            result = ''.join(eval_params)
            log.info(
                "Concatenated result: length=%d len_mod_4=%d preview=%r",
                len(result), len(result) % 4, result[:60].replace("\n", "\\n"),
            )
            return result
        except (TypeError, ValueError) as e:
            log.exception('Impossible to concatenate non-string arguments. ' + str(e))
            return ''

    def __repr__(self):
        return debug_repr("&", self.arg)
        return ' & '.join(map(repr, self.arg))

    def to_python(self, context, params=None, indent=0):
        r = ""
        first = True
        for arg in self.arg:
            if (not first):
                r += " + "
            first = False
            r += "coerce_to_str(" + to_python(arg, context, params=params) + ", zero_is_null=True)"
        return "(" + r + ")"


class Mod(VBA_Object):
    """
    VBA Modulo using the operator 'Mod'
    """

    def __init__(self, original_str, location, tokens):
        super(Mod, self).__init__(original_str, location, tokens)
        self.arg = tokens[0][::2]

    def eval(self, context, params=None):

        evaluated_args = eval_args(self.arg, context)
        if ((isinstance(evaluated_args, Iterable)) and ("**MATCH ANY**" in evaluated_args)):
            return "**MATCH ANY**"

        try:
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("Compute mod " + str(self.arg))
            result = reduce(lambda x, y: int(x) % int(y), coerce_args(evaluated_args, preferred_type="int"))
            log.info("Mod: args=%s => %d", str(evaluated_args), result)
            return result
        except (TypeError, ValueError) as e:
            log.error('Impossible to mod arguments of different types. ' + str(e))
            return ''
        except ZeroDivisionError:
            log.error('Mod division by zero error.')
            return ''

    def __repr__(self):
        return debug_repr("mod", self.arg)
        return ' mod '.join(map(repr, self.arg))

    def to_python(self, context, params=None, indent=0):
        r = ""
        first = True
        for arg in self.arg:
            if (not first):
                r += " % "
            first = False
            r += "coerce_to_num(" + to_python(arg, context, params=params) + ")"
        return "(" + r + ")"
