import logging
import re
from logger import log
import statements

def _transform_dummy_loop1(loop):

    loop_pat = r"Do\s+While\s+(\w+)\s*<\s*(\d+)\r?\n.{0,500}?Loop"
    loop_str = loop.original_str
    if (re.search(loop_pat, loop_str, re.DOTALL) is None):
        return loop

    info = re.findall(loop_pat, loop_str, re.DOTALL)
    loop_var = info[0][0].strip()
    loop_ub = int(info[0][1].strip())
    
    if_pat = r"If\s+\(?\s*(\w+)\s*=\s*(\d+)\s*\)\s+Then\s*\r?\n?(.{10,200}?)End\s+If"
    if_info = re.findall(if_pat, loop_str, re.DOTALL)
    if (len(if_info) == 0):
        return loop

    run_statements = []
    for curr_if in if_info:

        test_var = curr_if[0].strip()

        test_val = int(curr_if[1].strip())

        if (test_var != loop_var):
            continue

        if (test_val >= loop_ub):
            continue

        run_statement = curr_if[2].strip()
        if (run_statement.endswith("Else")):
            run_statement = run_statement[:-len("Else")]
        run_statements.append(run_statement)

    if (len(run_statements) == 0):
        return loop

    loop_repl = ""
    for run_statement in run_statements:
        loop_repl += run_statement + "\n"

    import statements
    try:
        obj = statements.statement_block.parseString(loop_repl, parseAll=True)[0]
    except:
        return loop
    return obj

def _transform_wait_loop(loop):

    loop_pat = r"[Ww]hile\s+\w+\s*<>\s*\"?\w+\"?\r?\n.{0,500}?[Ww]end"
    loop_str = loop.original_str
    if (re.search(loop_pat, loop_str, re.DOTALL) is None):
        return loop

    if ((len(loop.body) > 1) or (len(loop.body) == 0) or
        (not isinstance(loop.body[0], statements.Call_Statement))):
        return loop

    log.warning("Transformed possible infinite wait loop...")
    return loop.body[0]
    
def transform_loop(loop):

    import statements
    if (not isinstance(loop, statements.While_Statement)):
        return loop
    
    r = _transform_dummy_loop1(loop)
    r = _transform_wait_loop(r)
    
    return r
