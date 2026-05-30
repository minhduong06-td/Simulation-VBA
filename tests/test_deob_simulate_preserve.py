import os
import subprocess
import sys


ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir))
CORE = os.path.join(ROOT, "simulation_vba", "core")
if CORE not in sys.path:
    sys.path.insert(0, CORE)

import deobfuscation


HTA_SAMPLE = r'''<html>
<head>
<script language="VBScript">
Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
Const PREFIX = "cmd"
Dim moduleValue

Sub Main()
    On Error Resume Next
    Dim cmd
    cmd = Chr(99) & Chr(97) & Chr(108) & Chr(99)
    If Len(cmd) > 0 Then
        For i = 1 To 1
            Do
                Select Case i
                    Case 1
                        Shell cmd
                        GoTo Done
                End Select
                Exit Do
            Loop
        Next
    End If
Done:
    Set sh = CreateObject("WScript.Shell")
    sh.Run cmd
    Helper
End Sub

Sub Helper()
    Shell "helper-called"
End Sub

Function NotCalled()
    Shell "never-run"
    NotCalled = "keep"
End Function
</script>
</head>
</html>
'''


def test_deob_simulate_preserves_hta_vba_and_stubs_dangerous_actions(monkeypatch, tmpdir):
    def fail_execute(*args, **kwargs):
        raise AssertionError("simulate/deob must not execute subprocesses")

    monkeypatch.setattr(subprocess, "Popen", fail_execute)
    monkeypatch.setattr(subprocess, "call", fail_execute)
    monkeypatch.setattr(os, "system", fail_execute)
    monkeypatch.chdir(str(tmpdir))
    before = set(os.listdir(str(tmpdir)))

    deobfuscated, actions = deobfuscation.simulate_deobfuscation(HTA_SAMPLE)

    assert before == set(os.listdir(str(tmpdir)))
    assert 'Sub Main()' in deobfuscated
    assert 'Sub Helper()' in deobfuscated
    assert 'Function NotCalled()' in deobfuscated
    assert 'Declare PtrSafe Function GetTickCount' in deobfuscated
    assert 'Const PREFIX = "cmd"' in deobfuscated
    assert 'Dim moduleValue' in deobfuscated
    assert 'On Error Resume Next' in deobfuscated
    assert 'If Len(cmd) > 0 Then' in deobfuscated
    assert 'For i = 1 To 1' in deobfuscated
    assert 'Do' in deobfuscated
    assert 'Select Case i' in deobfuscated
    assert 'GoTo Done' in deobfuscated
    assert 'Done:' in deobfuscated
    assert 'cmd = "calc"' in deobfuscated

    action_text = "\n".join("%s %s %s" % action for action in actions)
    assert "Shell function stubbed" in action_text
    assert "WScript.Shell.Run() stubbed" in action_text
    assert "CreateObject stubbed" in action_text
    assert "calc" in action_text
    assert "helper-called" in action_text
    assert "never-run" not in action_text
