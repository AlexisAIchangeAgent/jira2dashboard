"""Patch 5: show test progress in the status bar (runs take 10-15 s; without progress users click again)."""
from pathlib import Path

p = Path(__file__).parent / "modTests.bas"
s = p.read_text(encoding="ascii").replace("\r\n", "\n")
old = """    Dim nm As Variant, ok As Long, ko As Long, r As Variant
    Set mRows = New Collection
    Randomize 1
    For Each nm In names
        mCurrent = nm
"""
new = """    Dim nm As Variant, ok As Long, ko As Long, r As Variant, i As Long
    Set mRows = New Collection
    Randomize 1
    For Each nm In names
        i = i + 1
        Application.StatusBar = "Test " & i & "/" & (UBound(names) - LBound(names) + 1) & ": " & nm & "..."
        mCurrent = nm
"""
assert s.count(old) == 1
s = s.replace(old, new)
old2 = """    WriteResults
    RunTests = ok & " OK, " & ko & " KO"
"""
new2 = """    Application.StatusBar = False
    WriteResults
    RunTests = ok & " OK, " & ko & " KO"
"""
assert s.count(old2) == 1
s = s.replace(old2, new2)
p.write_text(s, encoding="ascii", newline="\n")
print("patched")
