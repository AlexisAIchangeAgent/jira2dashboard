"""Patch 4: language-independent Repr for booleans in the test sheet."""
from pathlib import Path

p = Path(__file__).parent / "modTests.bas"
s = p.read_text(encoding="ascii").replace("\r\n", "\n")
old = '    If VarType(v) = vbDate Then Repr = Format$(v, "yyyy-mm-dd hh:nn:ss"): Exit Function\n'
new = old + '    If VarType(v) = vbBoolean Then Repr = IIf(v, "True", "False"): Exit Function   \' Str$(True) = "Vrai" in French\n'
assert s.count(old) == 1
p.write_text(s.replace(old, new), encoding="ascii", newline="\n")
print("patched")
