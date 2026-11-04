"""Patch 3: test folders next to the workbook (no %TEMP%), so everything stays under C:\\dev."""
from pathlib import Path

p = Path(__file__).parent / "modTests.bas"
s = p.read_text(encoding="ascii").replace("\r\n", "\n")
for old, new in [(r'Environ$("TEMP") & "\xlsm_vba_test"', r'ThisWorkbook.Path & "\tmp_files"'),
                 (r'Environ$("TEMP") & "\xlsm_vba_bg"', r'ThisWorkbook.Path & "\tmp_export"')]:
    assert s.count(old) == 1, old
    s = s.replace(old, new)
p.write_text(s, encoding="ascii", newline="\n")
print("patched")
