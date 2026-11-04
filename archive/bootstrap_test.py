"""Bootstrap test on Windows: rebuild the kit from docs/ in an empty folder and compare with the tested files."""
import filecmp
import shutil
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).parent
DEST = HERE / "bootstrap_test"
shutil.rmtree(DEST, ignore_errors=True)
(DEST / "docs").mkdir(parents=True)
(DEST / "tools").mkdir()
for md in ("cookbook_python.md", "vba-modules-tested.md"):
    shutil.copyfile(HERE / "docs" / md, DEST / "docs" / md)

# Simulate the one copy-paste: the first block of section 4 of the cookbook.
text = (DEST / "docs" / "cookbook_python.md").read_text(encoding="utf-8")
fence = "`" * 3
block = text[text.index("### `tools/cookbook_extract.py`"):].split(fence + "python\n", 1)[1].split("\n" + fence, 1)[0]
(DEST / "tools" / "cookbook_extract.py").write_text(block + "\n", encoding="utf-8", newline="\n")

run = subprocess.run([sys.executable, "tools/cookbook_extract.py", "docs/cookbook_python.md", "docs/vba-modules-tested.md"],
                     cwd=DEST, capture_output=True, text=True)
print(run.stdout.strip().splitlines()[-1], "| exit", run.returncode)
names = ["xlsm_ref.py", "_check_child.py", "_hang_child.py", "_run_tests_child.py", "check_reference.py",
         "run_vba_tests.py"] + [f"vba/{p.name}" for p in sorted((HERE / "vba").glob("*.bas"))]
same = [n for n in names if filecmp.cmp(DEST / n, HERE / n, shallow=False)]
print(f"identical to the tested files: {len(same)}/{len(names)}")
for n in names:
    if n not in same:
        print("  DIFFERENT:", n)
