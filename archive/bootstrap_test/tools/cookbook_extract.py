"""Recreate every file of the kit from the Markdown files, and check each MD5.

Usage, from the project folder (files are written relative to the current folder):
    py -3.14 tools/cookbook_extract.py docs/cookbook_python.md docs/vba-modules-tested.md
    add --force to overwrite files that already exist with another content.
Never retype the files by hand or by an AI: one wrong character breaks them silently.
"""
import hashlib
import re
import sys
from pathlib import Path

FENCE = "`" * 3    # never write three backticks literally here: it would end the Markdown code block
TICK = "`"
SECTION = re.compile(
    r"^### " + TICK + r"([^" + TICK + r"\n]+)" + TICK + r"  \(MD5 " + TICK + r"([0-9a-f]{32})" + TICK + r"\)\n\n"
    + FENCE + r"[a-z]*\n(.*?)\n" + FENCE + r"$", re.S | re.M)


def extract(md_path: Path, force: bool) -> int:
    """Write each file found in one Markdown file; return the number of problems."""
    problems = 0
    for rel, expected, code in SECTION.findall(md_path.read_text(encoding="utf-8")):
        data = (code + "\n").encode("utf-8")
        status = "OK"
        if hashlib.md5(data).hexdigest() != expected:
            status = "MD5 MISMATCH"
        elif rel.endswith(".bas") and not data.isascii():
            status = "NOT ASCII"
        else:
            target = Path(rel)
            if target.exists() and target.read_bytes() != data and not force:
                status = "EXISTS (use --force)"
            else:
                target.parent.mkdir(parents=True, exist_ok=True)
                target.write_bytes(data)
        problems += status != "OK"
        print(f"{status:22} {rel}")
    return problems


def main(args: list[str]) -> int:
    force = "--force" in args
    files = [Path(a) for a in args if not a.startswith("--")]
    if not files:
        print(__doc__)
        return 2
    problems = sum(extract(f, force) for f in files)
    print("all files OK" if problems == 0 else f"{problems} problem(s)")
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
