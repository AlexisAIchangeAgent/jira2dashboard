"""Mutation proof of a VBA test, for the skill vba-mutation: break a module on purpose, run the tests, restore.
Standard library only. A module is changed in place; its good text is kept in %TEMP%\\jira2dashboard\\mutation until `restore`.

    .venv\\Scripts\\python.exe .claude\\skills\\vba-mutation\\mutate.py apply vba\\modJira.bas --old "text" --new "other text"
    .venv\\Scripts\\python.exe .claude\\skills\\vba-mutation\\mutate.py restore vba\\modJira.bas
    .venv\\Scripts\\python.exe .claude\\skills\\vba-mutation\\mutate.py status

apply    the OLD text must occur exactly once (else nothing changes); refuses a module that is already mutated
restore  puts the good text back, and fails if the mutation is still there afterwards
status   lists the modules left mutated (exit 1 when there is one): run it before a commit
The file stays ASCII with LF endings, as the VBA modules must (tests/unit/test_run_vba_tests_pure.py).
"""
from __future__ import annotations

import argparse
import os
import shutil
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[3]
KEEP = Path(os.environ.get("J2D_MUTATION_DIR") or Path(tempfile.gettempdir()) / "jira2dashboard" / "mutation")


def _module(path: str) -> Path:
    module = Path(path)
    module = module if module.is_absolute() else ROOT / module
    if module.suffix.lower() != ".bas" or not module.is_file():
        raise SystemExit(f"{path}: not a .bas module of this repository")
    return module


def _good(module: Path) -> Path:
    return KEEP / (module.name + ".good")


def _marker(module: Path) -> Path:
    return KEEP / (module.name + ".mutation")


def apply(path: str, old: str, new: str) -> int:
    module = _module(path)
    if _good(module).exists():
        raise SystemExit(f"{module.name} is already mutated: restore it first")
    text = module.read_bytes().decode("ascii")
    count = text.count(old)
    if count != 1:
        raise SystemExit(f"the old text occurs {count} times in {module.name}, it must occur exactly once")
    if old == new:
        raise SystemExit("the new text is the old one: that is no mutation")
    mutated = text.replace(old, new)
    try:
        mutated.encode("ascii")                                # a VBA module is ASCII
    except UnicodeEncodeError:
        raise SystemExit("the new text must be ASCII: a VBA module is") from None
    KEEP.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(module, _good(module))
    _marker(module).write_text(new, encoding="ascii")
    module.write_bytes(mutated.encode("ascii"))                # bytes: LF stays LF
    print(f"mutated {module.name}; restore with: mutate.py restore {path}")
    return 0


def restore(path: str) -> int:
    module = _module(path)
    if not _good(module).exists():
        print(f"{module.name}: nothing to restore")
        return 0
    shutil.copyfile(_good(module), module)
    new = _marker(module).read_text(encoding="ascii")
    if new in module.read_bytes().decode("ascii") and new not in _good(module).read_bytes().decode("ascii"):
        raise SystemExit(f"{module.name}: the mutation is still there after the restore")
    _good(module).unlink()
    _marker(module).unlink(missing_ok=True)
    print(f"restored {module.name}")
    return 0


def status() -> int:
    left = sorted(p.name[: -len(".good")] for p in KEEP.glob("*.good")) if KEEP.is_dir() else []
    for name in left:
        print(f"MUTATED: {name}")
    print("no module left mutated" if not left else "restore the modules above before going on")
    return 1 if left else 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="mutate.py", description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = parser.add_subparsers(dest="command", required=True)
    one = sub.add_parser("apply", help="replace OLD (exactly once) by NEW in a module")
    one.add_argument("module")
    one.add_argument("--old", required=True)
    one.add_argument("--new", required=True)
    two = sub.add_parser("restore", help="put the good text of a module back")
    two.add_argument("module")
    sub.add_parser("status", help="list the modules left mutated")
    args = parser.parse_args(argv)
    if args.command == "apply":
        return apply(args.module, args.old, args.new)
    return restore(args.module) if args.command == "restore" else status()


if __name__ == "__main__":
    sys.exit(main())
