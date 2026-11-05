"""Arguments of every entry point: the module docstring is the help, so --help explains how the script works.

Each command parses its arguments first, so --help (and a wrong argument) stops before any clean, Excel or network.
"""
import argparse
from pathlib import Path


def parser(prog: str, doc: str, epilog: str = "") -> argparse.ArgumentParser:
    """`prog`: how the user runs it (for example "scripts\\run.ps1 build"). `doc`: the module docstring."""
    return argparse.ArgumentParser(prog=prog, description=doc.strip(), epilog=epilog.strip() or None,
                                   formatter_class=argparse.RawDescriptionHelpFormatter)


def bare_xlsm(name: str) -> str:
    """The workbook name of --out: a bare file name, so that nothing outside the folder out is ever overwritten."""
    if name != Path(name).name or not name.lower().endswith(".xlsm"):
        raise argparse.ArgumentTypeError(f"a bare file name ending in .xlsm (it goes under the folder out), "
                                         f"not {name!r}")
    return name
