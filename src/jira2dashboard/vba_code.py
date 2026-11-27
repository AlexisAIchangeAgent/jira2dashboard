"""The code of the workbook in one text file for Rovo (AL, 2026-11-27: "Rovo cannot read the macros of an xlsm, it reads an xlsx: compile all the VBA
code in one text file so that it can debug easily"). It replaces the button Debug prompt for Rovo (01.md, 02.md... of v0.20.0).

At build, flow/sheets.py writes the sheet VbaCode (hidden): one row per line of each DELIVERED module (build_workbook.MODULES, in that order), Module |
Line | Text. The workbook never reads its own VBA project (that needs "Trust access to the VBA project object model", and an endpoint protection reads a
macro that reads macros as malware). The button "VBA code for Rovo" (modRovoDebug.WriteVbaCode) writes `file_text` of that sheet and of the sheet
FlowFormulas (the formula of every column of the tables) into rovo-debug\\vba-code.txt next to the workbook, UTF-8 without BOM, lines ended by LF.
No value of Jira, no setting: the code and the formulas only. The Excel scenario compares the file with `file_text` character by character.
"""
from pathlib import Path

# The delivered modules, in the order of the build: no test module, no test button; modCoverage stays because the other modules call Cov (a no-op
# unless switched on). build_workbook.MODULES is this list.
MODULES = ["modCoverage", "modLog", "modJson", "modText", "modHttp", "modProxy", "modJira", "modFlow", "modParams", "modExtract",
           "modJiraConfig", "modPropose", "modHealth", "modFlowBuild", "modRovoDebug", "modActions"]
SHEET = "VbaCode"
HEADER = ("Module", "Line", "Text")
FILE_NAME = "vba-code.txt"
TITLE = "# jira2dashboard - the code of the workbook, for Rovo"
INTRO = ("This file holds the VBA modules of the workbook (the macros of the buttons) and the formula of every column of its tables (sheet "
         "FlowFormulas_from_excel). It holds no value of Jira and no setting. Read the code to explain a figure; quote the module and the line.")
MODULE_HEAD = "===== VBA module {name} ({lines} lines) ====="
FORMULAS_HEAD = "===== Formulas of the tables (sheet FlowFormulas_from_excel): Table | Column | Formula ====="


def module_lines(text: str) -> list[str]:
    """The lines of a .bas file as the VBA editor keeps them: without the line "Attribute VB_Name" (the name is the module), the last newline
    not counted."""
    lines = text.replace("\r\n", "\n").split("\n")
    if lines and lines[-1] == "":
        lines.pop()
    return [line for line in lines if not line.startswith("Attribute VB_Name")]


def rows(modules: list[tuple[str, str]]) -> list[tuple]:
    """The rows of the sheet VbaCode: (module, line number from 1, text) for each line of each (name, text of the .bas) in the order given."""
    return [(name, n, line) for name, text in modules for n, line in enumerate(module_lines(text), start=1)]


def delivered(vba_dir: Path, names: list[str]) -> list[tuple[str, str]]:
    """(name, text) of the delivered modules, read from vba/<name>.bas."""
    return [(name, (vba_dir / f"{name}.bas").read_text(encoding="ascii")) for name in names]


def file_text(code_rows: list[tuple], formulas: list[tuple], version: str) -> str:
    """The text the button writes: the title, the version, what the file holds, the list of the modules with their line counts, then each module and
    the formulas (Table | Column | Formula, a values column has an empty Formula). Twin of modRovoDebug.WriteVbaCode."""
    counts: dict[str, int] = {}
    for name, _n, _line in code_rows:
        counts[name] = counts.get(name, 0) + 1
    out = [TITLE, f"Version: {version}", INTRO, "Modules: " + ", ".join(f"{name} ({n} lines)" for name, n in counts.items())]
    current = None
    for name, _n, line in code_rows:
        if name != current:
            out += ["", MODULE_HEAD.format(name=name, lines=counts[name])]
            current = name
        out.append("" if line is None else str(line))
    out += ["", FORMULAS_HEAD]
    out += [" | ".join("" if cell is None else str(cell) for cell in row) for row in formulas]
    return "\n".join(out) + "\n"
