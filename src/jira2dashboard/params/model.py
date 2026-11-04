"""Load and validate the parameter tables of params/*.md. Pure (no Excel).

Vocabulary follows docs/safe-flow-metrics-and-data-checks.md. All problems are reported together.
"""
import datetime as dt
import re
from collections import Counter
from dataclasses import dataclass
from pathlib import Path

from jira2dashboard.params.mdtable import MarkdownTableError, parse_table
from jira2dashboard.paths import PARAMS_DIR

WORKFLOWS = ("Team", "Feature")
CATEGORIES = ("Active", "Wait", "Done", "Removed", "Backlog", "Unclassified")
ROLES = ("Required", "Optional", "Display")      # a role may continue: "Required for PI views"
YES_NO = ("yes", "no")


class ParamsError(ValueError):
    """One or more problems in params/*.md, one per line."""


@dataclass(frozen=True)
class SheetSpec:
    name: str                          # sheet name in the workbook
    filename: str                      # file in params/
    columns: tuple[str, ...]           # exact header, same as the sheet header
    int_columns: frozenset[str] = frozenset()


@dataclass(frozen=True)
class ParamSheet:
    name: str
    header: tuple[str, ...]
    rows: tuple[tuple, ...]            # cells: str, int or None (empty)


CONFIG_PARAMETERS = ("BaseUrl", "Email", "ApiToken", "Filter", "WorkTypeField", "ProxyServer", "ProxyBypass",
                     "ProxyWindowsAuth", "FlowLoadMode", "AsOf", "TimeoutSeconds", "MaxRetries")
PROJECT_ROLES = ("Program", "Team")
_PROJECT_KEY = re.compile(r"[A-Z][A-Z0-9_]+")

SPECS = (
    SheetSpec("Config", "config.md", ("Parameter", "Value", "Description")),
    SheetSpec("Projects", "projects.md", ("Project", "Role", "Team")),
    SheetSpec("Periods", "periods.md", ("Period", "Start", "End")),
    SheetSpec("Workflows", "workflows.md",
              ("Workflow", "Order", "Status", "Category", "IsStart", "IsEnd", "IsFree"),
              frozenset({"Order", "IsStart", "IsEnd", "IsFree"})),
    SheetSpec("Transitions", "transitions.md", ("Workflow", "From", "To", "Type")),
    SheetSpec("FieldMapping", "field_mapping.md",
              ("Extract", "Logical", "Type", "CanBeEmpty", "UsedBy", "Role", "IfMissing", "TypeConfirmed")),
)


def _typed_cell(spec: SheetSpec, column: str, raw: str, where: str, problems: list[str]):
    if column in spec.int_columns:
        if raw in ("", "-"):
            return None
        if raw.isdigit():
            return int(raw)
        problems.append(f"{where}: {column} must be a whole number or '-', got {raw!r}")
        return None
    return raw or None


def _read(spec: SheetSpec, params_dir: Path, problems: list[str]) -> ParamSheet | None:
    path = params_dir / spec.filename
    try:
        table = parse_table(path.read_text(encoding="utf-8"))
    except FileNotFoundError:
        problems.append(f"{spec.filename}: file not found in {params_dir}")
        return None
    except MarkdownTableError as err:
        problems.append(f"{spec.filename}: {err}")
        return None
    if table.header != spec.columns:
        problems.append(f"{spec.filename}: header is {list(table.header)}, expected {list(spec.columns)}")
        return None
    rows = []
    for number, raw_row in enumerate(table.rows, start=1):
        where = f"{spec.name} row {number}"
        rows.append(tuple(_typed_cell(spec, col, raw, where, problems)
                          for col, raw in zip(spec.columns, raw_row)))
    return ParamSheet(spec.name, spec.columns, tuple(rows))


def _records(sheet: ParamSheet):
    """(row number, {column: value}) for each row."""
    return [(n, dict(zip(sheet.header, row))) for n, row in enumerate(sheet.rows, start=1)]


def _check_workflows(sheet: ParamSheet, problems: list[str]) -> None:
    seen: Counter = Counter()
    for number, rec in _records(sheet):
        where = f"Workflows row {number}"
        if rec["Workflow"] not in WORKFLOWS:
            problems.append(f"{where}: Workflow {rec['Workflow']!r} is not one of {', '.join(WORKFLOWS)}")
        if not rec["Status"]:
            problems.append(f"{where}: Status is empty")
        if rec["Category"] not in CATEGORIES:
            problems.append(f"{where}: Category {rec['Category']!r} is not one of {', '.join(CATEGORIES)}")
        for flag in ("IsStart", "IsEnd", "IsFree"):
            if rec[flag] not in (0, 1):
                problems.append(f"{where}: {flag} must be 0 or 1, got {rec[flag]!r}")
        seen[(rec["Workflow"], rec["Status"])] += 1
    for (workflow, status), count in seen.items():
        if count > 1:
            problems.append(f"Workflows: workflow {workflow!r} lists status {status!r} twice")
    for workflow in WORKFLOWS:
        recs = [r for _n, r in _records(sheet) if r["Workflow"] == workflow]
        if not recs:
            problems.append(f"Workflows: workflow {workflow!r} has no status")
            continue
        for flag in ("IsStart", "IsEnd"):
            count = sum(1 for r in recs if r[flag] == 1)
            if count != 1:
                problems.append(f"Workflows: workflow {workflow!r} needs exactly one {flag} status, found {count}")


def _statuses(workflows: ParamSheet) -> dict[str, set[str]]:
    result: dict[str, set[str]] = {wf: set() for wf in WORKFLOWS}
    for _n, rec in _records(workflows):
        if rec["Workflow"] in result and rec["Status"]:
            result[rec["Workflow"]].add(rec["Status"])
    return result


def _check_transitions(sheet: ParamSheet, workflows: ParamSheet, problems: list[str]) -> None:
    statuses = _statuses(workflows)
    seen: Counter = Counter()
    for number, rec in _records(sheet):
        where = f"Transitions row {number}"
        workflow = rec["Workflow"]
        if workflow not in WORKFLOWS:
            problems.append(f"{where}: Workflow {workflow!r} is not one of {', '.join(WORKFLOWS)}")
            continue
        for column in ("From", "To"):
            if rec[column] not in statuses[workflow]:
                problems.append(f"{where}: {column} {rec[column]!r} is not a status of workflow {workflow!r}")
        if not rec["Type"]:
            problems.append(f"{where}: Type is empty")
        seen[(workflow, rec["From"], rec["To"])] += 1
    for (workflow, source, target), count in seen.items():
        if count > 1:
            problems.append(f"Transitions: {workflow!r} {source!r} -> {target!r} listed twice")


def _check_field_mapping(sheet: ParamSheet, problems: list[str]) -> None:
    seen: Counter = Counter()
    for number, rec in _records(sheet):
        where = f"FieldMapping row {number}"
        for column in ("Extract", "Logical"):
            if not rec[column]:
                problems.append(f"{where}: {column} is empty")
        role = rec["Role"] or ""
        if role.split(" ")[0] not in ROLES:
            problems.append(f"{where}: Role {role!r} must start with one of {', '.join(ROLES)}")
        if rec["CanBeEmpty"] not in YES_NO:
            problems.append(f"{where}: CanBeEmpty {rec['CanBeEmpty']!r} must be yes or no")
        seen[rec["Extract"]] += 1
    for extract, count in seen.items():
        if count > 1:
            problems.append(f"FieldMapping: column {extract!r} listed twice")


def _check_config(sheet: ParamSheet, problems: list[str]) -> None:
    names = [row[0] for row in sheet.rows]
    for number, name in enumerate(names, start=1):
        if name not in CONFIG_PARAMETERS:
            problems.append(f"Config row {number}: unknown parameter {name!r}")
    for missing in (p for p in CONFIG_PARAMETERS if p not in names):
        problems.append(f"Config: parameter {missing!r} is missing")
    for name, count in Counter(names).items():
        if count > 1:
            problems.append(f"Config: parameter {name!r} listed twice")
    values = {row[0]: row[1] for row in sheet.rows}
    if values.get("ApiToken"):
        problems.append("Config: ApiToken must stay empty in params/config.md (type the token in the workbook)")
    if values.get("ProxyWindowsAuth") not in ("Yes", "No"):
        problems.append(f"Config: ProxyWindowsAuth must be Yes or No, got {values.get('ProxyWindowsAuth')!r}")
    for name, minimum in (("TimeoutSeconds", 1), ("MaxRetries", 0)):
        text = values.get(name)
        if text is None or not text.isdigit() or int(text) < minimum:
            problems.append(f"Config: {name} must be a whole number >= {minimum}, got {text!r}")
    base = values.get("BaseUrl")
    if base and not base.lower().startswith(("http://", "https://")):
        problems.append(f"Config: BaseUrl must start with http:// or https://, got {base!r}")
    if values.get("FlowLoadMode") not in ("A", "B"):
        problems.append(f"Config: FlowLoadMode must be A or B, got {values.get('FlowLoadMode')!r}")
    as_of = values.get("AsOf")
    if as_of and _parse_moment(as_of) is None:
        problems.append(f"Config: AsOf must be a UTC date or date and time such as 2026-11-03 18:00, got {as_of!r}")


def _parse_moment(text: str):
    try:
        return dt.datetime.fromisoformat(text.strip())
    except ValueError:
        return None


def _check_periods(sheet: ParamSheet, problems: list[str]) -> None:
    seen: Counter = Counter()
    for number, rec in _records(sheet):
        where = f"Periods row {number}"
        start, end = _parse_moment(rec["Start"] or ""), _parse_moment(rec["End"] or "")
        if not rec["Period"]:
            problems.append(f"{where}: Period needs a name")
        if start is None or end is None:
            problems.append(f"{where}: Start and End must be dates such as 2026-10-05, got {rec['Start']!r} and {rec['End']!r}")
        elif start >= end:
            problems.append(f"{where}: Start must be before End")
        seen[rec["Period"]] += 1
    for name, count in seen.items():
        if count > 1:
            problems.append(f"Periods: period {name!r} listed twice")


def _check_projects(sheet: ParamSheet, problems: list[str]) -> None:
    seen: Counter = Counter()
    for number, rec in _records(sheet):
        where = f"Projects row {number}"
        key = rec["Project"] or ""
        if not _PROJECT_KEY.fullmatch(key):
            problems.append(f"{where}: Project {key!r} is not a Jira project key (capital letters, digits, _)")
        if rec["Role"] not in PROJECT_ROLES:
            problems.append(f"{where}: Role {rec['Role']!r} is not one of {', '.join(PROJECT_ROLES)}")
        if rec["Role"] == "Team" and not rec["Team"]:
            problems.append(f"{where}: a Team project needs a Team name")
        seen[key] += 1
    for key, count in seen.items():
        if count > 1:
            problems.append(f"Projects: project {key!r} listed twice")
    programs = [rec["Project"] for _n, rec in _records(sheet) if rec["Role"] == "Program"]
    if len(programs) != 1:
        problems.append(f"Projects: exactly one Program project is needed, found {len(programs)}")
    if not any(rec["Role"] == "Team" for _n, rec in _records(sheet)):
        problems.append("Projects: at least one Team project is needed")


def project_keys(sheets: list[ParamSheet] | None = None) -> tuple[str, ...]:
    """The project keys, the Program first (the order the fake Jira generator expects)."""
    rows = next(s for s in (sheets or load_all()) if s.name == "Projects").rows
    return tuple(r[0] for r in sorted(rows, key=lambda r: r[1] != "Program"))


def load_all(params_dir: Path = PARAMS_DIR) -> list[ParamSheet]:
    """Every parameter sheet, in workbook order. Raises ParamsError listing all the problems."""
    problems: list[str] = []
    sheets = {spec.name: _read(spec, params_dir, problems) for spec in SPECS}
    if sheets["Config"]:
        _check_config(sheets["Config"], problems)
    if sheets["Projects"]:
        _check_projects(sheets["Projects"], problems)
    if sheets["Periods"]:
        _check_periods(sheets["Periods"], problems)
    if sheets["Workflows"]:
        _check_workflows(sheets["Workflows"], problems)
    if sheets["Transitions"] and sheets["Workflows"]:
        _check_transitions(sheets["Transitions"], sheets["Workflows"], problems)
    if sheets["FieldMapping"]:
        _check_field_mapping(sheets["FieldMapping"], problems)
    if problems:
        raise ParamsError("\n".join(problems))
    return [sheets[spec.name] for spec in SPECS]
