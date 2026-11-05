"""The parameters of a filled workbook, read as a file (no Excel): what the Python extraction needs.

The sheets Config, Projects, Periods, PIs, Workflows and Transitions are read by header name, so a workbook built
before a column was added still works (Config without Example or Mandatory, as the VBA accepts). Cells are brought to
the form of params/*.md: text trimmed like a pasted value, dates as ISO text (a date typed in Excel is a serial
number), the Workflows numbers as whole numbers. Then the rules shared with params/*.md are applied
(params.model.shape_problems), and those of the VBA for Config: defaults for an empty value, the mandatory values
(error 608 of the VBA), a readable AsOf (604), at least one period (703).
"""
import datetime as dt
from dataclasses import dataclass
from pathlib import Path

from jira2dashboard.flow.checks import FIELD_PARAMETERS, CheckSettings, Moves
from jira2dashboard.flow.model import Period, Rules
from jira2dashboard.params.model import (CONFIG_PARAMETERS, SPECS, ParamSheet, ParamsError, _parse_moment,
                                         shape_problems)
from jira2dashboard.standalone.xlsx_read import read_sheets

SHEETS = ("Config", "Projects", "Periods", "PIs", "Workflows", "Transitions")
OPTIONAL = {("Config", "Example"), ("Config", "Mandatory"), ("Config", "Description")}
DATE_COLUMNS = {("Periods", "Start"), ("Periods", "End"), ("PIs", "Start"), ("PIs", "End")}
EPOCH = dt.datetime(1899, 12, 30)
WHITE = " \t\r\n "                                # what the VBA TrimWhite removes around a pasted value
YES = ("yes", "oui", "true", "vrai", "1")             # modParams.IsYes
DEFAULTS = {"ProxyWindowsAuth": "Yes", "TimeoutSeconds": "60", "MaxRetries": "4", "FlowLoadMode": "A",
            "LogLevel": "Steps"}


def is_yes(value) -> bool:
    return str(value or "").strip(WHITE).lower() in YES


def serial_to_iso(serial: float) -> str:
    """An Excel date as ISO text: 2026-06-01, or 2026-11-03 18:00:00 when it has a time (rounded to the second)."""
    moment = EPOCH + dt.timedelta(seconds=round(serial * 86400))
    return f"{moment:%Y-%m-%d}" if moment.time() == dt.time() else f"{moment:%Y-%m-%d %H:%M:%S}"


def text(value) -> str | None:
    """A cell as the text the VBA reads: trimmed, None when empty; True/False for a boolean, 60 not 60.0."""
    if value is None:
        return None
    if isinstance(value, bool):
        return "True" if value else "False"
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value).strip(WHITE) or None


def _whole(value):
    """A Workflows number (Order, IsStart, IsEnd, IsFree): 1 and 1.0 and TRUE and "1" are 1; empty and "-" are None.
    Anything else stays as it is, for the check to name it (bool("0") would be True: never let text through)."""
    if value is None or isinstance(value, bool):
        return None if value is None else int(value)
    if isinstance(value, (int, float)):
        return int(value) if float(value).is_integer() else value
    cleaned = str(value).strip(WHITE)
    if cleaned in ("", "-"):
        return None
    return int(cleaned) if cleaned.isdecimal() else cleaned


def _cell(spec, column: str, value, parameter: str | None):
    if (spec.name, column) in DATE_COLUMNS and isinstance(value, (int, float)) and not isinstance(value, bool):
        return serial_to_iso(value)
    if column in spec.int_columns:
        return _whole(value)
    if parameter == "AsOf" and column == "Value" and isinstance(value, (int, float)) and not isinstance(value, bool):
        return serial_to_iso(value)
    return text(value)


def sheet_from_grid(spec, grid: list[list], problems: list[str]) -> ParamSheet:
    """The rows of a sheet as a ParamSheet with the columns of `spec`, found by header name."""
    header = [text(h) for h in (grid[0] if grid else [])]
    positions = []
    for column in spec.columns:
        if column in header:
            positions.append(header.index(column))
        else:
            positions.append(None)
            if (spec.name, column) not in OPTIONAL:
                problems.append(f"Sheet {spec.name}: no column {column} in its first row")
    rows = []
    for raw in grid[1:]:
        cells = [raw[i] if i is not None and i < len(raw) else None for i in positions]
        if all(text(c) is None for c in cells):
            continue
        parameter = text(cells[0]) if spec.name == "Config" else None
        rows.append(tuple(_cell(spec, column, value, parameter) for column, value in zip(spec.columns, cells)))
    return ParamSheet(spec.name, spec.columns, tuple(rows))


def complete_config(config: ParamSheet) -> ParamSheet:
    """Every parameter of CONFIG_PARAMETERS, in that order: a row missing from an older workbook is empty."""
    by_name = {row[0]: row for row in config.rows if row[0]}
    empty = (None,) * (len(config.header) - 1)
    return ParamSheet(config.name, config.header, tuple(by_name.get(name, (name,) + empty) for name in CONFIG_PARAMETERS))


def config_values(config: ParamSheet) -> dict[str, str]:
    """Parameter -> value, the default of the VBA when it is empty (ConfigValue)."""
    column = config.header.index("Value")
    return {row[0]: row[column] or DEFAULTS.get(row[0], "") for row in config.rows}


def missing_mandatory(config: ParamSheet) -> list[str]:
    """The parameters whose Mandatory is yes and whose Value is empty (modParams.MissingMandatory)."""
    value, flag = config.header.index("Value"), config.header.index("Mandatory")
    return [row[0] for row in config.rows if is_yes(row[flag]) and not row[value]]


def read_params(path: Path) -> list[ParamSheet]:
    """The parameter sheets of the workbook at `path`, in workbook order. Raises ParamsError listing every problem."""
    grids = read_sheets(path, set(SHEETS))
    problems: list[str] = []
    sheets: dict[str, ParamSheet] = {}
    for spec in (s for s in SPECS if s.name in SHEETS):
        if spec.name not in grids:
            problems.append(f"The workbook has no sheet {spec.name}")
            continue
        sheets[spec.name] = sheet_from_grid(spec, grids[spec.name], problems)
    if "Config" in sheets:
        sheets["Config"] = complete_config(sheets["Config"])
        problems += config_problems(sheets["Config"])
    problems += shape_problems(sheets)
    if problems:
        raise ParamsError("\n".join(problems))
    return [sheets[name] for name in SHEETS]


def config_problems(config: ParamSheet) -> list[str]:
    """The Config values the VBA would refuse, or would compute wrong from (a value is never refused for being empty
    here: missing_mandatory names those)."""
    values, problems = config_values(config), []
    if values["FlowLoadMode"].strip().upper() not in ("A", "B"):
        problems.append(f"Config: FlowLoadMode must be A or B, got {values['FlowLoadMode']!r}")
    threshold = values["FillRateThreshold"]
    if threshold and not (threshold.isdecimal() and 0 <= int(threshold) <= 100):
        problems.append(f"Config: FillRateThreshold must be a whole number between 0 and 100, got {threshold!r}")
    if values["AsOf"] and _parse_moment(values["AsOf"]) is None:
        problems.append("Config: AsOf must be a UTC date or date and time such as 2026-11-03 18:00, "
                        f"got {values['AsOf']!r}")
    base = values["BaseUrl"]
    if base and "@" in base.split("//", 1)[-1].split("/", 1)[0]:
        problems.append("Config: BaseUrl must not contain a user name or a token (user:token@host)")
    return problems


@dataclass(frozen=True)
class Connection:
    base_url: str
    email: str
    token: str
    proxy_server: str                  # "" = find it in the Windows settings; "DIRECT" = no proxy
    proxy_bypass: str
    windows_auth: bool                 # answer the proxy's Windows authentication with the Windows session
    timeout_s: int
    max_retries: int


@dataclass(frozen=True)
class RunSettings:
    sheets: list[ParamSheet]
    connection: Connection
    filter: str
    work_type_field: str
    fields: dict[str, str]             # FIELD_PARAMETERS -> custom field id ("" = not configured)
    blocks_link_type: str
    as_of: dt.datetime
    mode: str                          # FlowLoadMode, A or B
    periods: list[Period]
    projects: list[tuple[str, str, str]]   # Project | Role | Team
    rules: Rules
    moves: Moves
    checks: CheckSettings


def _number(value: str, default: int, minimum: int) -> int:
    """The VBA reads it with Val: no number gives 0; below `minimum` gives `default` (timeout) or the minimum (retries)."""
    digits = value.strip()
    number = int(digits) if digits.lstrip("-").isdecimal() else 0
    return number if number >= minimum else default


def run_settings(sheets: list[ParamSheet]) -> RunSettings:
    """What a run needs. Raises ParamsError with the message of the VBA when a mandatory value or a period is missing."""
    by_name = {s.name: s for s in sheets}
    config = by_name["Config"]
    missing = missing_mandatory(config)
    if missing:
        raise ParamsError(f"Fill these values of the Config sheet first (Mandatory = True): {', '.join(missing)}")
    periods = [Period(name, dt.datetime.fromisoformat(start), dt.datetime.fromisoformat(end))
               for name, start, end in by_name["Periods"].rows]
    if not periods:
        raise ParamsError("Sheet Periods has no period: type the iterations (Period | Start | End) to compute")
    values = config_values(config)
    as_of = _parse_moment(values["AsOf"])
    if as_of is None:
        raise ParamsError("Config: AsOf is empty: type the UTC date and time of the extraction, for example "
                          "2026-11-03 18:00")
    connection = Connection(
        base_url=values["BaseUrl"], email=values["Email"], token=values["ApiToken"],
        proxy_server=values["ProxyServer"], proxy_bypass=values["ProxyBypass"],
        windows_auth=is_yes(values["ProxyWindowsAuth"]),
        timeout_s=_number(values["TimeoutSeconds"], 60, 1), max_retries=_number(values["MaxRetries"], 0, 0))
    return RunSettings(
        sheets=sheets, connection=connection, filter=values["Filter"], work_type_field=values["WorkTypeField"],
        fields={name: values[name] for name in FIELD_PARAMETERS}, blocks_link_type=values["BlocksLinkType"],
        as_of=as_of, mode=values["FlowLoadMode"].strip().upper(), periods=periods,
        projects=[tuple(row) for row in by_name["Projects"].rows],
        rules=Rules.from_params(sheets), moves=Moves.from_params(sheets), checks=CheckSettings.from_params(sheets))
