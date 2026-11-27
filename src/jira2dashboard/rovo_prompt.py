"""The prompt for Rovo: what to ask it so that it lists the values to type in the workbook (pure text).

flow/sheets.py writes the rows of `spec_rows` into the sheet RovoPromptSpec at build time. The button "Prompt for Rovo"
(modParams.WriteRovoPrompt) reads that sheet and writes the prompt into the sheet RovoPrompt, one line per row:
the lines of the parts `text` as they are, the part `config` replaced by one line per row of the sheet Config (name,
mandatory or optional, description, example: never its Value, which may hold the address, the e-mail or the token), and each
part `sheet` as "SHEET <name> - columns: <header of the sheet>. <rule>". The user copies the sheet into Rovo and exports
its answer as CSV (Sheet, Row, Column, Value, Note).
AL, 2026-11-23: the prompt asks only what the workbook cannot read in Jira by itself. It names no connection parameter at all (part `keep`: the
load of the CSV refuses them), no parameter that "Propose values from Jira" reads (part `skip`: the field ids, the lists of issue types, the
link type) and no sheet Iterations (the sprints) nor Transitions (the moves of the workflows).
"""
from jira2dashboard import sheet_names as S
from jira2dashboard.flow.checks import FIELD_PARAMETERS

SPEC_SHEET = "RovoPromptSpec"
PROMPT_SHEET = "RovoPrompt"
SPEC_HEADER = ("Part", "Name", "Text")
TEXT, CONFIG, SHEET, KEEP, SKIP, LOAD = "text", "config", "sheet", "keep", "skip", "load"
# The parameters that say where the credentials go: the CSV can tell them to the user but the button Load settings CSV never writes
# them (a CSV, made by Rovo from what it reads in Jira, must not be able to redirect the e-mail and the token to another host).
PROTECTED = ("BaseUrl", "Email", "ApiToken", "ProxyServer", "ProxyBypass", "ProxyWindowsAuth")
# The settings of the connection to Jira that the CSV may load (the rights of the account, the HTTP calls): not asked of Rovo either.
CONNECTION_SETTINGS = ("IsUserProjectAdmin", "TimeoutSeconds", "MaxRetries")
# The parameters of Config that the workbook reads in Jira by itself (the button Propose values from Jira): not asked of Rovo.
# The sheets that the CSV may still fill (the configuration CSV of the fake Jira, a CSV written by hand) though the prompt does not ask for them:
# the workbook reads them in Jira (the sprints, the moves of the workflows).
LOADED_TOO = ("Iterations", "Transitions")
FROM_JIRA = ("WorkTypeField", *FIELD_PARAMETERS, "FeatureTypes", "SubTaskTypes", "ParentRequiredTypes", "PointsRequiredTypes", "BlocksLinkType")
CSV_COLUMNS = ("Sheet", "Row", "Column", "Value", "Note")
UNKNOWN = "UNKNOWN"

INTRO = (
    "I am filling an Excel workbook (jira2dashboard) that reads our Jira Cloud site and computes the six SAFe flow metrics: "
    "Velocity, Flow Time, Flow Load, Flow Efficiency, Flow Distribution and Flow Predictability.",
    "The workbook reads by itself in Jira the statuses, the workflows and their moves, the custom fields, the issue types, the link "
    "types and the sprints: do not give them. Using what you can see in our Jira (projects, teams, Program Increments, the statuses "
    "where the work waits or that are no delivery), give me the value to type for each parameter below.",
    "Rules:",
    f"- Answer only from what you can see in Jira. When you cannot see a value, write {UNKNOWN} in Value and say in Note what "
    "I must check.",
    "- Never invent a field id, a status name or a date.",
    "- Never write an API token or a password.",
    "- This text holds no value of my workbook: only the names and the meaning of what to fill.",
    "",
)
CONFIG_TITLE = (f"SHEET {S.of('Config')} - one line per parameter: name, mandatory or optional, meaning, and an example of the form "
                "of the value:")
SHEET_RULES = {
    "Projects": "One row per Jira project to extract. Project = the Jira project key. Role = Program for the project that "
                "holds the Features (exactly one), Team for the others. Team = the team name shown in the metrics (optional: "
                "empty = the project key; for the Program project, the name of the train).",
    "PIs": "One row per Program Increment (the long cycle that holds several iterations). PI = the value of the PI field of "
           "Jira, exactly as Jira writes it (for example PI#01-2026): ONE PI per row, never a list of PIs in one cell. Start = first day (yyyy-mm-dd, included). "
           "End = the day the next PI starts (yyyy-mm-dd, excluded). A PI whose dates you cannot see still gets a row of its own, with "
           f"{UNKNOWN} in Start and End.",
    "Workflows": "One row per status for which the workbook must NOT simply take what Jira says (it reads the status categories and the "
                 "workflows from Jira by itself). Status = the status name exactly as in Jira. Workflow = Team or Feature to limit the "
                 "row to one of them, empty for both. Leave a cell empty to keep Jira's value. "
                 "Category = Backlog, Active (work in progress), Wait (waiting), Done, Removed (cancelled): Jira cannot say Wait or "
                 "Removed, and calls Done a status that is no delivery (cancelled, to deploy). IsEnd = 1 for a status that ends the "
                 "work (every Done status does by default), IsFree = 1 for a status reachable from any status; 0 or empty otherwise. "
                 "The work starts at the first status in progress: no column for it.",
}
FORMAT = (
    "",
    "Answer with ONE table that I can export as CSV (comma separated, UTF-8, one header row, a value that contains a comma "
    "between double quotes), with exactly these columns:",
    ",".join(CSV_COLUMNS),
    "- Sheet: " + ", ".join(S.of(name) for name in ("Config", *SHEET_RULES)) + ".",
    "- Row: empty for " + S.of("Config") + "; 1, 2, 3... for the other sheets (one number per line of the sheet).",
    "- Column: the parameter name for " + S.of("Config") + "; the column name for the other sheets.",
    "- Value: what to type in the cell.",
    "- Note: where you found it in Jira, or what I must check. Keep it short.",
    "Give one line for every parameter of " + S.of("Config") + " listed above, even when the Value is " + UNKNOWN + ", and "
    "one line for every cell of the other sheets. No text outside the table.",
)


def spec_rows() -> list[tuple[str, str, str]]:
    """The rows of the sheet RovoPromptSpec: (Part, Name, Text), in the order of the prompt."""
    rows = [(TEXT, "", line) for line in INTRO]
    rows.append((TEXT, "", CONFIG_TITLE))
    rows.append((CONFIG, "", ""))
    rows.append((TEXT, "", ""))
    rows += [(SHEET, S.of(name), rule) for name, rule in SHEET_RULES.items()]
    rows += [(TEXT, "", line) for line in FORMAT]
    rows += [(KEEP, name, "connection parameter: typed by hand, never loaded from the CSV, not in the prompt") for name in PROTECTED]
    rows += [(LOAD, S.of(name), "read in Jira by the workbook: not in the prompt, the CSV may fill it") for name in LOADED_TOO]
    rows += [(SKIP, name, "setting of the connection to Jira: not in the prompt") for name in CONNECTION_SETTINGS]
    rows += [(SKIP, name, "read in Jira by the workbook (Propose values from Jira): not in the prompt") for name in FROM_JIRA]
    return rows

