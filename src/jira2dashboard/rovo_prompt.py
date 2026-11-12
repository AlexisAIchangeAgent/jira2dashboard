"""The prompt for Rovo: what to ask it so that it lists the values to type in the workbook (pure text).

flow/sheets.py writes the rows of `spec_rows` into the sheet RovoPromptSpec at build time. The button "Prompt for Rovo"
(modParams.WriteRovoPrompt) reads that sheet and writes the prompt into the sheet RovoPrompt, one line per row:
the lines of the parts `text` as they are, the part `config` replaced by one line per row of the sheet Config (name,
mandatory or optional, description, example: never its Value, which may hold the address, the e-mail or the token), and each
part `sheet` as "SHEET <name> - columns: <header of the sheet>. <rule>". The user copies the sheet into Rovo and exports
its answer as CSV (Sheet, Row, Column, Value, Note).
"""
SPEC_SHEET = "RovoPromptSpec"
PROMPT_SHEET = "RovoPrompt"
SPEC_HEADER = ("Part", "Name", "Text")
TEXT, CONFIG, SHEET, KEEP = "text", "config", "sheet", "keep"
# The parameters that say where the credentials go: the CSV can tell them to the user but the button Load settings CSV never writes
# them (a CSV, made by Rovo from what it reads in Jira, must not be able to redirect the e-mail and the token to another host).
PROTECTED = ("BaseUrl", "Email", "ApiToken", "ProxyServer", "ProxyBypass", "ProxyWindowsAuth")
CSV_COLUMNS = ("Sheet", "Row", "Column", "Value", "Note")
UNKNOWN = "UNKNOWN"

INTRO = (
    "I am filling an Excel workbook (jira2dashboard) that reads our Jira Cloud site and computes the six SAFe flow metrics: "
    "Velocity, Flow Time, Flow Load, Flow Efficiency, Flow Distribution and Flow Predictability.",
    "Using what you can see in our Jira (projects, work types, workflows and their statuses, custom fields, sprints, teams), "
    "give me the value to type for each parameter below.",
    "Rules:",
    f"- Answer only from what you can see in Jira. When you cannot see a value, write {UNKNOWN} in Value and say in Note what "
    "I must check.",
    "- Never invent a field id, a status name or a date.",
    "- Never write an API token or a password: for ApiToken write TO BE TYPED BY THE USER.",
    "- This text holds no value of my workbook: only the names and the meaning of what to fill.",
    "- The connection parameters (" + ", ".join(PROTECTED) + ") are typed by hand: give your value for information, the workbook "
    "does not load them from your answer.",
    "",
)
CONFIG_TITLE = ("SHEET Config - one line per parameter: name, mandatory or optional, meaning, and an example of the form "
                "of the value:")
SHEET_RULES = {
    "Projects": "One row per Jira project to extract. Project = the Jira project key. Role = Program for the project that "
                "holds the Features (exactly one), Team for the others. Team = the team name shown in the metrics (for the "
                "Program project, the name of the train).",
    "Iterations": "One row per iteration of the Agile Release Train, from the oldest to the latest you can see (the sprints "
                  "of the boards are a good source). Iteration-label = a short unique name such as PI#01-2026 It1. "
                  "Start = first day (yyyy-mm-dd, included). End = the day the next iteration starts (yyyy-mm-dd, excluded).",
    "PIs": "One row per Program Increment (the long cycle that holds several iterations). PI = the value of the PI field of "
           "Jira, of the form PI#01-2026. Start and End as for the iterations.",
    "Workflows": "One row per status of each workflow used by the issue types of the projects above (say in Note which issue "
                 "types use the workflow). Order = position in the normal flow (1 = first status, - when none). Category = "
                 "Backlog, Active (work in progress), Wait (waiting), Done, Removed (cancelled) or Unclassified. IsStart = 1 "
                 "for the status where the work starts, IsEnd = 1 for the final status, IsFree = 1 for a status reachable "
                 "from any status; 0 otherwise.",
    "Transitions": "One row per allowed move between two statuses of a workflow. Workflow, From, To, and Type = forward, "
                   "back, back to backlog or exit.",
}
FORMAT = (
    "",
    "Answer with ONE table that I can export as CSV (comma separated, UTF-8, one header row, a value that contains a comma "
    "between double quotes), with exactly these columns:",
    ",".join(CSV_COLUMNS),
    "- Sheet: Config, " + ", ".join(SHEET_RULES) + ".",
    "- Row: empty for Config; 1, 2, 3... for the other sheets (one number per line of the sheet).",
    "- Column: the parameter name for Config; the column name for the other sheets.",
    "- Value: what to type in the cell.",
    "- Note: where you found it in Jira, or what I must check. Keep it short.",
    "Give one line for every parameter of Config, even when the Value is " + UNKNOWN + ", and one line for every cell of the "
    "other sheets. No text outside the table.",
)


def spec_rows() -> list[tuple[str, str, str]]:
    """The rows of the sheet RovoPromptSpec: (Part, Name, Text), in the order of the prompt."""
    rows = [(TEXT, "", line) for line in INTRO]
    rows.append((TEXT, "", CONFIG_TITLE))
    rows.append((CONFIG, "", ""))
    rows.append((TEXT, "", ""))
    rows += [(SHEET, name, rule) for name, rule in SHEET_RULES.items()]
    rows += [(TEXT, "", line) for line in FORMAT]
    rows += [(KEEP, name, "connection parameter: typed by hand, never loaded from the CSV") for name in PROTECTED]
    return rows

