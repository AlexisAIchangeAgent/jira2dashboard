"""The user manual of the delivered workbook: the first steps on the Dashboard and the sheet Manual.

Pure text. flow/sheets.py writes it at build time (`write_manual`). The button captions, the list of mandatory parameters and the counts of the
controls and data checks come from their sources, so the manual follows the code. The sheets are named by their real names
(`Config_from_user`): the suffix says where the data of a sheet comes from (sheet_names.py). AL, 2026-11-24: short lines, in the order of the
buttons, and a section to check the figures.
"""
from jira2dashboard import sheet_names as S
from jira2dashboard.flow.checks import CHECKS
from jira2dashboard.flow.dashboard_spec import (CANCEL_CAPTION, COMPUTE_CAPTION, CONFIG_CAPTION, DRAW_CAPTION, EXTRACT_CAPTION, PROPOSE_CAPTION,
                                                ROVO_CAPTION, ROVO_LOAD_CAPTION, STATUS_CELL)
from jira2dashboard.flow.reconciliation import CONTROLS
from jira2dashboard.params.model import TYPED_AT_CLIENT

MANUAL_SHEET = "Manual"
HEADING, TEXT = "heading", "text"

CONFIG, PROJECTS, ITERATIONS, PIS = S.of("Config"), S.of("Projects"), S.of("Iterations"), S.of("PIs")
WORKFLOWS, TRANSITIONS = S.of("Workflows"), S.of("Transitions")
CHARTS, METRICS, DATACHECKS, ITEMFLOW = S.of("Charts"), S.of("Metrics"), S.of("DataChecks"), S.of("ItemFlow")
RECONCILIATION, METRIC_ITEMS, STATUS_RULES = S.of("Reconciliation"), S.of("MetricItems"), S.of("StatusRules")
PROPOSAL_SHEETS = ", ".join(S.of(name) for name in S.JIRA if name.startswith("Jira"))
CHECK_SECTION = "Check the figures"

LEGEND = (f"Sheet names: {S.FROM_USER} = you type here, {S.FROM_JIRA} = read from Jira (check the data here), "
          f"{S.FROM_EXCEL} = calculated or drawn by Excel (check the results here).")


def first_steps() -> list[str]:
    """The block on the Dashboard: what to do the first time, in the order of the buttons. Line 1 is the title."""
    return [
        f"First time: do these steps in this order (details: sheet {S.of(MANUAL_SHEET)})",
        f"1. Sheet {CONFIG}: fill every value whose Mandatory is True. {', '.join(TYPED_AT_CLIENT)} are empty in a new workbook.",
        f"2. Sheet {PROJECTS}: your Jira projects, one of them the Program. Sheets {ITERATIONS} and {PIS}: your calendar. "
        f'"{ROVO_CAPTION}" and "{ROVO_LOAD_CAPTION}" can fill them with Rovo.',
        f'3. Click "{PROPOSE_CAPTION}" (or "{CONFIG_CAPTION}" to only read Jira), then check the yellow cells it wrote.',
        f"4. Sheet {WORKFLOWS}: type only what Jira cannot say (a status that waits, a Done status that is no delivery).",
        f'5. Click "{EXTRACT_CAPTION}" ("{CANCEL_CAPTION}" stops it), then "{COMPUTE_CAPTION}", then "{DRAW_CAPTION}". '
        "After a correction, click only the last two: Jira is not read again.",
        f'6. Read the charts in the sheet {CHARTS}. To check a figure: sheet {S.of(MANUAL_SHEET)}, section "{CHECK_SECTION}".',
        LEGEND,
        f"Before you send this workbook to anyone: empty ApiToken in the sheet {CONFIG}.",
    ]


def manual(mandatory: tuple[str, ...]) -> list[tuple[str, str]]:
    """The sheet Manual: (HEADING or TEXT, line). `mandatory`: the parameters of Config whose Mandatory is True."""
    sections = (
        ("Before the first run", [
            "A file from a download or an e-mail: right-click it > Properties > Unblock, then open it and click Enable Content.",
            "A button says the workbook is locked (J2D-SYS-003): click Enable Editing, or save a copy (File > Save As) and open it.",
            "It needs Excel for Microsoft 365 or Excel 2024 on Windows, with macros. Python is not needed.",
        ]),
        ("The sheets", [
            LEGEND,
            f"Do not edit {S.of('FlowFormulas')}, {S.of('ChartSpec')}, {S.of('ChecksSpec')}, {S.of('RawTypesSpec')} and "
            f"{S.of('RovoPromptSpec')}: they are the plan the workbook builds from.",
        ]),
        ("Set up (once)", [
            f"{CONFIG}: one row per parameter (Parameter, Value, Example, Mandatory, Description). Only Value is read; Example shows its form.",
            f"The mandatory parameters: {', '.join(mandatory)}.",
            "A button stops, before anything else, while a value it needs is empty, and names it.",
            "AsOf: the UTC date and time of the extraction, for example 2026-11-03 18:00. Open statuses run up to it.",
            "ProxyServer: empty = the proxy of Windows, DIRECT = no proxy, or host:port.",
            "ApiToken stays in clear in the sheet: empty it before you send the workbook to anyone.",
            f"{PROJECTS}: Project, Role, Team. One project is the Program (the Features); the others are Team projects (none for an ART of "
            "Features only).",
            f"{ITERATIONS}: Iteration-label, Start, End (End excluded, UTC). {PIS}: PI, Start, End, with names such as PI#01-2026.",
            f"{WORKFLOWS}: the exceptions to what Jira says (an empty cell takes Jira's value): Category = Wait for a status that waits (Jira has "
            "no such category), Removed for a Done status that is no delivery, IsStart = 1 when Jira's workflows cannot be read.",
            "Its column Workflow is a kind, not a Jira name: Feature = the tickets of the Program project, Team = those of the Team projects, empty "
            f"= both. The real Jira workflows of each status are in {STATUS_RULES}, column JiraWorkflows.",
            f"{TRANSITIONS}: extra moves to allow in the data check of the moves, when Jira's workflows cannot be read.",
        ]),
        (f'Button "{ROVO_CAPTION}"', [
            f"It writes in {S.of('RovoPrompt')} a prompt that asks Rovo only what the workbook cannot read in Jira (the projects, the PIs, the "
            "statuses that wait). It holds no value of this workbook. The lines are selected: Ctrl+C, then paste them into Rovo.",
            "Ask Rovo to answer as CSV (Sheet, Row, Column, Value, Note).",
        ]),
        (f'Button "{ROVO_LOAD_CAPTION}"', [
            f"Choose Rovo's CSV, then confirm. A parameter of {CONFIG} it names gets its Value; any other sheet of the file ({PROJECTS}, "
            f"{ITERATIONS}, {PIS}, {WORKFLOWS}, {TRANSITIONS}) has all its rows replaced.",
            "The connection parameters, UNKNOWN and empty values are never loaded. Save a copy first to keep the current values.",
        ]),
        (f'Button "{CONFIG_CAPTION}"', [
            f"It reads Jira (read only) for the projects of {PROJECTS} and writes {PROPOSAL_SHEETS}. It never changes the sheets you type.",
            "IsUserProjectAdmin = False: the workflows are not read (they need the project administrator right): declare the start statuses "
            f"(IsStart) in {WORKFLOWS}.",
        ]),
        (f'Button "{PROPOSE_CAPTION}"', [
            f'It does what "{CONFIG_CAPTION}" does, then fills the EMPTY cells of {CONFIG} and {ITERATIONS} with what Jira says: the field ids, '
            "the issue-type lists, BlocksLinkType, and the iterations made of the sprints of the scrum boards.",
            "Each cell it writes is yellow, with the note 'To confirm' on the first of a row. What you typed never changes. PI dates have no "
            "API: type them.",
        ]),
        (f'Button "{EXTRACT_CAPTION}"', [
            f"It reads Jira in slices: the bar moves, cell {STATUS_CELL} says what happens, Excel stays usable. \"{CANCEL_CAPTION}\" stops it; "
            "stopped before the raw sheets are written, the previous results stay.",
            f"It writes the raw sheets ({S.of('Issues')}, {S.of('StatusIntervals')}, {S.of('Changelog')}...) and deletes the previous tables "
            f'and charts: never old figures next to new data. Then click "{COMPUTE_CAPTION}".',
        ]),
        (f'Button "{COMPUTE_CAPTION}"', [
            f"It makes the status rules again ({STATUS_RULES}: Jira's, under the exceptions of {WORKFLOWS}), then every table and figure. It "
            f'never calls Jira and deletes the charts: then click "{DRAW_CAPTION}".',
            f"Its status in {STATUS_CELL} adds a line for each thing to look at: No start status, No Wait status, Done outside the iterations, "
            "Without a team, Not Features, Statuses not in Workflows, AsOf before the data, VelocityCheck gap, Reconciliation (controls to "
            "review), Cells in error.",
        ]),
        (f'Button "{DRAW_CAPTION}"', [
            f"It draws the charts in {CHARTS} (cell B2: the team of the Flow Distribution chart) and {S.of('IterationCharts')}, without "
            "calling Jira.",
            "The note under a chart counts its gaps (not computable points) and gives their most frequent cause, read in the check tables.",
        ]),
        ("Which button after a change", [
            f'The connection, the proxy, Filter, the projects of {PROJECTS}, a field id (...Field, WorkTypeField) or BlocksLinkType: '
            f'"{EXTRACT_CAPTION}", then the two others.',
            f"{WORKFLOWS}, {TRANSITIONS}, {ITERATIONS}, {PIS}, the Role or Team of {PROJECTS}, AsOf, FeatureTypes or a mode: "
            f'"{COMPUTE_CAPTION}", then "{DRAW_CAPTION}".',
        ]),
        ("What the figures mean", [
            f'{METRICS}: one row per iteration and team. A figure that cannot be computed reads "not computable", never 0. The ART counts the '
            "Features only (FeatureTypes); the other tickets of the Program are in the team ART (other types).",
            f"{S.of('PIMetrics')}: per PI, the Features engaged at its start, those released in it, and Feature Predictability (released / "
            "engaged; mode B also counts the added ones). Flow Predictability stays not computable: Jira does not hold the PI objectives.",
            f"{S.of('IterationFeatures')}: per iteration, PI and team, the Features engaged, added, deferred, released and carried over.",
            f"{S.of('PIStatusTime')}: per PI, the share of the cycle time of the released Features spent in each status (the chart Cycle time by "
            "status).",
            f"{DATACHECKS}: the {len(CHECKS)} data checks {CHECKS[0].id}-{CHECKS[-1].id}, with Items and OK, KO or not computable. "
            f"Filter {ITEMFLOW} on a column DCxx = 1 to see the tickets.",
        ]),
        (CHECK_SECTION, [
            f"1. After \"{COMPUTE_CAPTION}\", read the status in {STATUS_CELL} and the rows under it: they say what to look at first.",
            f"2. {RECONCILIATION}: {len(CONTROLS)} controls with Expected, Found and Result. OK = right, REVIEW = look at it, INFO = a point of "
            "attention. Below the table, the parameters in force (FeatureTypes, the start statuses, the Wait statuses, AsOf, the dates of the "
            "iterations): compare them with Jira. On the right, the ticket types counted and the statuses no rule knows.",
            "3. One check table per metric, one row per iteration and team. Gap must be 0 (else tell the support). Reason says why a point is "
            "not computable; Cause is the same without the counts.",
            f"{S.of('VelocityCheck')}: why the Velocity differs from Jira's count of the changes into an end status (created done, closed then "
            "reopened, other types of the Program).",
            f"{S.of('FlowTimeCheck')}: the done items, then those left out of Flow Time: no start status in the rules, never entered the start "
            "status, back to the backlog after the start (a To Do status after it: declare a later status IsStart, or the backlog one Wait), reopened.",
            f"{S.of('FlowEfficiencyCheck')}: the active and wait days, the Wait statuses declared. {S.of('FlowLoadCheck')}: the days measured, the "
            f"start statuses, the load at the end. {S.of('FlowDistCheck')}: the done items per work type, Unmapped.",
            f"4. {METRIC_ITEMS}: the tickets behind each figure, four lists side by side. Done in an iteration (filter Iteration and Team: the rows "
            "are the Velocity; with FlowTimeCounted, WhyNotInFlowTime, the active, wait and unclassified days, the work type).",
            "Then the done tickets that no figure counts, with the reason; those done within 24 hours of an iteration boundary; those in process "
            "at the end of each iteration (the Flow Load at the end).",
            f"5. To compare with Jira: take an iteration and a team in {METRIC_ITEMS}, then search Jira for status CHANGED TO \"<end status>\" "
            "DURING (\"<start>\", \"<end>\"). VelocityCheck explains the differences. Dates are UTC; the End of an iteration is excluded.",
            f"6. Most \"not computable\" come from {STATUS_RULES}: no status with IsStart = 1, or none in the category Wait. Correct it in "
            f'{WORKFLOWS}, then click "{COMPUTE_CAPTION}" and "{DRAW_CAPTION}": no new extraction.',
            f"7. {ITEMFLOW}: one row per ticket, with Account (every ticket is in one account), DoneIteration and WhyNotInFlowTime.",
        ]),
        ("When something stops", [
            'A status that starts with "Stopped" gives the error and its code (J2D-...): give the code to the support. '
            f"{S.of('ErrorCodes')} says what each code means and what to check.",
            r"The log: logs\jira2dashboard.log in the folder of the workbook, one line per step with its time (no token, e-mail, address or "
            f"Jira answer). The line 'status rules' counts the start and Wait statuses. {S.of('Steps')} lists the steps of the last run.",
            "After an error Notepad opens the log, unless OpenLogOnError is No. LogLevel: Steps (each step) or Errors.",
        ]),
    )
    rows = [(HEADING, "jira2dashboard - user manual")]
    for title, lines in sections:
        rows.append((HEADING, title))
        rows.extend((TEXT, line) for line in lines)
    return rows
