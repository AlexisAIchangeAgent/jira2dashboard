"""The user manual of the delivered workbook: the first steps on the Dashboard and the sheet Manual.

Pure text. flow/sheets.py writes it at build time (`write_manual`). The button captions and the list of mandatory
parameters come from their sources, so the manual follows the code. The sheets are named by their real names
(`Config_from_user`): the suffix says where the data of a sheet comes from (sheet_names.py).
"""
from jira2dashboard import sheet_names as S
from jira2dashboard.flow.dashboard_spec import (CANCEL_CAPTION, CONFIG_CAPTION, EXTRACT_CAPTION, PROPOSE_CAPTION,
                                                ROVO_CAPTION, ROVO_LOAD_CAPTION, STATUS_CELL)
from jira2dashboard.params.model import TYPED_AT_CLIENT

MANUAL_SHEET = "Manual"
HEADING, TEXT = "heading", "text"

CONFIG, PROJECTS, ITERATIONS, PIS = S.of("Config"), S.of("Projects"), S.of("Iterations"), S.of("PIs")
WORKFLOWS, TRANSITIONS = S.of("Workflows"), S.of("Transitions")
CHARTS, METRICS, DATACHECKS, ITEMFLOW = S.of("Charts"), S.of("Metrics"), S.of("DataChecks"), S.of("ItemFlow")
PROPOSAL_SHEETS = ", ".join(S.of(name) for name in S.JIRA if name.startswith("Jira"))

LEGEND = (f"Sheet names: {S.FROM_USER} = you type here, {S.FROM_JIRA} = read from Jira (check the data here), "
          f"{S.FROM_EXCEL} = calculated or drawn by Excel (check the results here).")


def first_steps() -> list[str]:
    """The block on the Dashboard: what to do the first time, in the order of the buttons. Line 1 is the title."""
    return [
        f"First time: do these steps in this order (details: sheet {S.of(MANUAL_SHEET)})",
        f"1. Sheet {CONFIG}: fill every value whose Mandatory is True. {', '.join(TYPED_AT_CLIENT)} are empty in a new "
        "workbook.",
        f"2. Sheet {PROJECTS}: the Jira projects, and which one is the Program. Sheet {ITERATIONS}: the iterations. "
        f'"{ROVO_CAPTION}" prepares a prompt for Rovo that lists the values to find, '
        f'"{ROVO_LOAD_CAPTION}" loads its answer (CSV).',
        f'3. Click "{PROPOSE_CAPTION}" (or "{CONFIG_CAPTION}" to only read). Read the proposal sheets {PROPOSAL_SHEETS}.',
        f"4. Check the yellow cells it wrote in the sheets {WORKFLOWS}, {TRANSITIONS} and {CONFIG}, then type what is left.",
        f'5. Click "{EXTRACT_CAPTION}" and wait for the bar to reach the end. "{CANCEL_CAPTION}" stops it.',
        f"6. Read the charts in the sheet {CHARTS}, the figures in the sheet {METRICS} and the data checks in the sheet {DATACHECKS}.",
        LEGEND,
        f"Before you send this workbook to anyone: empty ApiToken in the sheet {CONFIG}.",
    ]


def manual(mandatory: tuple[str, ...]) -> list[tuple[str, str]]:
    """The sheet Manual: (HEADING or TEXT, line). `mandatory`: the parameters of Config whose Mandatory is True."""
    sections = (
        ("Before the first run", [
            "If the file came from a download or an e-mail: right-click it > Properties > Unblock. Then open it and "
            "click Enable Content.",
            "It needs Excel for Windows with macros. Python is not needed.",
        ]),
        ("Where to type, where to check", [
            f"{S.FROM_USER}: what you type, or what a button proposes and you confirm: " + ", ".join(S.of(n) for n in S.USER) + ".",
            f"{S.FROM_JIRA}: what the Jira API answered, to check against Jira: " + ", ".join(S.of(n) for n in S.JIRA) + ". "
            f"The sheets {S.of('Issues')}, {S.of('StatusIntervals')} and {S.of('IssueLinks')} are rewritten at each extraction; "
            f"the others (Jira...) are the proposals of the button \"{CONFIG_CAPTION}\".",
            f"{S.FROM_EXCEL}: calculated by Excel formulas and macros, or drawn: the charts ({CHARTS}), the flow metrics "
            f"({METRICS}, {S.of('PIMetrics')}), the work items ({ITEMFLOW}, {S.of('DailyLoad')}), the data checks "
            f"({DATACHECKS}, {S.of('FillRates')}), the {S.of('Dashboard')}, the {S.of('Steps')} of the last run and the "
            f"{S.of('ErrorCodes')}. {S.of('FlowFormulas')}, {S.of('ChartSpec')}, {S.of('ChecksSpec')} and "
            f"{S.of('RovoPromptSpec')} are the plan the workbook builds from: do not edit them.",
        ]),
        (f"Sheet {CONFIG}", [
            "One row per parameter: Parameter, Value, Example, Mandatory, Description.",
            "Example shows the expected form of the value (it is never read: only Value counts).",
            f'Mandatory = True: "{EXTRACT_CAPTION}" stops before calling Jira while the value is empty, and names it. '
            f'"{CONFIG_CAPTION}" needs only BaseUrl, Email and ApiToken.',
            'Mandatory = False: an empty value is allowed. A default is used, nothing is added (no filter, no proxy), '
            'or the data checks that need the value show "not computable".',
            f"The mandatory parameters: {', '.join(mandatory)}.",
            "AsOf: the UTC date and time of the extraction, for example 2026-11-03 18:00. Open statuses run up to it.",
            "ProxyServer: empty = the proxy of the Windows settings, found as the browser finds it (automatic "
            "detection, proxy script or fixed proxy). DIRECT = no proxy. Or type host:port.",
            "ApiToken stays in clear in this sheet: empty it before you send the workbook to anyone.",
        ]),
        (f"Sheets {PROJECTS}, {ITERATIONS} and {PIS}", [
            "Projects: Project, Role, Team. Exactly one project is the Program (the features); the others are Team.",
            "Iterations: Iteration-label, Start, End (the end is excluded). They are the short cycles inside a PI, for example 6 of "
            "3 weeks in a PI of 18 weeks, and the iterations of the charts. Column Example shows the form of each value.",
            "PIs: PI, Start, End, with names such as PI#01-2026 (data check DC18).",
        ]),
        (f'Button "{CONFIG_CAPTION}"', [
            f"It reads Jira (read only) for the projects of the sheet {PROJECTS} and writes the proposal sheets "
            f"{PROPOSAL_SHEETS}. It never changes the parameter sheets.",
            "Decide by hand: Order, IsStart, IsEnd, Active or Wait, Removed instead of Done, and the Type of a "
            f"transition. Then copy into {WORKFLOWS} and {TRANSITIONS}, and the field ids into the ...Field rows of {CONFIG}.",
        ]),
        (f'Button "{PROPOSE_CAPTION}"', [
            f'It does what "{CONFIG_CAPTION}" does, then copies what fits of the proposal sheets into the sheets {WORKFLOWS}, '
            f"{TRANSITIONS} and {CONFIG}: the statuses and moves the sheets lack, and the ids of the fields that Jira names once "
            "(Sprint, Team, Story Points, Target start and end, PI, Business Value).",
            "It writes only into EMPTY cells: what you typed is never changed. Every cell it writes is yellow, the first cell of a "
            "row has the note 'To confirm': check it, then clear the colour if you like. Order, IsStart, IsEnd, Active or Wait, "
            "Removed instead of Done and the Type of a move stay yours to type.",
            f"IsUserProjectAdmin in {CONFIG}: True (default) = you administer the Jira projects, the full workflows are read. "
            "False = only the statuses and their categories are read (the right to browse is enough), without the moves.",
        ]),
        (f'Button "{ROVO_CAPTION}"', [
            f"It writes the sheet {S.of('RovoPrompt')}: a prompt that asks Rovo for every value to type in this workbook (the connection, "
            "the flow metrics, the projects, iterations, PIs, workflows and transitions). The lines are selected: Ctrl+C, "
            "then paste them into Rovo.",
            "Rovo answers with a table; export it as CSV (Sheet, Row, Column, Value, Note) and type the values in the "
            "sheets. The prompt holds no value of this workbook (no address, e-mail or token): only the names and the meaning "
            "of the parameters. A value Rovo cannot see is marked UNKNOWN.",
        ]),
        (f'Button "{ROVO_LOAD_CAPTION}"', [
            "A dialog opens in the folder of this workbook: choose the CSV that Rovo answered, in any folder. The button "
            f"then asks you to confirm. The Value of each parameter of {CONFIG} that the "
            f"file gives replaces the current one; for every other sheet in the file ({PROJECTS}, {ITERATIONS}, {PIS}, {WORKFLOWS}, "
            f"{TRANSITIONS}) all the current rows are replaced by its rows (the column Example stays).",
            "The connection parameters (BaseUrl, Email, ApiToken, ProxyServer, ProxyBypass, ProxyWindowsAuth) are never loaded: "
            "type them by hand, so that a CSV cannot send the credentials elsewhere. "
            f"A value UNKNOWN, an empty value and the ApiToken placeholder are not loaded: {CONFIG} keeps its value, a table "
            "cell stays empty. The summary on the Dashboard says what was loaded and the first problem, if any. Save a copy "
            "of the workbook first if you want to keep the current values.",
        ]),
        (f'Button "{EXTRACT_CAPTION}"', [
            f"It reads Jira in small slices: the bar moves, cell {STATUS_CELL} (under the buttons) says what happens, and "
            f"Excel stays usable. At the end, {STATUS_CELL} gives the result and the rows under it its details, one per line "
            "(name in column A, value in column B).",
            f'"{CANCEL_CAPTION}" stops it before the raw sheets are written: the previous results stay.',
            f"At the end the sheets {ITEMFLOW}, {S.of('DailyLoad')}, {METRICS}, {S.of('PIMetrics')}, {S.of('ChartData')} and the charts are rebuilt to the size of "
            "the new data. Run it as often as needed.",
            f"The charts are in the sheet {CHARTS}, right of the Dashboard. Its cell B2: the team of the Flow Distribution chart.",
            f"Feature Predictability, per PI of the sheet {PIS}: the Features done before the end of the PI over the Features "
            "planned in it (their PI field). A cancelled Feature stays planned, not done. It needs FeatureTypes, the PI field "
            f"(PIField) and the sheet {PIS}. A PI that Features carry but the sheet {PIS} lacks comes after them, not computable "
            f"(it has no end date): add it to the sheet {PIS}. Flow Predictability, the SAFe measure on the business value of the "
            "PI objectives, stays not computable: Jira does not hold the PI objectives.",
            f'A status that starts with "Stopped" gives the error, with its code (J2D-...): read it to the support. The sheet {S.of("ErrorCodes")} '
            'says what each code means and what to check. The local log tells more (below).',
        ]),
        ("Results", [
            f'{METRICS}: one row per iteration and team. A figure that cannot be computed reads "not computable", never 0.',
            f"{S.of('PIMetrics')}: one row per PI of the sheet {PIS}, with its dates, the Features planned in it, those done before its "
            f"end and Feature Predictability; then one row per PI of a Feature missing from the sheet {PIS}, without dates.",
            f"{DATACHECKS}: the 29 data checks DC01-DC29, with Items and OK, KO or not computable. To see the work items, "
            f"filter the table ItemFlow (sheet {ITEMFLOW}) on its column DCxx = 1.",
            f"Do not edit the sheets {S.of('FlowFormulas')}, {S.of('ChartSpec')} and {S.of('ChecksSpec')}: they describe the formulas, the charts and "
            "the checks.",
        ]),
        ("Local log", [
            r"logs\jira2dashboard.log, in the folder of the workbook: one line per step, with its time in seconds, and "
            "the parameters of the run (PARAM lines: the values you typed for the connection, the proxy and the filter read "
            "only (set) or (empty)), and a RETRY line when Jira asked to wait. The first line names the version of the "
            "workbook, the last line of a stopped run gives its error code. It holds no token, e-mail, "
            f"address or Jira answer. LogLevel in Config: Steps (each step) or Errors. The sheet {S.of('Steps')} lists the steps of the last run with their status and time. After an error the Dashboard names the log file and "
            "Notepad opens it, unless OpenLogOnError is No.",
        ]),
    )
    rows = [(HEADING, "jira2dashboard - user manual")]
    for title, lines in sections:
        rows.append((HEADING, title))
        rows.extend((TEXT, line) for line in lines)
    return rows
