"""The user manual of the delivered workbook: the first steps on the Dashboard and the sheet Manual.

Pure text. flow/sheets.py writes it at build time (`write_manual`). The button captions and the list of mandatory
parameters come from their sources, so the manual follows the code.
"""
from jira2dashboard.flow.dashboard_spec import CANCEL_CAPTION, CONFIG_CAPTION, EXTRACT_CAPTION
from jira2dashboard.params.model import TYPED_AT_CLIENT

MANUAL_SHEET = "Manual"
HEADING, TEXT = "heading", "text"
PROPOSAL_SHEETS = ("JiraIssueTypes, JiraStatuses, JiraWorkflows, JiraTransitions, JiraFields, JiraProjectWorkflows, "
                   "JiraProjectFields and JiraFieldUsage")


def first_steps() -> list[str]:
    """The block on the Dashboard: what to do the first time, in the order of the buttons. Line 1 is the title."""
    return [
        f"First time: do these steps in this order (details: sheet {MANUAL_SHEET})",
        f"1. Sheet Config: fill every value whose Mandatory is True. {', '.join(TYPED_AT_CLIENT)} are empty in a new "
        "workbook.",
        "2. Sheet Projects: the Jira projects, and which one is the Program. Sheet Iterations: the iterations.",
        f'3. Click "{CONFIG_CAPTION}". Read the proposal sheets {PROPOSAL_SHEETS}.',
        "4. Copy what fits into the sheets Workflows and Transitions, and the field ids into the ...Field rows of Config.",
        f'5. Click "{EXTRACT_CAPTION}" and wait for the bar to reach the end. "{CANCEL_CAPTION}" stops it.',
        "6. Read the charts here, the figures in the sheet Metrics and the data checks in the sheet DataChecks.",
        "Before you send this workbook to anyone: empty ApiToken in the sheet Config.",
    ]


def manual(mandatory: tuple[str, ...]) -> list[tuple[str, str]]:
    """The sheet Manual: (HEADING or TEXT, line). `mandatory`: the parameters of Config whose Mandatory is True."""
    sections = (
        ("Before the first run", [
            "If the file came from a download or an e-mail: right-click it > Properties > Unblock. Then open it and "
            "click Enable Content.",
            "It needs Excel for Windows with macros. Python is not needed.",
        ]),
        ("Sheet Config", [
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
        ("Sheets Projects, Iterations and PIs", [
            "Projects: Project, Role, Team. Exactly one project is the Program (the features); the others are Team.",
            "Iterations: Iteration-label, Start, End (the end is excluded). They are the short cycles inside a PI, for example 6 of "
            "3 weeks in a PI of 18 weeks, and the iterations of the charts. Column Example shows the form of each value.",
            "PIs: PI, Start, End, with names such as PI#01-2026 (data check DC18).",
        ]),
        (f'Button "{CONFIG_CAPTION}"', [
            "It reads Jira (read only) for the projects of the sheet Projects and writes the proposal sheets "
            f"{PROPOSAL_SHEETS}. It never changes the parameter sheets.",
            "Decide by hand: Order, IsStart, IsEnd, Active or Wait, Removed instead of Done, and the Type of a "
            "transition. Then copy into Workflows and Transitions, and the field ids into the ...Field rows of Config.",
        ]),
        (f'Button "{EXTRACT_CAPTION}"', [
            "It reads Jira in small slices: the bar moves, cell A3 says what happens, and Excel stays usable.",
            f'"{CANCEL_CAPTION}" stops it before the raw sheets are written: the previous results stay.',
            "At the end the sheets ItemFlow, DailyLoad, Metrics, ChartData and the charts are rebuilt to the size of "
            "the new data. Run it as often as needed.",
            "Cell B2 of the Dashboard: the team of the Flow Distribution chart.",
            'A status that starts with "Stopped" gives the error. The local log tells more (below).',
        ]),
        ("Results", [
            'Metrics: one row per iteration and team. A figure that cannot be computed reads "not computable", never 0.',
            "DataChecks: the 29 data checks DC01-DC29, with Items and OK, KO or not computable. To see the work items, "
            "filter the table ItemFlow on its column DCxx = 1.",
            "Do not edit the sheets FlowFormulas, ChartSpec and ChecksSpec: they describe the formulas, the charts and "
            "the checks.",
        ]),
        ("Local log", [
            r"%LOCALAPPDATA%\jira2dashboard\logs\jira2dashboard.log: one line per step. It holds no token, e-mail, "
            "address or Jira answer. LogLevel in Config: Steps (each step) or Errors. After an error the Dashboard names the log file and "
            "Notepad opens it, unless OpenLogOnError is No.",
        ]),
    )
    rows = [(HEADING, "jira2dashboard - user manual")]
    for title, lines in sections:
        rows.append((HEADING, title))
        rows.extend((TEXT, line) for line in lines)
    return rows
