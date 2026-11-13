"""The names of the worksheets of the workbook: a suffix says where the data of a sheet comes from.

| Suffix        | What the sheet is                                                                           |
|---------------|---------------------------------------------------------------------------------------------|
| `_from_user`  | typed by the person: the parameters (Config, Projects, Iterations, PIs, Workflows...)        |
| `_from_Jira`  | filled with what the Jira API answered: the tickets, their status history, the links, the   |
|               | proposals of "Read the Jira configuration"                                                  |
| `_from_excel` | calculated or drawn by Excel formulas and macros (the flow metrics, the charts), or written  |
|               | by the workbook itself (the dashboard, the manual, the plan and the texts built with it)     |

The code keeps the LOGICAL names (`Config`, `Issues`, `Metrics`...): the tables of Excel keep them too (`Issues[Key]`).
`of(name)` gives the real name of the sheet, where the code meets Excel (create a sheet, find it, write a formula that
points at it, show its name to the person). VBA: `modText.SheetName`, the same three lists (`test_sheet_names.py` pins it).
"""
FROM_USER = "_from_user"
FROM_JIRA = "_from_Jira"
FROM_EXCEL = "_from_excel"

SUFFIX = {"user": FROM_USER, "jira": FROM_JIRA, "excel": FROM_EXCEL}

USER = ("Config", "Projects", "Iterations", "PIs", "Workflows", "Transitions", "FieldMapping")
JIRA = ("Issues", "StatusIntervals", "IssueLinks", "JiraIssueTypes", "JiraStatuses", "JiraWorkflows", "JiraTransitions",
        "JiraFields", "JiraProjectWorkflows", "JiraProjectFields", "JiraFieldUsage", "JiraSprints")
EXCEL = ("Dashboard", "Charts", "Manual", "ErrorCodes", "Steps", "ItemFlow", "DailyLoad", "Metrics", "PIMetrics",
         "DataChecks", "FillRates", "Controls", "ChartData", "FlowFormulas", "ChartSpec", "ChecksSpec", "RovoPromptSpec",
         "RovoPrompt", "Run", "Errors")                  # Run and Errors: the sheets of the Python path only


def source(name: str) -> str | None:
    """"user", "jira" or "excel" for a logical name; None for a name the workbook does not know."""
    logical_name = logical(name)
    for kind, names in (("user", USER), ("jira", JIRA), ("excel", EXCEL)):
        if logical_name in names:
            return kind
    return None


def logical(name: str) -> str:
    """`Config_from_user` -> `Config`; any other text is returned as it is."""
    for suffix in SUFFIX.values():
        if name.endswith(suffix) and name[:-len(suffix)] in USER + JIRA + EXCEL:
            return name[:-len(suffix)]
    return name


def of(name: str) -> str:
    """The real name of the sheet: `Config` -> `Config_from_user`. A real name, or a name the workbook does not know
    (a scratch sheet of a test), comes back unchanged."""
    kind = source(name)
    return name if kind is None or logical(name) != name else name + SUFFIX[kind]
