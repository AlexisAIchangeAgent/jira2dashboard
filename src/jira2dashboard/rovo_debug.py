"""The debug prompt for Rovo: the files that let Rovo compare the figures of the workbook with Jira (pure text; AL, 2026-11-26).

flow/sheets.py writes `spec_rows` into the sheet RovoDebugSpec at build time. The button "Debug prompt for Rovo" (modRovoDebug.WriteRovoDebug)
reads that sheet and the tables of 2. Compute the metrics, then writes Markdown files 01.md, 02.md... into the folder rovo-debug next to the
workbook: file 1 holds how to answer, the causes, the context and the controls of the Reconciliation; then the iterations (the last
MAX_ITERATIONS that started), each with its keys and two JQL that return only the differences, as many per file as fit in MAX_CHARACTERS.

Why so (research of 2026-11-26): Rovo does not count reliably (ROVO-535 grouped counts, ROVO-287 wrong over about a hundred tickets, ROVO-249 a
CSV of 26 rows read as 16) and writes the history JQL badly. So it runs the JQL written here as it is, lists keys with one cause of a closed list,
and never gives a total. A .md file (not .csv: the table tools lose rows), in Rovo Chat (Studio's test mode has no upload). No size limit is
documented: 12 000 characters a file is a guess to calibrate at the client.

The VBA replaces each token {name} of a text by its value; a token that holds several lines (a list) gives them all. A text never holds a
value of the connection (address, e-mail, token, proxy), a summary of a ticket or a person's name. Rovo knows no data-check code: no DCnn here.
"""
from jira2dashboard import sheet_names as S
from jira2dashboard.flow import reconciliation as R
from jira2dashboard.flow.formulas import metricitems_column
from jira2dashboard.flow.trace import CUTOFF_HOURS

SPEC_SHEET = "RovoDebugSpec"
SPEC_HEADER = ("Part", "Name", "Text", "Hint")
FOLDER = "rovo-debug"
MAX_CHARACTERS = 12_000
MAX_ITERATIONS = 18
LIMIT, COLUMN, HEAD, NEXT, SUMMARY, ITERATION, TEAM, JQL, WORD = "limit", "column", "head", "next", "summary", "iteration", "team", "jql", "word"
JIRA, PARTIAL, WORKBOOK = "jira", "partial", "workbook"      # the group of a control: checkable with a JQL, partly, by the workbook alone
GROUPS = (JIRA, PARTIAL, WORKBOOK)

RECONCILIATION, STATUS_RULES, METRIC_ITEMS = S.of("Reconciliation"), S.of("StatusRules"), S.of("MetricItems")

# The columns of the sheet MetricItems that the VBA reads, by list (formulas.METRICITEMS_BLOCKS: 0 Done, 1 left out, 2 cut-off)
COLUMNS = (("Done", 0, ("Iteration", "Team", "Key", "FlowTimeCounted", "WhyNotInFlowTime")),
           ("LeftOut", 1, ("Key", "Team", "Iteration", "Account")),
           ("Cutoff", 2, ("Key", "Team", "Iteration", "Hours")))

HEAD_LINE = "# jira2dashboard - debug prompt for Rovo, file {file} of {files}"
NEXT_LINE = "Read file 1 first: it says how to answer and gives the context and the controls of the workbook."

INTRO = (
    "You help me find why figures of my Excel workbook jira2dashboard differ from Jira. The workbook read our Jira Cloud site, then "
    "computed the six SAFe flow metrics; its own controls, at the end of this file, say which figures to review.",
    "These files hold no address of our site, no e-mail of the account and no token; the Filter of the workbook is quoted as typed. Every date "
    "and time is UTC; the end of an iteration is excluded.",
    "",
    "## How to answer",
    "1. First, for each file: its number, the iterations it holds, and how many keys you read in each list. When a number differs from the "
    "count written in the file, say so: a part of the file was lost.",
    "2. Run each JQL exactly as it is written, in the Jira search. Never write or change a JQL yourself.",
    "3. For each JQL, list the keys it returns, each with ONE cause of the list below and what you see in Jira (status, created, "
    "resolved). When it returns nothing, write \"no difference\".",
    "4. Never give a total that you computed: give a number of results only when the search itself reports it.",
    "5. End with one table: Iteration, Team, Key, Cause, What to check. Then the causes, the most frequent first.",
    "6. Never invent a key, a date or a status. Never compute a metric.",
    "",
    "## Causes (one per key)",
    f"- Time zone: done within {CUTOFF_HOURS} hours of an iteration boundary; the workbook reads UTC, a JQL the time zone of your Jira profile.",
    "- Created done: created directly in an end status, so its history holds no change into it.",
    "- Reopened: done, reopened, then done again in another iteration; the workbook counts the last time.",
    "- Not a Feature: an issue type that the workbook does not count as a Feature (FeatureTypes), or a sub-task.",
    "- Status rules: an end status in Jira that the rules of the workbook do not count as one, or the reverse.",
    "- Filter or scope: the Filter or the projects of the workbook leave it in or out.",
    "- Rights: the account of the workbook may not see it.",
    "- AsOf: done after the AsOf date of the workbook.",
    "- Other: say what you see.",
    "",
)
CONTEXT = (
    "## Context",
    "- Workbook: {version}. AsOf (UTC): {as_of}.",
    "- Filter (in every JQL): {filter}",
    "- FeatureTypes (the issue types of the ART): {feature_types}",
    "- Projects (key: role, team):",
    "{projects}",
    f"- Status rules of the workbook ({STATUS_RULES}) per workflow kind (Feature = the Program project, Team = the team projects):",
    "{statuses}",
    "- Iterations: the controls count all {iterations_total}; the next files list the last {iterations_listed} that started, from "
    "{window_start} to {window_end}:",
    "{iteration_index}",
    "- JQL E, every ticket the workbook read: {jql_extraction}",
    "- JQL F, the Features of the Program: {jql_features}",
    "- What the last computation says to look at:",
    "{health}",
    f"- Parameters in force (below the table of {RECONCILIATION}):",
    "{parameters}",
    "",
    f"## The {len(R.CONTROLS)} controls of {RECONCILIATION}",
    "OK: nothing to explain. REVIEW: explain it. INFO: a point of attention.",
    "{controls}",
)
ITERATION_LINES = (
    "",
    "## Iteration {iteration}: from {start} to {end} UTC",
    "Left out of the metrics though done in it (key: team, reason): {left_out}",
    f"Done within {CUTOFF_HOURS} hours of a boundary (key: team, hours): " + "{cutoff}. Your Jira time zone may put them in the next or the "
    "previous iteration.",
    "{teams}",
)
TEAM_LINES = (
    "### {iteration}, team {team}",
    "- Velocity {velocity}{velocity_reason}. Flow Time counts {flow_time_count} of them.",
    "- Done in the workbook ({done_count} keys): {done_keys}",
    "- Left out of Flow Time (key: reason): {not_in_flow_time}",
    "- JQL A, done in Jira in this iteration and missing from the workbook: {jql_missing}",
    "- JQL B, done in the workbook and not done in Jira in this iteration: {jql_extra}",
)
# The JQL, filled by the VBA: {program} / {projects} / {end_statuses} are quoted lists, {keys} keys, {start} and {last} "yyyy-mm-dd hh:mm" (the
# last minute of the iteration: DURING includes both ends). A clause on a list that is empty is left out ({not_keys}, {filter}).
JQLS = (
    ("scope_art", "project in ({program}) AND issuetype in ({feature_types})"),
    ("scope_team", "project in ({projects})"),
    ("filter", " AND ({filter})"),
    ("missing", '{scope}{filter_jql} AND status in ({end_statuses}) AND status CHANGED TO ({end_statuses}) DURING ("{start}", "{last}"){not_keys}'),
    ("not_keys", " AND key not in ({keys})"),
    ("extra", 'key in ({keys}) AND NOT (status CHANGED TO ({end_statuses}) DURING ("{start}", "{last}"))'),
)
WORDS = (
    ("none", "none"),
    ("no_keys", "none: the workbook counts no ticket here"),
    ("no_end", f"none: the rules of the workbook ({STATUS_RULES}) have no end status for this team"),
    ("no_scope", f"none: the workbook has no project for this team, or no FeatureTypes ({S.of('Projects')}, {S.of('Config')})"),
)

# One text per control of the Reconciliation (reconciliation.CONTROLS), keyed by its name: (group, title for Rovo, what to do when it is not OK).
_WORKBOOK_PATHS = "Two paths of the workbook give the same count: a difference is in the formulas of the workbook, not in Jira."
CONTROL_TEXTS = {
    R.POPULATION: (JIRA, "Tickets read from Jira = tickets of the table ItemFlow",
                   "Run JQL E and give the number of results the search reports: a number other than Expected means the extraction missed "
                   "tickets (rights, Filter); a Found other than Expected is in the workbook."),
    R.ACCOUNTS_ADD_UP: (WORKBOOK, "Every ticket sits in one account (done in an iteration, done outside them, not a Feature, removed, open)",
                        _WORKBOOK_PATHS),
    R.VELOCITY_TOTAL: (JIRA, R.VELOCITY_TOTAL,
                       "JQL A and B of each iteration and team, in the next files, give the tickets that differ from Jira."),
    R.VELOCITY_ROWS: (JIRA, R.VELOCITY_ROWS,
                      "JQL A and B of each iteration and team, in the next files, give the tickets that differ from Jira."),
    R.FLOW_TIME_COUNT: (WORKBOOK, R.FLOW_TIME_COUNT, "The line \"Left out of Flow Time\" of each team says why a done ticket is not counted."),
    R.FLOW_TIME_DAYS: (WORKBOOK, R.FLOW_TIME_DAYS, _WORKBOOK_PATHS),
    R.FLOW_TIME_GAP: (WORKBOOK, R.FLOW_TIME_GAP, _WORKBOOK_PATHS),
    R.FLOW_EFFICIENCY_GAP: (WORKBOOK, R.FLOW_EFFICIENCY_GAP, _WORKBOOK_PATHS),
    R.FLOW_LOAD_GAP: (WORKBOOK, R.FLOW_LOAD_GAP, _WORKBOOK_PATHS),
    R.FLOW_DISTRIBUTION_GAP: (WORKBOOK, R.FLOW_DISTRIBUTION_GAP, _WORKBOOK_PATHS),
    R.FLOW_LOAD_LIST: (WORKBOOK, R.FLOW_LOAD_LIST, f"The in-process list of {METRIC_ITEMS} holds the tickets: " + _WORKBOOK_PATHS),
    R.CARRY_OVER: (WORKBOOK, R.CARRY_OVER, _WORKBOOK_PATHS),
    R.UNIQUE: (WORKBOOK, R.UNIQUE, "A key listed twice in the next files is the ticket counted twice."),
    R.CUTOFF: (PARTIAL, R.CUTOFF,
               "Each iteration lists them. JQL reads dates in the time zone of your Jira profile, the workbook in UTC: such a ticket may fall in "
               "the next or the previous iteration in Jira (cause Time zone)."),
    R.CLASSIFIED: (PARTIAL, "Statuses of the counted tickets all classified by the rules of the workbook",
                   "Say whether each status of the rules listed as Unclassified still exists in the Jira workflows of the projects (a retired "
                   "workflow leaves old statuses in the histories)."),
    R.FEATURE_TYPES: (JIRA, R.FEATURE_TYPES,
                      "Run JQL F: when it returns nothing, give the issue types of the Program project as Jira names them (FeatureTypes must "
                      "name one)."),
    R.START_STATUS: (PARTIAL, R.START_STATUS,
                     "Give, for each Jira workflow of the projects, the first in-progress status after the creation status: the workbook must "
                     "declare it IsStart."),
    R.WAIT_STATUS: (PARTIAL, R.WAIT_STATUS,
                    "Give the statuses of the Jira workflows where the work waits (ready for, to test, blocked): Jira has no Wait category, "
                    "the workbook must declare them."),
    R.HISTORIES: (WORKBOOK, "Status histories consistent (each change starts from the previous status, none before the creation)",
                  "The histories come from the extraction; Rovo cannot read them. The table ItemFlow of the workbook names the tickets."),
    R.EXISTENCE: (WORKBOOK, R.EXISTENCE, _WORKBOOK_PATHS),
}
assert tuple(CONTROL_TEXTS) == tuple(name for name, _ in R.CONTROLS), "one text per control, in the order of the catalogue"


def spec_rows() -> list[tuple[str, str, str, str]]:
    """The rows of the sheet RovoDebugSpec: (Part, Name, Text, Hint), in the order of the files."""
    rows = [(LIMIT, "MaxCharacters", str(MAX_CHARACTERS), ""), (LIMIT, "MaxIterations", str(MAX_ITERATIONS), "")]
    rows += [(COLUMN, f"{block}.{name}", str(metricitems_column(number, name)), "") for block, number, names in COLUMNS for name in names]
    rows += [(HEAD, "", HEAD_LINE, ""), (NEXT, "", NEXT_LINE, "")]
    rows += [(SUMMARY, "", line, "") for line in INTRO + CONTEXT]
    rows += [(group, name, title, hint) for name, (group, title, hint) in CONTROL_TEXTS.items()]
    rows += [(ITERATION, "", line, "") for line in ITERATION_LINES]
    rows += [(TEAM, "", line, "") for line in TEAM_LINES]
    rows += [(JQL, name, text, "") for name, text in JQLS]
    rows += [(WORD, name, text, "") for name, text in WORDS]
    return rows
