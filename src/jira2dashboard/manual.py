"""The user manual of the delivered workbook: the first steps on the Dashboard and the sheet Manual.

Pure text. flow/sheets.py writes it at build time (`write_manual`). The button captions and the list of mandatory
parameters come from their sources, so the manual follows the code. The sheets are named by their real names
(`Config_from_user`): the suffix says where the data of a sheet comes from (sheet_names.py).
"""
from jira2dashboard import sheet_names as S
from jira2dashboard.flow.dashboard_spec import (CANCEL_CAPTION, COMPUTE_CAPTION, CONFIG_CAPTION, DRAW_CAPTION, EXTRACT_CAPTION, PROPOSE_CAPTION,
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
        f"4. Check the yellow cells it wrote in the sheets {CONFIG} and {ITERATIONS}, then type what is left. The statuses, their categories, the start and the end of "
        f"each flow and the moves are read from Jira by the workbook itself: type in {WORKFLOWS} only what Jira cannot say (a status that waits, a Done status that is "
        "no delivery, such as cancelled).",
        f'5. Click "{EXTRACT_CAPTION}" and wait for the bar to reach the end ("{CANCEL_CAPTION}" stops it), then "{COMPUTE_CAPTION}", then '
        f'"{DRAW_CAPTION}". After a change of {WORKFLOWS}, {ITERATIONS} or {PIS}, click only the last two: Jira is not read again.',
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
            "If a button says that the workbook is locked (J2D-SYS-003): click Enable Editing in the yellow bar at the top of "
            "Excel if there is one, or save a copy (File > Save As) and open that one, then click the button again.",
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
            f"{S.of('ErrorCodes')}. {S.of('FlowFormulas')}, {S.of('ChartSpec')}, {S.of('ChecksSpec')}, "
            f"{S.of('RawTypesSpec')} and {S.of('RovoPromptSpec')} are the plan the workbook builds from: do not edit them.",
        ]),
        (f"Sheet {CONFIG}", [
            "One row per parameter: Parameter, Value, Example, Mandatory, Description.",
            "Example shows the expected form of the value (it is never read: only Value counts).",
            f'Mandatory = True: a button stops while a value it needs is empty, and names it. "{EXTRACT_CAPTION}" needs BaseUrl, Email, '
            f'ApiToken, AsOf and BlocksLinkType (it stops before calling Jira), "{COMPUTE_CAPTION}" needs FlowLoadMode, PredictabilityMode, AsOf, '
            f'FeatureTypes and FillRateThreshold, "{CONFIG_CAPTION}" only BaseUrl, Email and ApiToken.',
            'Mandatory = False: an empty value is allowed. A default is used, nothing is added (no filter, no proxy), '
            'or the data checks that need the value show "not computable".',
            f"The mandatory parameters: {', '.join(mandatory)}.",
            "AsOf: the UTC date and time of the extraction, for example 2026-11-03 18:00. Open statuses run up to it.",
            "ProxyServer: empty = the proxy of the Windows settings, found as the browser finds it (automatic "
            "detection, proxy script or fixed proxy). DIRECT = no proxy. Or type host:port.",
            "ApiToken stays in clear in this sheet: empty it before you send the workbook to anyone.",
        ]),
        (f"Sheets {PROJECTS}, {ITERATIONS} and {PIS}", [
            "Projects: Project, Role, Team. Exactly one project is the Program (the features); the others are Team (none for an ART of Features only).",
            "Iterations: Iteration-label, Start, End (the end is excluded). They are the short cycles inside a PI, for example 6 of "
            "3 weeks in a PI of 18 weeks, and the iterations of the charts. Column Example shows the form of each value.",
            "PIs: PI, Start, End, with names such as PI#01-2026 (data check DC18).",
        ]),
        (f'Button "{CONFIG_CAPTION}"', [
            f"It reads Jira (read only) for the projects of the sheet {PROJECTS} and writes the proposal sheets "
            f"{PROPOSAL_SHEETS}. It never changes the parameter sheets.",
            f"The metrics read the statuses, the workflow of each issue type and its moves from these sheets (also read by the Extract button when they are missing). "
            f"Decide by hand only the exceptions, in {WORKFLOWS}: Wait (Jira has no such category), Removed instead of Done, a status that does not end the flow. "
            f"The field ids go into the ...Field rows of {CONFIG}.",
        ]),
        (f'Button "{PROPOSE_CAPTION}"', [
            f'It does what "{CONFIG_CAPTION}" does, then copies what fits of the proposal sheets into the sheets {CONFIG} and {ITERATIONS} '
            f"(never into {WORKFLOWS} and {TRANSITIONS}: the statuses, workflows and moves are read from Jira by the workbook): the ids "
            "of the fields that Jira names once (Sprint, Team, Story Points, Target start and end, PI, Business Value, work type "
            "category), the issue types of SubTaskTypes (the sub-tasks of Jira) and FeatureTypes (the types of the Program project at its "
            "highest hierarchy level), the types of ParentRequiredTypes and PointsRequiredTypes that no team project has taken out, and "
            f"BlocksLinkType (the link type that blocks, {S.of('JiraLinkTypes')}). It reads the sprints of the scrum boards "
            f"({S.of('JiraSprints')}) and proposes the iterations they make in {ITERATIONS}: the sprints of the teams that start and end "
            "on the same days are one iteration, and no day belongs to two iterations. PI dates have no API: type them.",
            "It writes only into EMPTY cells: what you typed is never changed. Every cell it writes is yellow, the first cell of a "
            "row has the note 'To confirm': check it, then clear the colour if you like. IsStart, Active or Wait and Removed instead "
            "of Done stay yours to type. In " + CONFIG + " a value that still is the delivered default (the same text as its Example) is "
            "replaced when Jira says otherwise, in yellow; a value you changed is never touched.",
            f"IsUserProjectAdmin in {CONFIG}: True (default) = you administer the Jira projects, the full workflows are read. "
            "False = only the statuses and their categories are read (the right to browse is enough), without the moves; the "
            "extraction does not ask for the workflows either, so declare the start statuses (IsStart) in " + S.of("Workflows") + ".",
        ]),
        (f'Button "{ROVO_CAPTION}"', [
            f"It writes the sheet {S.of('RovoPrompt')}: a prompt that asks Rovo for what the workbook cannot read in Jira by itself (the "
            "parameters of the flow metrics, the projects, the PIs, the statuses that wait or are no delivery). It names no parameter "
            "of the connection, no field id, no issue type, no sprint and no move: the workbook reads them. The lines are selected: Ctrl+C, "
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
        ("Three steps: which button after a change", [
            f'"{EXTRACT_CAPTION}" reads Jira, "{COMPUTE_CAPTION}" makes the tables and the figures, "{DRAW_CAPTION}" draws the charts. The last two '
            "never call Jira and take a few seconds: correct a parameter, then compute and draw again.",
            f'A change of the connection, the proxy, Filter, the projects listed in {PROJECTS}, a field id of {CONFIG} (the ...Field values, WorkTypeField) '
            f'or BlocksLinkType: "{EXTRACT_CAPTION}", then the two others.',
            f'A change of {WORKFLOWS} (a status that waits, a start status), {TRANSITIONS}, {ITERATIONS}, {PIS}, the Role or the Team of {PROJECTS}, '
            f'AsOf, FeatureTypes, FlowLoadMode, PredictabilityMode or FillRateThreshold: "{COMPUTE_CAPTION}", then "{DRAW_CAPTION}".',
            f"Each step ends with a status in cell {STATUS_CELL} (under the buttons) and its details in the rows under it, one per line (name in "
            'column A, value in column B). A status that starts with "Stopped" gives the error, with its code (J2D-...): read it to the support. '
            f'The sheet {S.of("ErrorCodes")} says what each code means and what to check. The local log tells more (below).',
        ]),
        (f'Button "{EXTRACT_CAPTION}"', [
            f"It reads Jira in small slices: the bar moves, cell {STATUS_CELL} says what happens, and Excel stays usable. At the end it says "
            f'Extracted, the number of issues, status intervals and changelog rows read, and that "{COMPUTE_CAPTION}" comes next.',
            f'"{CANCEL_CAPTION}" stops it before the raw sheets are written: the previous results stay. Once they are written, the results of the '
            "previous computation are deleted: never old figures next to new data.",
            f"Besides the raw sheets it keeps what Jira said about the statuses ({S.of('JiraStatusCatalog')}, {S.of('JiraWorkflows')}): "
            f'"{COMPUTE_CAPTION}" reads them, so that a change of {WORKFLOWS} needs no new extraction.',
        ]),
        (f'Button "{COMPUTE_CAPTION}"', [
            f"It makes the rules of the statuses again (what Jira says, under the exceptions of {WORKFLOWS}: the sheet "
            f"{S.of('StatusRules')}), then the tables and formulas of the figures. It never calls Jira. The charts are deleted (they read the "
            f'tables made again): click "{DRAW_CAPTION}".',
            "When something keeps a figure empty, the status adds a field: No start status (a workflow with tickets and no IsStart status: no Flow "
            f"Time, Flow Efficiency or Flow Load; set IsStart in {WORKFLOWS}), No Wait status (no Flow Efficiency), Not Features (the tickets of the "
            "Program that are not of a type of "
            "FeatureTypes are in the team ART (other types) of ItemFlow, left out of the ART metrics; if the ART has no Feature, FeatureTypes is "
            "not the name of Jira), Done outside the iterations (add iterations), Without a team, "
            f"Statuses not in Workflows (no Jira source knows them either: give them a Category in {WORKFLOWS}), Started in the second of the "
            "backlog entry, AsOf before the data, VelocityCheck gap, Reconciliation (controls to review) and Cells in error (Table.Column x n: a "
            "formula that fails on your data).",
            f"The sheets {ITEMFLOW}, {S.of('DailyLoad')}, {METRICS}, {S.of('PIMetrics')} and the others are rebuilt to the size of the data. Run it "
            "as often as needed.",
            f"Feature Predictability, per PI of the sheet {PIS}: the Features released in the PI over the Features engaged in it at "
            "its start, read in the history of the PI field (the start of the first iteration of the PI): a Feature deferred later "
            "(its PI field emptied or moved) stays planned, a cancelled one stays planned, not done, and a Feature added after the "
            "start is outside the ratio. PredictabilityMode B counts the added Features too (the ratio can pass 100 %). It needs "
            f"FeatureTypes, the PI field (PIField), the sheet {PIS} and iterations that cover the PI. A PI that Features carry but "
            f"the sheet {PIS} lacks comes after them, not computable (it has no dates): add it to the sheet {PIS}. Flow "
            "Predictability, the SAFe measure on the business value of the PI objectives, stays not computable: Jira does not hold "
            "the PI objectives.",
            f"The sheet {S.of('IterationFeatures')} gives, per iteration, PI and team (the ART, then the teams of the Features: the "
            "Team field), the Features engaged at its start, added, deferred, released and carried over; the PI figures are sums of "
            f"it. Its charts are in the sheet {S.of('IterationCharts')}.",
        ]),
        (f'Button "{DRAW_CAPTION}"', [
            f'It draws the charts from the tables of "{COMPUTE_CAPTION}", in a few seconds, without calling Jira: the sheet {CHARTS}, right of '
            f"the Dashboard (its cell B2: the team of the Flow Distribution chart), and the sheet {S.of('IterationCharts')}. Its status repeats "
            "the fields of what to look at.",
        ]),
        ("Results", [
            f'{METRICS}: one row per iteration and team. A figure that cannot be computed reads "not computable", never 0.',
            f"{S.of('PIMetrics')}: one row per PI of the sheet {PIS}, with its dates, the Features engaged in it at its start, those "
            f"released in it and Feature Predictability; then one row per PI of a Feature missing from the sheet {PIS}, without dates.",
            f"{S.of('PIStatusTime')}: where the cycle time of the Features goes. For each PI of the sheet {PIS} and each Active or Wait status of the "
            "Feature workflow: the Features released in the PI, the average days each spent in the status, and its share of the total (the "
            "statuses of a PI add up to 100 %). The chart Cycle time by status stacks these shares, one bar per PI: the biggest block "
            "is the status that takes most of the cycle time.",
            f"{S.of('FlowTimeCheck')}: why the Flow Time of each iteration and team is what it is (the rows of Metrics): Done (counted again from "
            "the status intervals), the done items left out of Flow Time, each under one reason (reopened, no start status in the rules: set IsStart "
            "in Workflows, no start after the backlog), Counted (the Flow Time count), Gap (0) and Reason: why the point is not computable, in clear.",
            f"{S.of('FlowEfficiencyCheck')}: the same for Flow Efficiency: the counted items, those with time in a status no rule knows, the active "
            "and wait days summed again from the status intervals, the Wait statuses of the rules, the efficiency recounted, Gap (0) and Reason "
            "(no Wait status: declare one in Workflows; nothing counted; unclassified time; no active or wait time).",
            f"{S.of('FlowLoadCheck')}: the same for Flow Load: the days of the iteration, those measured (not after AsOf), the start statuses of "
            "the workflow of the team, the average load and the load at the end counted again, Gap (0) and Reason (no start status: declare one "
            "in Workflows; every day after AsOf).",
            f"{S.of('FlowDistCheck')}: the done items counted again per work type (the case of the work type does not matter), Unmapped, "
            "Gap (0), Reason (no item done) and a Note when every done item is Unmapped (WorkTypeField empty, or no known work type).",
            f"{S.of('VelocityCheck')}: the Velocity counted again from the status changes of the changelog, one row per iteration and team "
            "(the rows of Metrics). Velocity = the items done in the iteration (the chart); VelocityFeatures = of which the types of "
            "FeatureTypes (the team ART holds every type of the Program project); CreatedDone = done because they were CREATED in the end "
            "status (the changelog holds no change for them); ChangedToEnd = the changes into an end status (IsEnd) in the iteration, "
            "one per change (an item closed twice counts twice), ChangedToEndFeatures = of which Features (what the JQL status CHANGED TO "
            "... DURING counts in Jira); ReopenedLater = those changes that the item left again. Gap = Velocity - CreatedDone - "
            "(ChangedToEnd - ReopenedLater) is 0 when the table explains the Velocity. Reason says why the Velocity differs from Jira's "
            "count of the changes into an end status (created done, closed then reopened, other types of the Program), or that nothing was done.",
            f"{S.of('MetricItems')}: the tickets BEHIND the figures, three lists side by side (formulas that follow the data). Left: every ticket done in "
            "an iteration, by iteration, team, date and key, with its type, whether Flow Time counts it (FlowTimeCounted 1 or 0) and why not, "
            "its active, wait and unclassified days (Flow Efficiency) and its work-type bucket (Flow Distribution) "
            "(Reopened; No start status in the rules: set IsStart in Workflows; No start after the backlog): the rows of an iteration and team are its "
            "Velocity, so filter the list to see the keys of a chart point. Middle: the "
            "done tickets that no figure counts, with the reason (done outside the iterations, or left out because they are not Features: the team "
            f"ART (other types)): the missing ones. Right: the tickets done within 24 hours of an iteration boundary (cut-off). {ITEMFLOW} has the columns "
            "behind them: DoneIteration, DonePI, Account (every ticket is in exactly one account), WhyNotInFlowTime, BoundaryHours.",
            f"{S.of('Reconciliation')}: the controls an accountant would run, each with Expected, Found and a Result (OK, or REVIEW: look at it): the tickets "
            "read = the tickets of ItemFlow, every ticket in one account, the Velocity and the Flow Time recounted from the lists, no ticket twice, the tickets "
            "near a boundary (INFO: a point of attention, not counted on the Dashboard), the statuses of the counted tickets all classified (the statuses that "
            "no rule classifies are listed by name right of the census), FeatureTypes matching a ticket, consistent status histories, the done tickets found again in the "
            "changelog. Next to them the census of the ticket types (how many, in the ART, left out, done, done in an iteration) and, below, the parameters "
            "in force (FeatureTypes, the statuses that start and end the flow, the dates of the iterations): compare them with Jira. The status of the "
            "Dashboard says how many controls are to review.",
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
