"""The Excel formulas of the flow sheets, as text (English function names, structured references).

Pure: nothing here touches Excel. blueprint.py writes them into the sheet FlowFormulas and modFlowBuild.bas puts them in the tables. The formulas follow docs section 2 and are what the
differential test compares, cell by cell, with the Python reference (itemflow.py, load.py, metrics.py).

Rules kept here:
* no status name and no mapping in a formula: the status flags and categories are read from the sheet Workflows,
  the teams and roles from Projects, the iterations from Iterations, the mode and the as-of date from Config;
* every comparison with a date adds or removes half a second: dates are whole seconds, and Excel turns a bound
  into text with 15 digits, so a strict `>` on the exact serial could let an equal date through;
* a blank result is the empty text "", which COUNTIFS / MINIFS / SUMIFS ignore in numeric criteria.
"""
from jira2dashboard.flow import check_formulas as C
from jira2dashboard.flow import reconciliation as R
from jira2dashboard.flow import trace as T
from jira2dashboard.flow.model import ART, OTHER_TYPES
from jira2dashboard.flow.refs import (HALF, IF_, IS, PIS_LAST, SI, TRANSITIONS_LAST, bounded, last_row_cell,
                                      row_index)

FAR = "2958465"                      # 9999-12-31: "no end"
NOT_COMPUTABLE = "not computable"
CATEGORIES_IN_PROCESS = ("Active", "Wait", "Unclassified")
BUCKETS = ("Feature", "Enabler", "Defect", "Autre", "Unmapped")
WORK_TYPES = BUCKETS[:-1]

ISSUES = ("Key", "Id", "Project", "Type", "WorkType", "CreatedUtc", "Status", "Parent", "StatusCategory", "ResolvedUtc",
          "FixVersions", "StoryPoints", "PI", "BV", "TeamValue", "SprintState", "SprintStart", "SprintEnd", "SprintCount",
          "TargetStart", "TargetEnd", "KeyChanges")
# The sheet Changelog (raw, from Jira): one row per changelog item, every field. No author. A text (FromString, ToString) is cut at
# CHANGELOG_TEXT_MAX characters; the person fields keep no value (a name is personal data). Same columns as modFlow.ChangelogHeader.
CHANGELOG_RAW = ("Key", "IssueId", "ChangeId", "CreatedUtc", "CreatedText", "FieldId", "Field", "FieldType", "From", "FromString",
                 "To", "ToString")
CHANGELOG_TEXT_MAX = 255
CHANGELOG_PERSON_FIELDS = ("assignee", "reporter", "creator")
INTERVALS_RAW = ("Key", "Workflow", "Seq", "Status", "Category", "EnteredUtc", "LeftUtc", "DurationDays", "IsOpen",
                 "ChangeId", "Check")

_row_index = row_index


# --- StatusIntervals: helper columns after the raw ones ------------------------------------------------

INTERVAL_HELPERS = (
    ("IsStart", "=SUMIFS(Workflows_from_user!$D:$D,Workflows_from_user!$A:$A,[@Workflow],Workflows_from_user!$B:$B,[@Status])"),
    ("IsEnd", "=SUMIFS(Workflows_from_user!$E:$E,Workflows_from_user!$A:$A,[@Workflow],Workflows_from_user!$B:$B,[@Status])"),
    ("PrevCategory", f'=IF([@Seq]=1,"",INDEX({SI}[Category],{_row_index(SI)}-1))'),
    ("Team", f"=INDEX({IF_}[Team],MATCH([@Key],{IF_}[Key],0))"),
    ("Bucket", f"=INDEX({IF_}[WorkTypeBucket],MATCH([@Key],{IF_}[Key],0))"),
    ("FirstStart", f"=INDEX({IF_}[FirstStartUtc],MATCH([@Key],{IF_}[Key],0))"),
    ("LeftEff", f'=IF([@LeftUtc]="",{FAR},[@LeftUtc])'),
    ("Counted", "=IF(OR(" + ",".join(f'[@Category]="{c}"' for c in CATEGORIES_IN_PROCESS) + "),1,0)"),
    ("PrevStatus", f'=IF([@Seq]=1,"",INDEX({SI}[Status],{_row_index(SI)}-1))'),
    # DC28: the move from the previous status is not in Transitions (same workflow) and the status is not free (IsFree)
    ("BadMove", f'=IF([@Seq]=1,0,IF(OR(COUNTIFS({bounded("Transitions", "A", TRANSITIONS_LAST)},[@Workflow],'
                f'{bounded("Transitions", "B", TRANSITIONS_LAST)},[@PrevStatus],'
                f'{bounded("Transitions", "C", TRANSITIONS_LAST)},[@Status])>0,SUMIFS(Workflows_from_user!$F:$F,Workflows_from_user!$A:$A,[@Workflow],'
                'Workflows_from_user!$B:$B,[@Status])>0),0,1))'),
    # PIStatusTime: the time of a Feature in each Active / Wait status, between its StartUtc (included) and its DoneUtc (excluded), the
    # window of ItemFlow.ActiveDays / WaitDays; 0 for an interval of an item that is not done, has no valid start or is excluded from flow time
    ("ItemRow", f"=MATCH([@Key],{IF_}[Key],0)"),                       # one lookup per interval, the three columns below read it
    ("ItemStart", f"=INDEX({IF_}[StartUtc],[@ItemRow])"),
    ("ItemDone", f"=INDEX({IF_}[DoneUtc],[@ItemRow])"),
    ("ItemExcluded", f"=INDEX({IF_}[Excluded],[@ItemRow])"),
    ("ItemIsFeature", f"=INDEX({IF_}[IsFeature],[@ItemRow])"),                 # VelocityCheck: the Feature types of Config.FeatureTypes
    ("InCycle", f'=IF(AND(OR([@Category]="Active",[@Category]="Wait"),[@ItemStart]<>"",[@ItemDone]<>"",[@ItemExcluded]=0),'
                f"IF(AND([@EnteredUtc]>=[@ItemStart]-{HALF},[@EnteredUtc]<[@ItemDone]-{HALF}),1,0),0)"),
)


# --- ItemFlow (docs 2.1) ----------------------------------------------------------------------------------

def _starts_after_backlog(function: str) -> str:
    return (f"{function}({SI}[EnteredUtc],{SI}[Key],[@Key],{SI}[IsStart],1,"
            f'{SI}[EnteredUtc],">"&([@LastBacklogUtc]+{HALF}))')


def _days(category: str) -> str:
    return (f'=IF(OR([@StartUtc]="",[@DoneUtc]=""),"",SUMIFS({SI}[DurationDays],{SI}[Key],[@Key],'
            f'{SI}[Category],"{category}",{SI}[EnteredUtc],">="&([@StartUtc]-{HALF}),'
            f'{SI}[EnteredUtc],"<"&([@DoneUtc]-{HALF})))')


ITEMFLOW_COLUMNS = (
    ("Key", f"=INDEX({IS}[Key],{_row_index(IF_)})"),
    ("Project", f"=INDEX({IS}[Project],{_row_index(IF_)})"),
    ("Type", f"=INDEX({IS}[Type],{_row_index(IF_)})"),
    ("WorkType", f'=IF(INDEX({IS}[WorkType],{_row_index(IF_)})="","",INDEX({IS}[WorkType],{_row_index(IF_)}))'),
    ("Role", '=IFERROR(INDEX(Projects_from_user!$B:$B,MATCH([@Project],Projects_from_user!$A:$A,0)),"")'),
    ("Workflow", '=IF([@Role]="Program","Feature","Team")'),
    # the team of a ticket of the Program: ART for a Feature (type in Config.FeatureTypes), "ART (other types)" for the rest (sub-tasks...: no metric of
    # the ART counts them); of a Team project: its column Team, or the project key when that cell is empty (AL, 2026-11-18)
    ("Team", f'=IF([@Role]="Program",IF({C.in_list("[@Type]", "FeatureTypes")},"{ART}","{OTHER_TYPES}"),'
             'IFERROR(IF(INDEX(Projects_from_user!$C:$C,MATCH([@Project],Projects_from_user!$A:$A,0))="",[@Project],'
             'INDEX(Projects_from_user!$C:$C,MATCH([@Project],Projects_from_user!$A:$A,0))),""))'),
    ("CreatedUtc", f"=INDEX({IS}[CreatedUtc],{_row_index(IF_)})"),
    ("LastBacklogUtc", f'=MAXIFS({SI}[EnteredUtc],{SI}[Key],[@Key],{SI}[Category],"Backlog")'),
    ("FirstStartUtc", f'=IF(COUNTIFS({SI}[Key],[@Key],{SI}[IsStart],1)=0,"",'
                      f"MINIFS({SI}[EnteredUtc],{SI}[Key],[@Key],{SI}[IsStart],1))"),
    ("StartUtc", f'=IF(COUNTIFS({SI}[Key],[@Key],{SI}[IsStart],1,{SI}[EnteredUtc],">"&([@LastBacklogUtc]+{HALF}))=0,"",'
                 f"{_starts_after_backlog('MINIFS')})"),
    ("BacklogReturns", f'=IF([@FirstStartUtc]="",0,COUNTIFS({SI}[Key],[@Key],{SI}[Category],"Backlog",'
                       f'{SI}[EnteredUtc],">"&([@FirstStartUtc]+{HALF})))'),
    ("DoneUtc", f'=IF(SUMIFS({SI}[IsEnd],{SI}[Key],[@Key],{SI}[IsOpen],TRUE)>0,'
                f'MAXIFS({SI}[EnteredUtc],{SI}[Key],[@Key],{SI}[IsEnd],1),"")'),
    ("RemovedUtc", f'=IF(COUNTIFS({SI}[Key],[@Key],{SI}[IsOpen],TRUE,{SI}[Category],"Removed")>0,'
                   f'MAXIFS({SI}[EnteredUtc],{SI}[Key],[@Key],{SI}[Category],"Removed"),"")'),
    ("FlowTimeDays", '=IF(OR([@StartUtc]="",[@DoneUtc]=""),"",[@DoneUtc]-[@StartUtc])'),
    ("LeadTimeDays", '=IF([@DoneUtc]="","",[@DoneUtc]-[@CreatedUtc])'),
    ("ActiveDays", _days("Active")),
    ("WaitDays", _days("Wait")),
    ("UnclassifiedDays", _days("Unclassified")),
    ("Reopened", f'=COUNTIFS({SI}[Key],[@Key],{SI}[PrevCategory],"Done",{SI}[Category],"<>Done")'),
    ("SkippedStart", '=IF(AND([@DoneUtc]<>"",[@StartUtc]=""),"Y","N")'),
    ("Excluded", '=IF(OR([@SkippedStart]="Y",[@Reopened]>0),1,0)'),
    ("WorkTypeBucket", "=IF(OR(" + ",".join(f'[@WorkType]="{w}"' for w in WORK_TYPES) + '),[@WorkType],"Unmapped")'),
    ("EndUtc", f'=IF([@DoneUtc]<>"",[@DoneUtc],IF([@RemovedUtc]<>"",[@RemovedUtc],{FAR}))'),
    # the trace of a ticket (flow/trace.py): the iteration and the PI where it was done, its account, why Flow Time leaves it out, its distance to a boundary.
    # The iterations are those of the table Metrics (its key columns are values: no circular reference), the PIs those of PIMetrics.
    ("DoneIterationRow", f'=IF([@DoneUtc]="","",IFERROR(MATCH(1,(Metrics[IterationStart]-{HALF}<=[@DoneUtc])*([@DoneUtc]<Metrics[IterationEnd]-{HALF}),0),""))'),
    ("DoneIteration", '=IF([@DoneIterationRow]="","",INDEX(Metrics[Iteration-label],[@DoneIterationRow]))'),
    ("DoneIterationStart", '=IF([@DoneIterationRow]="","",INDEX(Metrics[IterationStart],[@DoneIterationRow]))'),
    ("DoneIterationEnd", '=IF([@DoneIterationRow]="","",INDEX(Metrics[IterationEnd],[@DoneIterationRow]))'),
    ("DonePI", f'=IF([@DoneUtc]="","",IFERROR(INDEX(PIMetrics[PI],MATCH(1,IFERROR((PIMetrics[Start]-{HALF}<=[@DoneUtc])*([@DoneUtc]<PIMetrics[End]-{HALF}),0),0)),""))'),
    ("Account", f'=IF([@Team]="{OTHER_TYPES}","{T.ACCOUNT_OTHER}",IF([@DoneUtc]<>"",IF([@DoneIteration]<>"","{T.ACCOUNT_COUNTED}","{T.ACCOUNT_OUTSIDE}"),'
                f'IF([@RemovedUtc]<>"","{T.ACCOUNT_REMOVED}","{T.ACCOUNT_OPEN}")))'),
    ("WhyNotInFlowTime", f'=IF([@Excluded]=0,"",TEXTJOIN(", ",TRUE,IF([@Reopened]>0,"{T.REOPENED}",""),IF([@SkippedStart]="Y","{T.NO_START}","")))'),
    ("BoundaryHours", '=IF([@DoneUtc]="","",MIN(MIN(ABS(Metrics[IterationStart]-[@DoneUtc])),MIN(ABS(Metrics[IterationEnd]-[@DoneUtc])))*24)'),
    # data checks: raw fields, helpers and one 0/1 column per check (check_formulas.py, reference in checks.py)
) + C.ITEMFLOW_CHECK_COLUMNS

# --- DailyLoad: one row per day and team ------------------------------------------------------------------

def _snapshot(day: str) -> str:
    return f"({day}+1+{HALF})"


def load_count(day: str, team: str, bucket: str | None = None) -> str:
    """Issues in process at the end of `day` (midnight of day + 1), for the mode in cfg_FlowLoadMode.

    Mode A counts the StatusIntervals that contain the instant in a category that is work in process, for issues
    that have started; the intervals of one issue do not overlap, so intervals and issues are the same count.
    Mode B counts the issues whose last start is not after the instant and whose end is.
    """
    t = _snapshot(day)
    a_bucket = f",{SI}[Bucket],{bucket}" if bucket else ""
    b_bucket = f",{IF_}[WorkTypeBucket],{bucket}" if bucket else ""
    mode_a = (f'COUNTIFS({SI}[EnteredUtc],"<="&{t},{SI}[LeftEff],">"&{t},{SI}[Counted],1,'
              f'{SI}[FirstStart],"<="&{t},{SI}[Team],{team}{a_bucket})')
    mode_b = (f'COUNTIFS({IF_}[Team],{team},{IF_}[StartUtc],"<="&{t},{IF_}[EndUtc],">"&{t}{b_bucket})')
    return f'IF(cfg_FlowLoadMode="A",{mode_a},IF(cfg_FlowLoadMode="B",{mode_b},NA()))'


DAILYLOAD_LOAD = f'=IF(ROUND(([@Date]+1-AsOfUtc)*86400,0)>0,"",{load_count("[@Date]", "[@Team]")})'


# --- Metrics: one row per iteration and team (docs 2.2) -------------------------------------------------------

DONE_IN = (f'{IF_}[DoneUtc],">="&([@IterationStart]-{HALF}),{IF_}[DoneUtc],"<"&([@IterationEnd]-{HALF}),'
           f"{IF_}[Team],[@Team]")
_DONE_ARRAY = (f"({IF_}[DoneUtc]>=[@IterationStart]-{HALF})*({IF_}[DoneUtc]<[@IterationEnd]-{HALF})*"
               f"({IF_}[Team]=[@Team])*({IF_}[Excluded]=0)")
_LOAD_WINDOW = (f'DailyLoad[Team],[@Team],DailyLoad[Date],">="&([@IterationStart]-{HALF}),'
                f'DailyLoad[Date],"<"&([@IterationEnd]-{HALF})')
_NC = f'"{NOT_COMPUTABLE}"'


def _sum(column: str) -> str:
    return f"SUMIFS({IF_}[{column}],{DONE_IN},{IF_}[Excluded],0)"


METRICS_KEYS = ("Iteration-label", "IterationStart", "IterationEnd", "Team")

METRICS_COLUMNS = (
    ("Velocity", f"=COUNTIFS({DONE_IN})"),
    ("FlowTimeCount", f"=COUNTIFS({DONE_IN},{IF_}[Excluded],0)"),
    ("FlowTimeExcluded", "=[@Velocity]-[@FlowTimeCount]"),
    ("FlowTimeMedian", f"=IF([@FlowTimeCount]=0,{_NC},MEDIAN(FILTER({IF_}[FlowTimeDays],{_DONE_ARRAY})))"),
    ("FlowTimeP85", f"=IF([@FlowTimeCount]=0,{_NC},PERCENTILE.INC(FILTER({IF_}[FlowTimeDays],{_DONE_ARRAY}),0.85))"),
    ("FlowLoadAverage", f"=IFERROR(AVERAGEIFS(DailyLoad[Load],{_LOAD_WINDOW}),{_NC})"),
    ("EndDay", f'=IF(COUNTIFS({_LOAD_WINDOW},DailyLoad[Load],">=0")=0,"",'
               f'MAXIFS(DailyLoad[Date],{_LOAD_WINDOW},DailyLoad[Load],">=0"))'),
    ("FlowLoadEnd", f'=IF([@EndDay]="",{_NC},SUMIFS(DailyLoad[Load],DailyLoad[Team],[@Team],DailyLoad[Date],[@EndDay]))'),
    ("FlowEfficiency", f"=IF([@FlowTimeCount]=0,{_NC},IF({_sum('UnclassifiedDays')}>0,{_NC},"
                       f"IF({_sum('ActiveDays')}+{_sum('WaitDays')}=0,{_NC},"
                       f"{_sum('ActiveDays')}/({_sum('ActiveDays')}+{_sum('WaitDays')}))))"),
) + tuple(
    (f"DistDone{b}", f'=IF([@Velocity]=0,{_NC},COUNTIFS({DONE_IN},{IF_}[WorkTypeBucket],"{b}")/[@Velocity])')
    for b in BUCKETS
) + tuple(
    (f"WipCount{b}", f'=IF([@EndDay]="","",{load_count("[@EndDay]", "[@Team]", chr(34) + b + chr(34))})')
    for b in BUCKETS
) + tuple(
    (f"DistWip{b}", f'=IF(OR([@EndDay]="",[@FlowLoadEnd]=0),{_NC},[@WipCount{b}]/[@FlowLoadEnd])')
    for b in BUCKETS
) + (("FlowPredictability", f"={_NC}"),)

# --- the table VelocityCheck: the Velocity counted again from the status changes of the changelog (AL, 2026-11-18) ----------------
# Same keys as Metrics (one row per iteration and team: VBA writes them with the same code). All figures are counts over StatusIntervals; an
# iteration with nothing gives 0, never "not computable". Reference: flow/velocity_check.py (its docstring says what each column means).
_END_IN = (f'{SI}[IsEnd],1,{SI}[EnteredUtc],">="&([@IterationStart]-{HALF}),{SI}[EnteredUtc],"<"&([@IterationEnd]-{HALF}),'
           f"{SI}[Team],[@Team]")
VELOCITYCHECK_KEYS = METRICS_KEYS
VELOCITYCHECK_COLUMNS = (
    ("Velocity", f"=COUNTIFS({DONE_IN})"),
    ("VelocityFeatures", f"=COUNTIFS({DONE_IN},{IF_}[IsFeature],1)"),
    # the tickets of the Program that are not Features, done in the iteration (the ART rows only): the team "ART (other types)" of ItemFlow
    ("OtherTypes", f'=IF([@Team]="{ART}",COUNTIFS({IF_}[DoneUtc],">="&([@IterationStart]-{HALF}),{IF_}[DoneUtc],"<"&([@IterationEnd]-{HALF}),'
                   f'{IF_}[Team],"{OTHER_TYPES}"),0)'),
    ("CreatedDone", f"=COUNTIFS({_END_IN},{SI}[Seq],1,{SI}[IsOpen],TRUE)"),
    ("ChangedToEnd", f'=COUNTIFS({_END_IN},{SI}[Seq],">1")'),
    ("ChangedToEndFeatures", f'=COUNTIFS({_END_IN},{SI}[Seq],">1",{SI}[ItemIsFeature],1)'),
    ("ReopenedLater", f'=COUNTIFS({_END_IN},{SI}[Seq],">1",{SI}[IsOpen],FALSE)'),
    ("Gap", "=[@Velocity]-[@CreatedDone]-([@ChangedToEnd]-[@ReopenedLater])"),
)

# --- the sheet MetricItems: the tickets behind the figures (AL, 2026-11-18) ---------------------------------------------------------
# Three lists side by side, each column ONE spilled formula (Excel 365 dynamic arrays, as FILTER is already used): the tickets of ItemFlow that
# fall in an account, sorted. No VBA writes a key: the lists follow the data. Reference: flow/trace.metric_items.
_C = f'({IF_}[Account]="{T.ACCOUNT_COUNTED}")'
_LEFT_OUT = f'({IF_}[DoneUtc]<>"")*({IF_}[Account]<>"{T.ACCOUNT_COUNTED}")'
_NEAR = f'ISNUMBER({IF_}[BoundaryHours])*({IF_}[BoundaryHours]<={T.CUTOFF_HOURS})*({IF_}[Account]<>"{T.ACCOUNT_OTHER}")'


def _col(column: str, cond: str, sort: tuple[str, ...], wrap: str = "{}") -> str:
    """One column of a list: `column` of the rows of ItemFlow that meet `cond` (wrapped by `wrap`, "{}" = itself), sorted by `sort`. A column per
    formula, all sorted by the same keys, so the cells of one row stay together; an empty result gives the empty text."""
    def fil(name: str) -> str:
        return f"FILTER({IF_}[{name}],{cond})"
    keys = ",".join(f"{fil(c)},1" for c in sort)
    return f'=IFERROR(SORTBY({wrap.format(fil(column))},{keys}),"")'


METRICITEMS_TITLES = ("Done in an iteration: the tickets behind Velocity, Flow Time, Flow Efficiency and Flow Distribution",
                      "Done but not counted: the missing ones, with the reason",
                      f"Cut-off: done within {T.CUTOFF_HOURS} hours of an iteration boundary")
_DONE_SORT = ("DoneIterationStart", "Team", "DoneUtc", "Key")
METRICITEMS_DONE = tuple((name, _col(column, _C, _DONE_SORT, wrap)) for name, column, wrap in (
    ("Iteration", "DoneIteration", "{}"), ("Start", "DoneIterationStart", "{}"), ("End", "DoneIterationEnd", "{}"), ("Team", "Team", "{}"),
    ("Key", "Key", "{}"), ("Type", "Type", "{}"), ("DoneUtc", "DoneUtc", "{}"), ("FlowTimeCounted", "Excluded", "1-{}"),
    ("FlowTimeDays", "FlowTimeDays", "{}"), ("WhyNotInFlowTime", "WhyNotInFlowTime", "{}")))
_LEFT_SORT = ("Account", "DoneUtc", "Key")
METRICITEMS_LEFT_OUT = tuple((name, _col(column, _LEFT_OUT, _LEFT_SORT)) for name, column in (
    ("Key", "Key"), ("Type", "Type"), ("Team", "Team"), ("Iteration", "DoneIteration"), ("DoneUtc", "DoneUtc"), ("Account", "Account")))
_NEAR_SORT = ("BoundaryHours", "Key")
METRICITEMS_CUTOFF = tuple((name, _col(column, _NEAR, _NEAR_SORT, wrap)) for name, column, wrap in (
    ("Key", "Key", "{}"), ("Team", "Team", "{}"), ("Iteration", "DoneIteration", "{}"), ("DoneUtc", "DoneUtc", "{}"),
    ("Hours", "BoundaryHours", "ROUND({},1)")))
METRICITEMS_BLOCKS = (METRICITEMS_DONE, METRICITEMS_LEFT_OUT, METRICITEMS_CUTOFF)

# --- the sheet Reconciliation: the controls on the ledger, the census of the types, the parameters in force (AL, 2026-11-18) ------------
# The table Reconciliation: one row per control (Control, Assertion: written by VBA from RECONCILIATION_CONTROLS), Expected and Found are a formula of
# their own in each row (RECONCILIATION_EXPECTED / FOUND), Difference and Result are calculated columns. The ledger is the first list of MetricItems
# (its spilled columns: A Iteration, D Team, E Key, H FlowTimeCounted, I FlowTimeDays) and the third (S Key). Reference: flow/reconciliation.py.
_M = "MetricItems_from_excel"
_LEDGER = {"iteration": f"{_M}!$A$3#", "team": f"{_M}!$D$3#", "key": f"{_M}!$E$3#", "counted": f"{_M}!$H$3#", "days": f"{_M}!$I$3#"}
_ACCOUNTS_ARRAY = "{" + ",".join(f'"{a}"' for a in T.ACCOUNTS) + "}"
RECONCILIATION_FORMULAS = {
    R.POPULATION: ('=COUNTIF(Issues[Key],"?*")', '=COUNTIF(ItemFlow[Key],"?*")'),
    R.ACCOUNTS_ADD_UP: ('=COUNTIF(ItemFlow[Key],"?*")', f"=SUMPRODUCT(COUNTIF(ItemFlow[Account],{_ACCOUNTS_ARRAY}))"),
    R.VELOCITY_TOTAL: ("=SUM(Metrics[Velocity])", f'=COUNTIF({_LEDGER["key"]},"?*")'),
    R.VELOCITY_ROWS: ("=0", f'=MAX(ABS(Metrics[Velocity]-COUNTIFS({_LEDGER["iteration"]},Metrics[Iteration-label],{_LEDGER["team"]},Metrics[Team])))'),
    R.FLOW_TIME_COUNT: ("=SUM(Metrics[FlowTimeCount])", f'=SUM({_LEDGER["counted"]})'),
    R.FLOW_TIME_DAYS: (f'=SUMIFS(ItemFlow[FlowTimeDays],ItemFlow[Account],"{T.ACCOUNT_COUNTED}",ItemFlow[Excluded],0)',
                       f'=SUMIFS({_LEDGER["days"]},{_LEDGER["counted"]},1)'),
    R.UNIQUE: ("=0", f'=COUNTIF({_LEDGER["key"]},"?*")-IFERROR(ROWS(UNIQUE(FILTER({_LEDGER["key"]},{_LEDGER["key"]}<>""))),0)'),
    R.CUTOFF: ("=0", f'=COUNTIF({_M}!$S$3#,"?*")'),
    R.CLASSIFIED: ("=0", f'=COUNTIFS({SI}[Category],"Unclassified",{SI}[Team],"<>{OTHER_TYPES}")'),
    R.FEATURE_TYPES: ("=1", f'=IF(COUNTIFS(ItemFlow[Role],"Program")=0,1,MIN(1,COUNTIFS(ItemFlow[Team],"{ART}")))'),
    R.HISTORIES: ("=0", "=SUM(ItemFlow[DC06])+SUM(ItemFlow[DC07])"),
    R.EXISTENCE: ("=COUNT(ItemFlow[DoneUtc])", f"=COUNTIFS({SI}[IsEnd],1,{SI}[IsOpen],TRUE)"),
}
assert tuple(RECONCILIATION_FORMULAS) == tuple(name for name, _ in R.CONTROLS), "one pair of formulas per control, in the order of the catalogue"
RECONCILIATION_COLUMNS = (("Control", ""), ("Assertion", ""), ("Expected", ""), ("Found", ""),
                          ("Difference", "=[@Found]-[@Expected]"), ("Result", '=IF(ABS([@Difference])<0.000001,"OK",IF(OR(' + ",".join(f'[@Assertion]="{a}"' for a in R.INFO_ASSERTIONS) + '),"INFO","REVIEW"))'))
# the census of the types (spilled formulas right of the table: the list of the types, then one count per type)
_TYPES = "$H$2#"
CENSUS_COLUMNS = (
    ("Type", '=IFERROR(SORT(UNIQUE(ItemFlow[Type])),"")'),
    ("Tickets", f'=IFERROR(COUNTIFS(ItemFlow[Type],{_TYPES}),"")'),
    ("InART", f'=IFERROR(COUNTIFS(ItemFlow[Type],{_TYPES},ItemFlow[Team],"{ART}"),"")'),
    ("LeftOut", f'=IFERROR(COUNTIFS(ItemFlow[Type],{_TYPES},ItemFlow[Team],"{OTHER_TYPES}"),"")'),
    ("Done", f'=IFERROR(COUNTIFS(ItemFlow[Type],{_TYPES},ItemFlow[DoneUtc],">0"),"")'),
    ("DoneInAnIteration", f'=IFERROR(COUNTIFS(ItemFlow[Type],{_TYPES},ItemFlow[DoneIteration],"?*"),"")'),
)
# --- the table PIMetrics: Feature Predictability per PI (AL's variant, 2026-11-12) -----------------------------
# The key PI comes from the sheet PIs, then from the Features whose PI the sheet lacks (modFlowBuild.PiNames); its dates
# are looked up in the sheet: a PI missing there or whose date is not a date gives "" and its figures are not computable.
# A PI of the sheet is judged on what was ENGAGED at its start (AL, 2026-11-16): the figures are sums of the rows of the ART in
# IterationFeatures (below), read in the history of the PI field, so a Feature deferred later is still planned and a cancelled
# one stays planned. Planned = Engaged at the first iteration of the PI; done = ReleasedPlanned summed over its iterations
# (mode A of cfg_PredictabilityMode) or Released (mode B: also the Features added after the start). A PI the sheet lacks has no
# iteration: its Features are counted by their PI field as it is now, and its ratio is not computable.
PIMETRICS_KEYS = ("PI",)
_PI_ART = 'IterationFeatures[PI],[@PI],IterationFeatures[Team],"ART"'


def _pi_date(column: str) -> str:
    # a date typed in Excel is a number; Load settings CSV writes an ISO text ('2026-12-01): DATEVALUE reads it, as DC18
    return f'=IFERROR(LET(v,INDEX(PIs_from_user!${column}:${column},MATCH([@PI],PIs_from_user!$A:$A,0)),IF(ISNUMBER(v),IF(v>0,v,""),DATEVALUE(v))),"")'


PIMETRICS_COLUMNS = (
    ("Start", _pi_date("B")),
    ("End", _pi_date("C")),
    ("FeaturesPlanned", f'=IF(ISNUMBER([@Start]),SUMIFS(IterationFeatures[Engaged],{_PI_ART},IterationFeatures[IterationStart],'
                        f'MINIFS(IterationFeatures[IterationStart],{_PI_ART})),COUNTIFS(ItemFlow[IsFeature],1,ItemFlow[PI],[@PI]))'),
    ("FeaturesDone", f'=IF(ISNUMBER([@Start]),IF(cfg_PredictabilityMode="B",SUMIFS(IterationFeatures[Released],{_PI_ART}),'
                     f'SUMIFS(IterationFeatures[ReleasedPlanned],{_PI_ART})),"")'),
    ("FeaturePredictability", f"=IF(OR([@FeaturesPlanned]=0,NOT(ISNUMBER([@FeaturesDone]))),{_NC},"
                              "[@FeaturesDone]/[@FeaturesPlanned])"),
)

# --- the table PIStatusTime: where the cycle time of the Features goes, status by status, PI by PI (AL, 2026-11-17) -------------
# One row per PI (the keys of PIMetrics) and per Active / Wait status of the Feature workflow (VBA writes the keys: modFlowBuild.StatusNames, in
# the order of the sheet Workflows). Items = the Features released in the PI (DoneUtc in [Start, End), not excluded from flow time);
# TotalDays = the time those items spent in the status (StatusIntervals[InCycle]); AverageDays = TotalDays / Items; Share = TotalDays /
# the total of the statuses of the PI. A PI without dates or without a released Feature is not computable. Reference: flow/pi_status_time.py.
PISTATUSTIME_KEYS = ("PI", "Status")
_PI_DONE = f'ItemFlow[DoneUtc],">="&([@Start]-{HALF}),ItemFlow[DoneUtc],"<"&([@End]-{HALF})'
PISTATUSTIME_COLUMNS = (
    ("Start", _pi_date("B")),
    ("End", _pi_date("C")),
    ("Items", f'=IF(AND(ISNUMBER([@Start]),ISNUMBER([@End])),COUNTIFS(ItemFlow[Team],"{ART}",{_PI_DONE},ItemFlow[Excluded],0),{_NC})'),
    ("TotalDays", f'=IF(ISNUMBER([@Items]),SUMIFS({SI}[DurationDays],{SI}[Team],"{ART}",{SI}[Status],[@Status],{SI}[InCycle],1,'
                  f'{SI}[ItemDone],">="&([@Start]-{HALF}),{SI}[ItemDone],"<"&([@End]-{HALF})),{_NC})'),
    ("AverageDays", f"=IF(ISNUMBER([@Items]),IF([@Items]=0,{_NC},[@TotalDays]/[@Items]),{_NC})"),
    ("Share", f"=IF(ISNUMBER([@Items]),IF([@Items]=0,{_NC},IF(SUMIFS(PIStatusTime[TotalDays],PIStatusTime[PI],[@PI])>0,"
              f"[@TotalDays]/SUMIFS(PIStatusTime[TotalDays],PIStatusTime[PI],[@PI]),{_NC})),{_NC})"),
)

# --- PIIntervals: the PI field of each issue through time (raw: modFlow.PiIntervalRows; reference: flow/pi_history.py) --------
# One row per stretch in which an issue carried a PI: FromUtc included, ToUtc excluded (empty = still: IsOpen). The helpers give each
# row the Feature flag, the team of the Feature (the Team field; "(none)" when empty) and when it was done (far away when not).

PIINTERVALS_RAW = ("Key", "PI", "FromUtc", "ToUtc", "IsOpen")
_BY_KEY = "MATCH([@Key],ItemFlow[Key],0)"
PIINTERVAL_HELPERS = (
    # an empty table keeps one blank row (no stretch at all): IFERROR gives it neutral values instead of #N/A
    ("IsFeature", f"=IFERROR(INDEX(ItemFlow[IsFeature],{_BY_KEY}),0)"),
    ("FeatureTeam", f'=IFERROR(LET(v,TRIM(INDEX(ItemFlow[TeamValue],{_BY_KEY})),IF(v="","(none)",v)),"(none)")'),
    ("DoneKey", f'=IFERROR(LET(d,INDEX(ItemFlow[DoneUtc],{_BY_KEY}),IF(d="",{FAR},d)),{FAR})'),
    ("ToEff", f'=IF([@ToUtc]="",{FAR},[@ToUtc])'),
)

# --- IterationFeatures: the Features of the ART iteration by iteration (reference: flow/iteration_features.py) ------------
# One row per iteration, PI that holds its start (VBA writes the keys) and team (ART, then the teams of the Features). The counts are
# COUNTIFS on the stretches of PIIntervals: "at t" = FromUtc <= t < ToUtc; "through" = one stretch holds the start and the end.

ITERATIONFEATURES_KEYS = ("Iteration-label", "IterationStart", "IterationEnd", "Team", "PI")
_START, _END, _BASE = "[@IterationStart]", "[@IterationEnd]", "[@Baseline]"
_WHO = ('PIIntervals[PI],[@PI],PIIntervals[IsFeature],1,'
        'PIIntervals[FeatureTeam],IF([@Team]="ART","*",[@Team])')


def _held(first: str, last: str) -> str:
    """The stretches that begin at or before `first` and end after `last`."""
    return f'PIIntervals[FromUtc],"<="&({first}+{HALF}),PIIntervals[ToEff],">"&({last}+{HALF})'


_NOT_DONE_BEFORE_START = f'PIIntervals[DoneKey],">="&({_START}-{HALF})'
_DONE_IN = f'{_NOT_DONE_BEFORE_START},PIIntervals[DoneKey],"<"&({_END}-{HALF})'


def _count(first: str, last: str, done: str) -> str:
    return f"COUNTIFS({_WHO},{_held(first, last)},{done})"


_HOLDS_START = _count(_START, _START, _NOT_DONE_BEFORE_START)          # Engaged
_THROUGH = _count(_START, _END, _NOT_DONE_BEFORE_START)                # held the start and the end by one stretch, not done before
_HOLDS_END = _count(_END, _END, _NOT_DONE_BEFORE_START)
_RELEASED_START = _count(_START, _START, _DONE_IN)
_RELEASED_END = _count(_END, _END, _DONE_IN)
_RELEASED_THROUGH = _count(_START, _END, _DONE_IN)

ITERATIONFEATURES_COLUMNS = (
    ("Baseline", '=IF([@PI]="","",MINIFS(IterationFeatures[IterationStart],IterationFeatures[PI],[@PI]))'),
    ("Engaged", f'=IF([@PI]="",0,{_HOLDS_START})'),
    ("Added", f'=IF([@PI]="",0,{_HOLDS_END}-{_THROUGH})'),
    ("Deferred", f'=IF([@PI]="",0,{_HOLDS_START}-{_THROUGH}-({_RELEASED_START}-{_RELEASED_THROUGH}))'),
    ("Released", f'=IF([@PI]="",0,{_RELEASED_START}+{_RELEASED_END}-{_RELEASED_THROUGH})'),
    ("ReleasedPlanned", f'=IF([@PI]="",0,{_count(_BASE, _START, _DONE_IN)})'),
    ("ReleasedAdded", "=[@Released]-[@ReleasedPlanned]"),
    ("CarryOver", "=[@Engaged]+[@Added]-[@Deferred]-[@Released]"),
)
ITERATIONFEATURES_NAMES = tuple(name for name, _ in ITERATIONFEATURES_COLUMNS)

# --- control cells of the Metrics sheet: the as-of date and the Flow Load mode, read from Config ------------

CONFIG_VALUE = 'INDEX(Config_from_user!$B:$B,MATCH("{name}",Config_from_user!$A:$A,0))'
# A date typed in the cell is a number, ISO text (2026-11-03 18:00) is parsed, an empty cell is #N/A (not 1900).
AS_OF_CELL = ('=LET(v,' + CONFIG_VALUE.format(name="AsOf") + ',IF(ISNUMBER(v),v,'
              'LET(t,SUBSTITUTE(TRIM(v),"T"," "),'
              'IF(t="",NA(),DATEVALUE(LEFT(t,10))+IF(LEN(t)>10,TIMEVALUE(MID(t,12,8)),0)))))')
MODE_CELL = "=UPPER(TRIM(" + CONFIG_VALUE.format(name="FlowLoadMode") + "))"
PREDICTABILITY_MODE_CELL = "=UPPER(TRIM(" + CONFIG_VALUE.format(name="PredictabilityMode") + "))"
TRANSITIONS_LAST_CELL, PIS_LAST_CELL = last_row_cell("Transitions"), last_row_cell("PIs")

# the statuses that no rule classifies, by name (spilled formulas right of the census, from column O): the intervals of the counted tickets (what the
# control counts) and of the tickets left out
_UNK = "$O$2#"
UNCLASSIFIED_COLUMNS = (
    ("Status", f'=IFERROR(SORT(UNIQUE(FILTER({SI}[Status],{SI}[Category]="Unclassified"))),"")'),
    ("IntervalsOfCounted", f'=IFERROR(COUNTIFS({SI}[Status],{_UNK},{SI}[Category],"Unclassified",{SI}[Team],"<>{OTHER_TYPES}"),"")'),
    ("IntervalsOfLeftOut", f'=IFERROR(COUNTIFS({SI}[Status],{_UNK},{SI}[Category],"Unclassified",{SI}[Team],"{OTHER_TYPES}"),"")'),
)
# the parameters in force: what the figures read, written next to them so that the person can compare it with Jira
PARAMETERS_IN_FORCE = (
    ("FeatureTypes (the types of the ART)", "=" + CONFIG_VALUE.format(name="FeatureTypes")),
    ("Program project(s)", '=TEXTJOIN(", ",TRUE,FILTER(Projects_from_user!$A$2:$A$300,Projects_from_user!$B$2:$B$300="Program",""))'),
    ("Statuses that start the flow (IsStart)", '=TEXTJOIN(", ",TRUE,FILTER(Workflows_from_user!$A$2:$A$500&": "&Workflows_from_user!$B$2:$B$500,'
                                               'Workflows_from_user!$D$2:$D$500=1,""))'),
    ("Statuses that end it (IsEnd)", '=TEXTJOIN(", ",TRUE,FILTER(Workflows_from_user!$A$2:$A$500&": "&Workflows_from_user!$B$2:$B$500,'
                                     'Workflows_from_user!$E$2:$E$500=1,""))'),
    ("First iteration start", "=MIN(Metrics[IterationStart])"),
    ("Last iteration end (excluded)", "=MAX(Metrics[IterationEnd])"),
    ("AsOf (UTC)", "=AsOfUtc"),
    ("Team field", "=" + CONFIG_VALUE.format(name="TeamField")),
    ("PI field", "=" + CONFIG_VALUE.format(name="PIField")),
)

TABLES = {
    "Issues": ISSUES,
    "StatusIntervals": INTERVALS_RAW + tuple(name for name, _ in INTERVAL_HELPERS),
    "ItemFlow": tuple(name for name, _ in ITEMFLOW_COLUMNS),
    "DailyLoad": ("Date", "Team", "Load"),
    "IssueLinks": C.LINKS_RAW + tuple(name for name, _ in C.LINKS_COLUMNS),
    "FillRates": C.FILL_KEYS + tuple(name for name, _ in C.FILL_COLUMNS),
    "DataChecks": C.DATACHECKS_KEYS + tuple(name for name, _ in C.DATACHECKS_COLUMNS),
    "Metrics": METRICS_KEYS + tuple(name for name, _ in METRICS_COLUMNS),
    "PIIntervals": PIINTERVALS_RAW + tuple(name for name, _ in PIINTERVAL_HELPERS),
    "IterationFeatures": ITERATIONFEATURES_KEYS + ITERATIONFEATURES_NAMES,
    "PIMetrics": PIMETRICS_KEYS + tuple(name for name, _ in PIMETRICS_COLUMNS),
    "PIStatusTime": PISTATUSTIME_KEYS + tuple(name for name, _ in PISTATUSTIME_COLUMNS),
    "VelocityCheck": VELOCITYCHECK_KEYS + tuple(name for name, _ in VELOCITYCHECK_COLUMNS),
}
