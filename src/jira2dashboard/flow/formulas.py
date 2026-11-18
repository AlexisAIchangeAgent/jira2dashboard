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
    # the team of a Team project: its column Team, or the project key when that cell is empty (AL, 2026-11-18)
    ("Team", '=IF([@Role]="Program","ART",IFERROR(IF(INDEX(Projects_from_user!$C:$C,MATCH([@Project],Projects_from_user!$A:$A,0))="",[@Project],'
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
    ("Items", f'=IF(AND(ISNUMBER([@Start]),ISNUMBER([@End])),COUNTIFS(ItemFlow[Workflow],"Feature",{_PI_DONE},ItemFlow[Excluded],0),{_NC})'),
    ("TotalDays", f'=IF(ISNUMBER([@Items]),SUMIFS({SI}[DurationDays],{SI}[Workflow],"Feature",{SI}[Status],[@Status],{SI}[InCycle],1,'
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
}
