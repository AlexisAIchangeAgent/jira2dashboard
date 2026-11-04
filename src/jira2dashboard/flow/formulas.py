"""The Excel formulas of the flow sheets, as text (English function names, structured references).

Pure: nothing here touches Excel. sheets.py writes them. The formulas follow docs section 2 and are what the
differential test compares, cell by cell, with the Python reference (itemflow.py, load.py, metrics.py).

Rules kept here:
* no status name and no mapping in a formula: the status flags and categories are read from the sheet Workflows,
  the teams and roles from Projects, the periods from Periods, the mode and the as-of date from Config;
* every comparison with a date adds or removes half a second: dates are whole seconds, and Excel turns a bound
  into text with 15 digits, so a strict `>` on the exact serial could let an equal date through;
* a blank result is the empty text "", which COUNTIFS / MINIFS / SUMIFS ignore in numeric criteria.
"""
HALF = "(0.5/86400)"
FAR = "2958465"                      # 9999-12-31: "no end"
NOT_COMPUTABLE = "not computable"
CATEGORIES_IN_PROCESS = ("Active", "Wait", "Unclassified")
BUCKETS = ("Feature", "Enabler", "Defect", "Autre", "Unmapped")
WORK_TYPES = BUCKETS[:-1]

ISSUES = ("Key", "Id", "Project", "Type", "WorkType", "CreatedUtc", "Status")
INTERVALS_RAW = ("Key", "Workflow", "Seq", "Status", "Category", "EnteredUtc", "LeftUtc", "DurationDays", "IsOpen",
                 "ChangeId", "Check")

SI, IF_, IS = "StatusIntervals", "ItemFlow", "Issues"


def _row_index(table: str) -> str:
    return f"ROW()-ROW({table}[#Headers])"


# --- StatusIntervals: helper columns after the raw ones ------------------------------------------------

INTERVAL_HELPERS = (
    ("IsStart", "=SUMIFS(Workflows!$E:$E,Workflows!$A:$A,[@Workflow],Workflows!$C:$C,[@Status])"),
    ("IsEnd", "=SUMIFS(Workflows!$F:$F,Workflows!$A:$A,[@Workflow],Workflows!$C:$C,[@Status])"),
    ("PrevCategory", f'=IF([@Seq]=1,"",INDEX({SI}[Category],{_row_index(SI)}-1))'),
    ("Team", f"=INDEX({IF_}[Team],MATCH([@Key],{IF_}[Key],0))"),
    ("Bucket", f"=INDEX({IF_}[WorkTypeBucket],MATCH([@Key],{IF_}[Key],0))"),
    ("FirstStart", f"=INDEX({IF_}[FirstStartUtc],MATCH([@Key],{IF_}[Key],0))"),
    ("LeftEff", f'=IF([@LeftUtc]="",{FAR},[@LeftUtc])'),
    ("Counted", "=IF(OR(" + ",".join(f'[@Category]="{c}"' for c in CATEGORIES_IN_PROCESS) + "),1,0)"),
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
    ("Role", '=IFERROR(INDEX(Projects!$B:$B,MATCH([@Project],Projects!$A:$A,0)),"")'),
    ("Workflow", '=IF([@Role]="Program","Feature","Team")'),
    ("Team", '=IF([@Role]="Program","ART",IFERROR(INDEX(Projects!$C:$C,MATCH([@Project],Projects!$A:$A,0)),""))'),
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
)


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
    return f'IF(cfg_FlowLoadMode="A",{mode_a},{mode_b})'


DAILYLOAD_LOAD = f'=IF(ROUND(([@Date]+1-AsOfUtc)*86400,0)>0,"",{load_count("[@Date]", "[@Team]")})'


# --- Metrics: one row per period and team (docs 2.2) -------------------------------------------------------

DONE_IN = (f'{IF_}[DoneUtc],">="&([@PeriodStart]-{HALF}),{IF_}[DoneUtc],"<"&([@PeriodEnd]-{HALF}),'
           f"{IF_}[Team],[@Team]")
_DONE_ARRAY = (f"({IF_}[DoneUtc]>=[@PeriodStart]-{HALF})*({IF_}[DoneUtc]<[@PeriodEnd]-{HALF})*"
               f"({IF_}[Team]=[@Team])*({IF_}[Excluded]=0)")
_LOAD_WINDOW = (f'DailyLoad[Team],[@Team],DailyLoad[Date],">="&([@PeriodStart]-{HALF}),'
                f'DailyLoad[Date],"<"&([@PeriodEnd]-{HALF})')
_NC = f'"{NOT_COMPUTABLE}"'


def _sum(column: str) -> str:
    return f"SUMIFS({IF_}[{column}],{DONE_IN},{IF_}[Excluded],0)"


METRICS_KEYS = ("Period", "PeriodStart", "PeriodEnd", "Team")

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

# --- control cells of the Metrics sheet: the as-of date and the Flow Load mode, read from Config ------------

CONFIG_VALUE = 'INDEX(Config!$B:$B,MATCH("{name}",Config!$A:$A,0))'
AS_OF_CELL = ('=LET(t,SUBSTITUTE(TRIM(' + CONFIG_VALUE.format(name="AsOf") + '),"T"," "),'
              'DATEVALUE(LEFT(t,10))+IF(LEN(t)>10,TIMEVALUE(MID(t,12,8)),0))')
MODE_CELL = "=UPPER(TRIM(" + CONFIG_VALUE.format(name="FlowLoadMode") + "))"

TABLES = {
    "Issues": ISSUES,
    "StatusIntervals": INTERVALS_RAW + tuple(name for name, _ in INTERVAL_HELPERS),
    "ItemFlow": tuple(name for name, _ in ITEMFLOW_COLUMNS),
    "DailyLoad": ("Date", "Team", "Load"),
    "Metrics": METRICS_KEYS + tuple(name for name, _ in METRICS_COLUMNS),
}
