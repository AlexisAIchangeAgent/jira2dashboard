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
from jira2dashboard.flow import metric_checks as MC
from jira2dashboard.flow import reconciliation as R
from jira2dashboard.flow import trace as T
from jira2dashboard.flow.model import ART, OTHER_TYPES
from jira2dashboard.flow.refs import (HALF, IF_, IS, PIS_LAST, SI, TRANSITIONS_LAST, bounded, jira_range, last_row_cell,
                                      row_index)

RULES_SHEET = "StatusRules_from_excel"       # the effective rules of every status (Jira's, under the exceptions of Workflows_from_user): modExtract writes it
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
    ("IsStart", "=SUMIFS(StatusRules_from_excel!$D:$D,StatusRules_from_excel!$A:$A,[@Workflow],StatusRules_from_excel!$B:$B,[@Status])"),
    ("IsEnd", "=SUMIFS(StatusRules_from_excel!$E:$E,StatusRules_from_excel!$A:$A,[@Workflow],StatusRules_from_excel!$B:$B,[@Status])"),
    ("PrevCategory", f'=IF([@Seq]=1,"",INDEX({SI}[Category],{_row_index(SI)}-1))'),
    ("Team", f"=INDEX({IF_}[Team],MATCH([@Key],{IF_}[Key],0))"),
    ("Bucket", f"=INDEX({IF_}[WorkTypeBucket],MATCH([@Key],{IF_}[Key],0))"),
    ("FirstStart", f"=INDEX({IF_}[FirstStartUtc],MATCH([@Key],{IF_}[Key],0))"),
    ("LeftEff", f'=IF([@LeftUtc]="",{FAR},[@LeftUtc])'),
    ("Counted", "=IF(OR(" + ",".join(f'[@Category]="{c}"' for c in CATEGORIES_IN_PROCESS) + "),1,0)"),
    ("PrevStatus", f'=IF([@Seq]=1,"",INDEX({SI}[Status],{_row_index(SI)}-1))'),
    ("JiraWorkflow", f'=IFERROR(INDEX({IF_}[JiraWorkflow],MATCH([@Key],{IF_}[Key],0)),"")'),
    # DC28: the move from the previous status is allowed by neither the Jira workflow of the ticket (JiraTransitions) nor the sheet Transitions (same
    # workflow kind) and the status is not free (IsFree). A ticket whose moves nobody knows (no row of Transitions for its kind, no Jira workflow read for it)
    # is not judged: 0 (flow/checks.Moves.permits)
    ("BadMove", f'=IF([@Seq]=1,0,IF(OR(COUNTIFS({bounded("Transitions", "A", TRANSITIONS_LAST)},[@Workflow],'
                f'{bounded("Transitions", "B", TRANSITIONS_LAST)},[@PrevStatus],'
                f'{bounded("Transitions", "C", TRANSITIONS_LAST)},[@Status])>0,'
                f'AND([@JiraWorkflow]<>"",COUNTIFS({jira_range("JiraTransitions", "A")},[@JiraWorkflow],{jira_range("JiraTransitions", "B")},[@PrevStatus],'
                f'{jira_range("JiraTransitions", "C")},[@Status])>0),'
                'SUMIFS(StatusRules_from_excel!$F:$F,StatusRules_from_excel!$A:$A,[@Workflow],StatusRules_from_excel!$B:$B,[@Status])>0,'
                f'AND(COUNTIFS({bounded("Transitions", "A", TRANSITIONS_LAST)},[@Workflow])=0,'
                f'OR([@JiraWorkflow]="",COUNTIFS({jira_range("JiraTransitions", "A")},[@JiraWorkflow])=0))),0,1))'),
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


def _starts_of(workflow: str) -> str:
    """How many statuses start the flow of the workflow kind `workflow` in the sheet StatusRules (SUMIFS stops at the used range)."""
    return f"SUMIFS(StatusRules_from_excel!$D:$D,StatusRules_from_excel!$A:$A,{workflow})"


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
    # the canonical name of a known work type whatever its case (= ignores it), else Unmapped: model.work_type_bucket
    ("WorkTypeBucket", "=" + "".join(f'IF([@WorkType]="{w}","{w}",' for w in WORK_TYPES) + '"Unmapped"' + ")" * len(WORK_TYPES)),
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
    # why a done ticket has no start: no IsStart status in the rules of its workflow (a setting), or none entered after its last backlog entry
    ("WhyNotInFlowTime", f'=IF([@Excluded]=0,"",TEXTJOIN(", ",TRUE,IF([@Reopened]>0,"{T.REOPENED}",""),IF([@SkippedStart]="Y",'
                         f'IF({_starts_of("[@Workflow]")}=0,"{T.NO_START_RULE}","{T.NO_START}"),"")))'),
    ("BoundaryHours", '=IF([@DoneUtc]="","",MIN(MIN(ABS(Metrics[IterationStart]-[@DoneUtc])),MIN(ABS(Metrics[IterationEnd]-[@DoneUtc])))*24)'),
    # data checks: raw fields, helpers and one 0/1 column per check (check_formulas.py, reference in checks.py)
) + C.ITEMFLOW_CHECK_COLUMNS + (
    # the Jira workflow of the ticket: that of its project and issue type in the sheet JiraIssueTypes ("" when the configuration was not read). It reads Project
    # and Type only, and StatusIntervals.JiraWorkflow reads it (check DC28)
    ("JiraWorkflow", f'=IFERROR(INDEX({jira_range("JiraIssueTypes", "C")},MATCH(1,INDEX(({jira_range("JiraIssueTypes", "A")}=[@Project])*'
                     f'({jira_range("JiraIssueTypes", "B")}=[@Type]),0),0))&"","")'),
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
    return f'IF(cfg_FlowLoadMode="A",{mode_a},IF(cfg_FlowLoadMode="B",{mode_b},NA()))'


# a day after AsOf has no snapshot; a team whose workflow has no start status has nothing ever in process: its load is unknown, never a false 0
# (AL, 2026-11-22). The ART follows the Feature workflow, a team of a Team project the Team workflow.
_TEAM_WORKFLOW = f'IF([@Team]="{ART}","Feature","Team")'
DAILYLOAD_LOAD = (f'=IF(ROUND(([@Date]+1-AsOfUtc)*86400,0)>0,"",IF({_starts_of(_TEAM_WORKFLOW)}=0,"",'
                  f'{load_count("[@Date]", "[@Team]")}))')


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
    # Flow Efficiency needs a status declared Wait: Jira cannot say which in-progress statuses wait, and without one the ratio would be 100 % by construction
    ("FlowEfficiency", f'=IF(COUNTIF({RULES_SHEET}!$C$2:$C$5000,"Wait")=0,{_NC},IF([@FlowTimeCount]=0,{_NC},IF({_sum("UnclassifiedDays")}>0,{_NC},'
                       f'IF({_sum("ActiveDays")}+{_sum("WaitDays")}=0,{_NC},'
                       f'{_sum("ActiveDays")}/({_sum("ActiveDays")}+{_sum("WaitDays")})))))'),
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

# --- the table FlowTimeCheck: why the Flow Time of each iteration and team is what it is (AL, 2026-11-22) ----------------------------------------
# The keys of Metrics. Done counted again from the intervals (an open end status entered in the iteration: an independent path), the done items left
# out of Flow Time each under ONE reason (reopened, then no start status in the rules, then no start after the backlog: WhyNotInFlowTime), Counted
# = Metrics.FlowTimeCount, Gap = Done - the three - Counted (0), Reason in clear when the point is not computable. Reference: flow/metric_checks.py.
FLOWTIMECHECK_KEYS = METRICS_KEYS
FLOWTIMECHECK_COLUMNS = (
    ("Done", f"=COUNTIFS({_END_IN},{SI}[IsOpen],TRUE)"),
    ("ExclReopened", f'=COUNTIFS({DONE_IN},{IF_}[Reopened],">0")'),
    ("ExclNoStartRule", f'=COUNTIFS({DONE_IN},{IF_}[Reopened],0,{IF_}[WhyNotInFlowTime],"{T.NO_START_RULE}")'),
    ("ExclNotStarted", f'=COUNTIFS({DONE_IN},{IF_}[Reopened],0,{IF_}[WhyNotInFlowTime],"{T.NO_START}")'),
    ("Counted", "=SUMIFS(Metrics[FlowTimeCount],Metrics[Iteration-label],[@[Iteration-label]],Metrics[Team],[@Team])"),
    ("Gap", "=[@Done]-[@ExclReopened]-[@ExclNoStartRule]-[@ExclNotStarted]-[@Counted]"),
    ("Reason", f'=IF([@Counted]>0,"",IF([@Done]=0,"{MC.NO_ITEM}","{MC.ALL}"&[@Done]&"{MC.LEFT_OUT}"&TEXTJOIN(", ",TRUE,'
               f'IF([@ExclNoStartRule]>0,[@ExclNoStartRule]&"{MC.NO_START_RULE}",""),IF([@ExclNotStarted]>0,[@ExclNotStarted]&"{MC.NOT_STARTED}",""),'
               f'IF([@ExclReopened]>0,[@ExclReopened]&"{MC.REOPENED}",""))))'),
)

# --- the table FlowEfficiencyCheck: why the Flow Efficiency of each iteration and team is what it is (AL, 2026-11-22) ----------------------------
# The keys of Metrics. Counted (Metrics.FlowTimeCount), the counted items with time in a status no rule knows, the Active and Wait days summed again
# from StatusIntervals (InCycle: the window of ItemFlow.ActiveDays / WaitDays, an independent path), the Wait statuses of StatusRules, the efficiency
# recounted, Reason in the order of the formula of Metrics, Gap = recount - Metrics.FlowEfficiency where it is computable. Reference: metric_checks.py.
_ITEM_DONE_IN = f'{SI}[ItemDone],">="&([@IterationStart]-{HALF}),{SI}[ItemDone],"<"&([@IterationEnd]-{HALF})'
_METRIC_OF_ROW = "XLOOKUP(1,(Metrics[Iteration-label]=[@[Iteration-label]])*(Metrics[Team]=[@Team]),Metrics[{column}])"
FLOWEFFICIENCYCHECK_KEYS = METRICS_KEYS
FLOWEFFICIENCYCHECK_COLUMNS = (
    ("Counted", "=SUMIFS(Metrics[FlowTimeCount],Metrics[Iteration-label],[@[Iteration-label]],Metrics[Team],[@Team])"),
    ("WithUnclassified", f'=COUNTIFS({DONE_IN},{IF_}[Excluded],0,{IF_}[UnclassifiedDays],">0")'),
    ("ActiveDays", f'=SUMIFS({SI}[DurationDays],{SI}[Team],[@Team],{SI}[InCycle],1,{SI}[Category],"Active",{_ITEM_DONE_IN})'),
    ("WaitDays", f'=SUMIFS({SI}[DurationDays],{SI}[Team],[@Team],{SI}[InCycle],1,{SI}[Category],"Wait",{_ITEM_DONE_IN})'),
    ("WaitStatuses", f'=COUNTIF({RULES_SHEET}!$C$2:$C$5000,"Wait")'),
    ("Efficiency", '=IF([@ActiveDays]+[@WaitDays]=0,"",[@ActiveDays]/([@ActiveDays]+[@WaitDays]))'),
    ("Gap", f'=IF([@Reason]<>"",0,IFERROR(ROUND([@Efficiency]-{_METRIC_OF_ROW.format(column="FlowEfficiency")},{MC.GAP_DIGITS}),1))'),
    ("Reason", f'=IF([@WaitStatuses]=0,"{MC.NO_WAIT}",IF([@Counted]=0,"{MC.NOTHING_COUNTED}",IF([@WithUnclassified]>0,'
               f'[@WithUnclassified]&"{MC.UNCLASSIFIED_TIME}",IF([@ActiveDays]+[@WaitDays]=0,"{MC.NO_TIME}",""))))'),
)

# --- the table FlowLoadCheck: why the Flow Load of each iteration and team is what it is (AL, 2026-11-22) ----------------------------------------
# The keys of Metrics. The days of the iteration, those with a snapshot (not after AsOf), the start statuses of the workflow of the team, the average
# of DailyLoad over the measured days, the items in process at the end of the last measured day counted again with SUMPRODUCT (another writing of
# the rule than the COUNTIFS of DailyLoad: mode A over StatusIntervals, mode B over ItemFlow), Gap against Metrics, Reason. Reference: metric_checks.
_LOAD_END = f"[@IterationStart]+[@DaysMeasured]+{HALF}"
_IN_PROCESS_A = (f'SUMPRODUCT(({SI}[Team]=[@Team])*({SI}[Counted]=1)*({SI}[EnteredUtc]<=t)*({SI}[LeftEff]>t)*({SI}[FirstStart]<>"")'
                 f'*({SI}[FirstStart]<=t))')
_IN_PROCESS_B = f'SUMPRODUCT(({IF_}[Team]=[@Team])*({IF_}[StartUtc]<>"")*({IF_}[StartUtc]<=t)*({IF_}[EndUtc]>t))'
FLOWLOADCHECK_KEYS = METRICS_KEYS
FLOWLOADCHECK_COLUMNS = (
    ("Days", "=[@IterationEnd]-[@IterationStart]"),
    ("DaysMeasured", f"=IFERROR(MAX(0,MIN([@IterationEnd],INT(AsOfUtc+{HALF}))-[@IterationStart]),0)"),      # AsOf unreadable: no snapshot
    ("StartStatuses", f"={_starts_of(_TEAM_WORKFLOW)}"),
    ("Average", f'=IF([@Reason]<>"","",SUMIFS(DailyLoad[Load],{_LOAD_WINDOW})/[@DaysMeasured])'),
    ("AtEnd", f'=IF([@Reason]<>"","",LET(t,{_LOAD_END},IF(cfg_FlowLoadMode="A",{_IN_PROCESS_A},{_IN_PROCESS_B})))'),
    ("Gap", f'=IF([@Reason]<>"",0,IFERROR(ROUND([@Average]-{_METRIC_OF_ROW.format(column="FlowLoadAverage")},{MC.GAP_DIGITS})'
            f'+([@AtEnd]-{_METRIC_OF_ROW.format(column="FlowLoadEnd")}),1))'),
    ("Reason", f'=IF([@StartStatuses]=0,"{MC.NO_START_STATUS}",IF([@DaysMeasured]=0,"{MC.AFTER_AS_OF}",""))'),
)

# --- the table FlowDistCheck: what the Flow Distribution of each iteration and team counts (AL, 2026-11-22) ------------------------------
# The keys of Metrics. The done items counted again per work type from the raw WorkType (COUNTIFS ignores the case, as WorkTypeBucket), Unmapped the
# rest, Gap = sum of |count - DistDone x Velocity| of Metrics, Reason (no item done) and a Note when every done item is Unmapped. Reference: metric_checks.
FLOWDISTRIBUTIONCHECK_KEYS = METRICS_KEYS
_CONFIG_WORK_TYPE_FIELD = 'INDEX(Config_from_user!$B:$B,MATCH("WorkTypeField",Config_from_user!$A:$A,0))'   # as CONFIG_VALUE, defined below
FLOWDISTRIBUTIONCHECK_COLUMNS = (
    ("Done", f"=COUNTIFS({DONE_IN})"),
    *((w, f'=COUNTIFS({DONE_IN},{IF_}[WorkType],"{w}")') for w in WORK_TYPES),
    ("Unmapped", "=[@Done]-" + "-".join(f"[@{w}]" for w in WORK_TYPES)),
    ("Gap", "=ROUND(" + "+".join(f'ABS([@{b}]-IFERROR({_METRIC_OF_ROW.format(column="DistDone" + b)}*{_METRIC_OF_ROW.format(column="Velocity")},0))'
                                 for b in BUCKETS) + f",{MC.GAP_DIGITS})"),
    ("Reason", f'=IF([@Done]=0,"{MC.NO_ITEM}","")'),
    ("Note", f'=IF(AND([@Done]>0,[@Unmapped]=[@Done]),"{MC.ALL_UNMAPPED}"&IF(TRIM(SUBSTITUTE(CLEAN(IFERROR({_CONFIG_WORK_TYPE_FIELD},"")&""),CHAR(160)," "))="",'
             f'"{MC.NO_WORK_TYPE_FIELD}","{MC.NO_KNOWN_TYPE}"),"")'),
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
    R.FLOW_TIME_GAP: ("=0", "=SUMPRODUCT(ABS(FlowTimeCheck[Gap]))"),
    R.FLOW_EFFICIENCY_GAP: ("=0", "=SUMPRODUCT(ABS(FlowEfficiencyCheck[Gap]))"),
    R.FLOW_LOAD_GAP: ("=0", "=SUMPRODUCT(ABS(FlowLoadCheck[Gap]))"),
    R.FLOW_DISTRIBUTION_GAP: ("=0", "=SUMPRODUCT(ABS(FlowDistCheck[Gap]))"),
    R.UNIQUE: ("=0", f'=COUNTIF({_LEDGER["key"]},"?*")-IFERROR(ROWS(UNIQUE(FILTER({_LEDGER["key"]},{_LEDGER["key"]}<>""))),0)'),
    R.CUTOFF: ("=0", f'=COUNTIF({_M}!$S$3#,"?*")'),
    R.CLASSIFIED: ("=0", f'=COUNTIFS({SI}[Category],"Unclassified",{SI}[Team],"<>{OTHER_TYPES}")'),
    R.FEATURE_TYPES: ("=1", f'=IF(COUNTIFS(ItemFlow[Role],"Program")=0,1,MIN(1,COUNTIFS(ItemFlow[Team],"{ART}")))'),
    R.START_STATUS: ("=0", "=" + "+".join(f'(COUNTIFS(ItemFlow[Workflow],"{k}",ItemFlow[Key],"?*")>0)*({_starts_of(chr(34) + k + chr(34))}=0)'
                                          for k in ("Feature", "Team"))),
    R.WAIT_STATUS: ("=1", f'=IF(COUNTIF(ItemFlow[Key],"?*")=0,1,MIN(1,COUNTIF({RULES_SHEET}!$C$2:$C$5000,"Wait")))'),
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
    # a date typed in Excel is a number; Load settings CSV writes an ISO text ('2026-12-01, '2026-12-01T00:00:00): DATEVALUE of its first 10
    # characters reads it, as DC18 and as the VBA (modFlowBuild.CheckPis refuses the rest)
    return (f'=IFERROR(LET(v,INDEX(PIs_from_user!${column}:${column},MATCH([@PI],PIs_from_user!$A:$A,0)),IF(ISNUMBER(v),IF(v>0,v,""),'
            f'DATEVALUE(LEFT(TRIM(v),10)))),"")')


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
# A final Z (UTC) and a non-breaking space (a value pasted from Confluence) are dropped, as modParams.ParseMoment and params/model._parse_moment do.
AS_OF_CELL = ('=LET(v,' + CONFIG_VALUE.format(name="AsOf") + ',IF(ISNUMBER(v),v,'
              'LET(t,SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(TRIM(SUBSTITUTE(v,CHAR(160)," ")),"Z",""),"z",""),"T"," "),'
              'IF(t="",NA(),DATEVALUE(LEFT(t,10))+IF(LEN(t)>10,TIMEVALUE(MID(t,12,8)),0)))))')
# TRIM keeps a non-breaking space: "A" pasted with one passed modFlowBuild.CheckModes and made every Flow Load #N/A (parity check, 2026-11-23)
MODE_CELL = '=UPPER(TRIM(SUBSTITUTE(CLEAN(' + CONFIG_VALUE.format(name="FlowLoadMode") + '),CHAR(160)," ")))'
PREDICTABILITY_MODE_CELL = '=UPPER(TRIM(SUBSTITUTE(CLEAN(' + CONFIG_VALUE.format(name="PredictabilityMode") + '),CHAR(160)," ")))'
TRANSITIONS_LAST_CELL, PIS_LAST_CELL = last_row_cell("Transitions"), last_row_cell("PIs")

# the statuses that no rule classifies, by name (spilled formulas right of the census, from column O): the intervals of the counted tickets (what the
# control counts) and of the tickets left out
_UNK = "$O$2#"
UNCLASSIFIED_COLUMNS = (
    ("Status", f'=IFERROR(SORT(UNIQUE(FILTER({SI}[Status],{SI}[Category]="Unclassified"))),"")'),
    ("IntervalsOfCounted", f'=IFERROR(COUNTIFS({SI}[Status],{_UNK},{SI}[Category],"Unclassified",{SI}[Team],"<>{OTHER_TYPES}"),"")'),
    ("IntervalsOfLeftOut", f'=IFERROR(COUNTIFS({SI}[Status],{_UNK},{SI}[Category],"Unclassified",{SI}[Team],"{OTHER_TYPES}"),"")'),
)
# the statuses that end the flow (IsEnd = 1 in the rules) and the tickets that sit in each now (spilled from column S): the person declares as exceptions the Done
# statuses of Jira that are no delivery
_ENDED = "$S$2#"
DONE_STATUS_COLUMNS = (
    ("EndStatus", f'=IFERROR(SORT(UNIQUE(FILTER({RULES_SHEET}!$B$2:$B$5000,{RULES_SHEET}!$E$2:$E$5000=1,""))),"")'),
    ("TicketsNow", f'=IFERROR(COUNTIFS(Issues[Status],{_ENDED}),"")'),
)
# the parameters in force: what the figures read, written next to them so that the person can compare it with Jira
PARAMETERS_IN_FORCE = (
    ("FeatureTypes (the types of the ART)", "=" + CONFIG_VALUE.format(name="FeatureTypes")),
    ("Program project(s)", '=TEXTJOIN(", ",TRUE,FILTER(Projects_from_user!$A$2:$A$300,Projects_from_user!$B$2:$B$300="Program",""))'),
    ("Statuses that start the flow (IsStart)", f'=TEXTJOIN(", ",TRUE,UNIQUE(FILTER({RULES_SHEET}!$B$2:$B$5000,{RULES_SHEET}!$D$2:$D$5000=1,"")))'),
    ("Statuses that end it (IsEnd)", f'=TEXTJOIN(", ",TRUE,UNIQUE(FILTER({RULES_SHEET}!$B$2:$B$5000,{RULES_SHEET}!$E$2:$E$5000=1,"")))'),
    ("Statuses declared Wait or Removed", f'=TEXTJOIN(", ",TRUE,UNIQUE(FILTER({RULES_SHEET}!$B$2:$B$5000,'
                                         f'({RULES_SHEET}!$C$2:$C$5000="Wait")+({RULES_SHEET}!$C$2:$C$5000="Removed"),"")))'),
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
    "FlowTimeCheck": FLOWTIMECHECK_KEYS + tuple(name for name, _ in FLOWTIMECHECK_COLUMNS),
    "FlowEfficiencyCheck": FLOWEFFICIENCYCHECK_KEYS + tuple(name for name, _ in FLOWEFFICIENCYCHECK_COLUMNS),
    "FlowLoadCheck": FLOWLOADCHECK_KEYS + tuple(name for name, _ in FLOWLOADCHECK_COLUMNS),
    "FlowDistCheck": FLOWDISTRIBUTIONCHECK_KEYS + tuple(name for name, _ in FLOWDISTRIBUTIONCHECK_COLUMNS),
}
