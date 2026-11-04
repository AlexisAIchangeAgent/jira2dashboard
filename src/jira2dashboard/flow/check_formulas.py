"""The Excel formulas of the data checks, as text (pure). The reference is flow/checks.py; the scenario of flow_excel.feature
compares both, issue by issue.

ItemFlow gets the raw fields of Issues (same row), helper columns and one 0/1 column per check. A check that reads a
Jira field is 0 for everybody while the id of that field is empty in Config (`_ready`), and the sheet DataChecks then shows
"not computable". Lists of types and field ids are read from Config, never written here.
"""
from jira2dashboard.flow.checks import FILL_FIELDS, PER_ISSUE
from jira2dashboard.flow.refs import HALF, IF_, IS, SI, row_index

# the raw fields of Issues copied into ItemFlow (an empty cell stays "")
RAW_FIELDS = ("Parent", "StatusCategory", "ResolvedUtc", "FixVersions", "StoryPoints", "PI", "BV", "TeamValue",
              "SprintState", "SprintStart", "SprintEnd", "SprintCount", "TargetStart", "TargetEnd", "KeyChanges")


def cfg(name: str) -> str:
    """The text of a Config value ("" when empty)."""
    return f'TRIM(INDEX(Config!$B:$B,MATCH("{name}",Config!$A:$A,0)))'


def _list(name: str) -> str:
    """Config list "a, b,c" as ",a,b,c," (no space around the commas): to look a value up with FIND (case-sensitive)."""
    raw = f'SUBSTITUTE({cfg(name)},CHAR(160)," ")'                     # a pasted list may hold a non-breaking space
    return f'","&TEXTJOIN(",",TRUE,TRIM(TEXTSPLIT({raw},",")))&","'


def in_list(value: str, name: str) -> str:
    return f'ISNUMBER(FIND(","&{value}&",",{_list(name)}))'


def _ready(*needs: str) -> str:
    """True when every field id (Config parameter) is filled in; "PIs" = the sheet PIs has rows."""
    parts = ["COUNTA(PIs!$A:$A)>1" if n == "PIs" else f'{cfg(n)}<>""' for n in needs]
    return parts[0] if len(parts) == 1 else "AND(" + ",".join(parts) + ")"


def _flag(rule: str, *needs: str) -> str:
    """=IF(rule,1,0), or 0 while a field that the rule reads is not configured."""
    if needs:
        return f"=IF({_ready(*needs)},IF({rule},1,0),0)"
    return f"=IF({rule},1,0)"


def valid_pi(cell: str) -> str:
    """PI#XX-YYYY, XX and YYYY digits only (the case of PI# counts, like in the reference)."""
    return (f'AND(LEN({cell})=10,EXACT(LEFT({cell},3),"PI#"),MID({cell},6,1)="-",'
            f'SUMPRODUCT(--ISNUMBER(FIND(MID({cell},{{4,5,7,8,9,10}},1),"0123456789")))=6)')


def _copy(column: str) -> str:
    return f'=IF(INDEX({IS}[{column}],{row_index(IF_)})="","",INDEX({IS}[{column}],{row_index(IF_)}))'


def _lookup_parent(column: str) -> str:
    return f'=IF([@Parent]="","",IFERROR(INDEX({IF_}[{column}],MATCH([@Parent],{IF_}[Key],0)),""))'


COPIES = tuple((name, _copy(name)) for name in RAW_FIELDS)

HELPERS = (
    ("IsFeature", f"=IF({in_list('[@Type]', 'FeatureTypes')},1,0)"),
    ("IsParentType", f"=IF({in_list('[@Type]', 'ParentRequiredTypes')},1,0)"),
    ("IsSubTask", f"=IF({in_list('[@Type]', 'SubTaskTypes')},1,0)"),
    ("IsPointsType", f"=IF({in_list('[@Type]', 'PointsRequiredTypes')},1,0)"),
    ("ParentType", _lookup_parent("Type")),
    ("ParentStart", _lookup_parent("TargetStart")),
    ("ParentEnd", _lookup_parent("TargetEnd")),
    ("ChildCount", f"=COUNTIFS({IF_}[Parent],[@Key])"),
    ("CurrentCategory", f'=XLOOKUP(1,({SI}[Key]=[@Key])*({SI}[IsOpen]=TRUE),{SI}[Category],"")'),
    ("SubParentTypes", f'=IF([@IsSubTask]=1,IFERROR(ROWS(UNIQUE(FILTER({IF_}[ParentType],({IF_}[Team]=[@Team])*'
                       f'({IF_}[IsSubTask]=1)*({IF_}[ParentType]<>"")))),0),0)'),
)

_PI_START = f'INDEX(PIs!$B:$B,MATCH([@PI],PIs!$A:$A,0))'
_PI_END = f'INDEX(PIs!$C:$C,MATCH([@PI],PIs!$A:$A,0))'


def _date(cell: str) -> str:
    return f"LET(v,{cell},IF(ISNUMBER(v),v,DATEVALUE(v)))"


_PI_WINDOW = (f"LET(s,{_date(_PI_START)},e,{_date(_PI_END)},"
              f'COUNTIFS({IF_}[Parent],[@Key],{IF_}[SprintStart],"<"&(s-{HALF}))+'
              f'COUNTIFS({IF_}[Parent],[@Key],{IF_}[SprintEnd],">"&(e+{HALF})))')

FLAGS = {
    "DC01": _flag('AND([@Workflow]="Team",[@IsParentType]=1,[@Parent]="")'),
    "DC02": _flag(f'AND([@Workflow]="Team",[@IsParentType]=1,[@ParentType]<>"",NOT({in_list("[@ParentType]", "FeatureTypes")}))'),
    "DC03": _flag('AND([@IsSubTask]=1,[@ParentType]<>"",[@SubParentTypes]>1)'),
    "DC04": _flag('AND([@StatusCategory]="done",[@ResolvedUtc]="")'),
    "DC05": _flag('AND([@StatusCategory]<>"done",[@ResolvedUtc]<>"")'),
    "DC06": f'=IF(COUNTIFS({SI}[Key],[@Key],{SI}[Check],"*MISMATCH*")>0,1,0)',
    "DC07": f'=IF(COUNTIFS({SI}[Key],[@Key],{SI}[Check],"*BEFORE_CREATED*")>0,1,0)',
    "DC08": f'=IF(COUNTIFS({SI}[Key],[@Key],{SI}[Category],"Unclassified")>0,1,0)',
    "DC09": '=IF([@SkippedStart]="Y",1,0)',
    "DC10": '=IF([@Reopened]>0,1,0)',
    "DC11": _flag('AND([@IsPointsType]=1,OR([@SprintState]="active",[@SprintState]="future"),[@StoryPoints]="")',
                  "StoryPointsField", "SprintField"),
    "DC12": _flag('AND([@IsSubTask]=1,[@StoryPoints]<>"")', "StoryPointsField"),
    "DC13": _flag('AND(OR([@IsFeature]=1,[@IsParentType]=1),[@TeamValue]="")', "TeamField"),
    "DC14": _flag('AND([@Workflow]="Team",[@TeamValue]<>"",[@TeamValue]<>[@Team])', "TeamField"),
    "DC15": _flag('AND([@Workflow]="Team",[@StatusCategory]="indeterminate",[@SprintCount]=0)', "SprintField"),
    "DC16": _flag("[@SprintCount]>=3", "SprintField"),
    "DC17": _flag(f"AND([@IsFeature]=1,NOT({valid_pi('[@PI]')}))", "PIField"),
    "DC18": _flag(f"AND([@IsFeature]=1,{valid_pi('[@PI]')},COUNTIFS(PIs!$A:$A,[@PI])>0,IFERROR({_PI_WINDOW},0)>0)",
                  "PIField", "SprintField", "PIs"),
    "DC19": _flag('AND([@IsFeature]=1,[@CurrentCategory]<>"Backlog",[@CurrentCategory]<>"Removed",'
                  'OR([@TargetStart]="",[@TargetEnd]=""))', "TargetStartField", "TargetEndField"),
    "DC20": _flag('OR(AND([@TargetStart]<>"",[@TargetEnd]<>"",[@TargetEnd]<[@TargetStart]),'
                  'AND([@ParentStart]<>"",[@TargetStart]<>"",[@TargetStart]<[@ParentStart]),'
                  'AND([@ParentEnd]<>"",[@TargetEnd]<>"",[@TargetEnd]>[@ParentEnd]))', "TargetStartField", "TargetEndField"),
    "DC21": _flag('AND([@IsFeature]=1,[@BV]<>"")', "BusinessValueField"),
    "DC22": _flag("COUNTIFS(IssueLinks[Blocked],[@Key],IssueLinks[Risk],1)>0", "SprintField"),
    "DC23": _flag("COUNTIFS(IssueLinks[Blocker],[@Key],IssueLinks[Cross],1,IssueLinks[BlockerNoSprint],1)+"
                  "COUNTIFS(IssueLinks[Blocked],[@Key],IssueLinks[Cross],1,IssueLinks[BlockedNoSprint],1)>0", "SprintField"),
    "DC24": _flag("AND([@IsFeature]=1,[@ChildCount]=0)"),
    "DC26": _flag('AND([@IsFeature]=1,[@FixVersions]="")'),
    "DC27": _flag("[@KeyChanges]>0"),
    "DC28": f"=IF(COUNTIFS({SI}[Key],[@Key],{SI}[BadMove],1)>0,1,0)",
    "DC29": _flag(f'AND([@PI]<>"",NOT({valid_pi("[@PI]")}))', "PIField"),
}
assert tuple(FLAGS) == PER_ISSUE, "one formula per issue-level check, in the order of the catalogue"

ITEMFLOW_CHECK_COLUMNS = COPIES + HELPERS + tuple(FLAGS.items())

# --- IssueLinks: one row per "blocks" link (the raw columns come from the extraction) ----------------------------------

LINKS_RAW = ("Blocker", "Blocked")


def _of(column: str, key: str) -> str:
    return f'IFERROR(INDEX({IF_}[{column}],MATCH({key},{IF_}[Key],0)),"")'


LINKS_COLUMNS = (
    ("Known", f"=IF(AND(ISNUMBER(MATCH([@Blocker],{IF_}[Key],0)),ISNUMBER(MATCH([@Blocked],{IF_}[Key],0))),1,0)"),
    ("BlockerEnd", "=" + _of("SprintEnd", "[@Blocker]")),
    ("BlockedStart", "=" + _of("SprintStart", "[@Blocked]")),
    ("BlockerTeam", "=" + _of("Team", "[@Blocker]")),
    ("BlockedTeam", "=" + _of("Team", "[@Blocked]")),
    ("Risk", f"=IF(AND([@Known]=1,ISNUMBER([@BlockerEnd]),ISNUMBER([@BlockedStart])),"
             f"IF([@BlockerEnd]>[@BlockedStart]+{HALF},1,0),0)"),
    ("Cross", '=IF(AND([@Known]=1,[@BlockerTeam]<>[@BlockedTeam]),1,0)'),
    ("BlockerNoSprint", f"=IF({_of('SprintCount', '[@Blocker]')}=0,1,0)"),
    ("BlockedNoSprint", f"=IF({_of('SprintCount', '[@Blocked]')}=0,1,0)"),
)

# --- FillRates: field x team x issue type (DC25); Field, Team and Type are keys written by VBA -------------------------

FILL_KEYS = ("Field", "Team", "Type")
FILL_COLUMNS = (
    ("Count", f"=COUNTIFS({IF_}[Team],[@Team],{IF_}[Type],[@Type])"),
    ("Filled", f"=IFERROR(SUMPRODUCT(({IF_}[Team]=[@Team])*({IF_}[Type]=[@Type])*"
               f"(LEN(INDEX({IF_},0,MATCH([@Field],{IF_}[#Headers],0)))>0)),0)"),
    ("Rate", '=IF([@Count]=0,"",[@Filled]/[@Count])'),
    ("Below", f"=IF(AND([@Count]>0,[@Filled]*100<[@Count]*VALUE({cfg('FillRateThreshold')})),1,0)"),
)
FILL_FIELD_NEEDS = FILL_FIELDS            # (ItemFlow column, Config parameter that must be set or "")

# --- DataChecks: one row per check (the keys come from the sheet ChecksSpec) ----------------------------------------------

DATACHECKS_KEYS = ("Id", "Title", "Severity", "Need1", "Need2", "Need3", "Source")


def _need_ok(cell: str) -> str:
    return f'IF({cell}="",TRUE,IF({cell}="PIs",COUNTA(PIs!$A:$A)>1,TRIM(INDEX(Config!$B:$B,MATCH({cell},Config!$A:$A,0)))<>""))'


DATACHECKS_COLUMNS = (
    ("Ready", "=IF(AND(" + ",".join(_need_ok(f"[@Need{n}]") for n in (1, 2, 3)) + "),1,0)"),
    ("Items", f'=IF([@Ready]=0,"",IF([@Source]="FillRates",SUM(FillRates[Below]),'
              f"SUM(INDEX({IF_},0,MATCH([@Id],{IF_}[#Headers],0)))))"),
    ("Result", '=IF([@Ready]=0,"not computable",IF([@Items]=0,"OK","KO"))'),
)
