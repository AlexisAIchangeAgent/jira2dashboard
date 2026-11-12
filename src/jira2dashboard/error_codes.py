"""Support codes of the errors: what a person reads out to the support, and what the support looks up here.

A code reads `J2D-<AREA>-<NNN>`. It is never reused and never renumbered: a code that stops being raised stays in the table.

| AREA | Covers |
|------|--------|
| CFG  | the parameters: Config, Projects, Iterations, a value the run cannot use |
| JIRA | what Jira answered. 4xx and 5xx codes carry the HTTP status itself (`J2D-JIRA-401`); 0xx are ours |
| NET  | no answer at all (connection, proxy, certificate, time out). 12xxx codes carry the WinHTTP error |
| DATA | an answer that cannot be read (JSON) |
| FLOW | building the tables, the formulas and the charts |
| ROVO | the Rovo prompt and the Rovo CSV |
| SYS  | Excel and the run itself (cancel, a run already going), the machine (`diagnose`, `check-machine`) |
| SCR  | the scripts of `scripts\\` (setup, run, release, manual-test, start-mock, scenario) |
| VBA  | an error of VBA or Excel that this tool does not know: `J2D-VBA-<its number>` (hexadecimal when negative) |
| GEN  | an error raised by this tool that has no code yet: `J2D-GEN-<its number>` (a gap to fill in this table) |

The VBA side is `modLog.ErrorCode`; `tests/unit/test_error_codes.py` compares it with this table. The sheet `ErrorCodes` of the
delivered workbook shows the same table (written at build by `flow/sheets.py`), the VBA log and the sheet `Steps` carry the code,
and so do `logs\\python.log` and the status of the Dashboard (`Stopped`, then the row `Code: J2D-CFG-002`).
"""
from __future__ import annotations

import re
from typing import NamedTuple

CODE_FORMAT = re.compile(r"^J2D-(CFG|JIRA|NET|DATA|FLOW|ROVO|SYS|SCR|VBA|GEN)-[0-9A-F]{2,8}$")


class ErrorCode(NamedTuple):
    code: str
    vba: int | None            # the number raised as vbObjectError + vba; None when the code has no such number
    meaning: str
    cause: str                 # what usually lies behind it
    check: str                 # what the person at the keyboard checks first
    support: str               # what the support asks for next


TABLE: tuple[ErrorCode, ...] = (
    # --- CFG: parameters -------------------------------------------------------------------------------------------------
    ErrorCode("J2D-CFG-001", 600, "No project to extract", "The sheet Projects has no row", "Type the Jira project keys in Projects (Project, Role, Team)", "A screenshot of Projects (keys may be hidden)"),
    ErrorCode("J2D-CFG-002", 608, "A mandatory parameter is empty", "A Config row with Mandatory = True has no Value", "Fill the parameters named in the message; the buttons stop before any call to Jira", "The names listed in the message"),
    ErrorCode("J2D-CFG-003", 603, "The Filter contains ORDER BY", "The Filter of Config is added to the JQL, which already sorts", "Remove ORDER BY from the Filter", "The Filter with its values hidden"),
    ErrorCode("J2D-CFG-004", 604, "AsOf is missing or not a date", "AsOf in Config is empty or typed as text", "Type a UTC date and time such as 2026-11-03 18:00", "The format of the cell (date or text)"),
    ErrorCode("J2D-CFG-005", 602, "BaseUrl does not start with https://", "The API token would travel in clear text, the run refuses", "Fix BaseUrl in Config (https://<site>.atlassian.net)", "None: the check is local"),
    ErrorCode("J2D-CFG-006", 620, "The Jira configuration has no project to read", "The sheet Projects is empty when the button Read the Jira configuration runs", "Type the Jira project keys in Projects", "A screenshot of Projects"),
    ErrorCode("J2D-CFG-010", 701, "An iteration has no valid date", "Start or End of the sheet Iterations is not a date", "Type real dates in Iterations (Start, End)", "The iteration label named in the message"),
    ErrorCode("J2D-CFG-011", 704, "An iteration ends before it starts", "End is earlier than Start in Iterations", "Swap or correct the dates", "The iteration label named in the message"),
    ErrorCode("J2D-CFG-020", None, "The parameter sheets are not as expected", "A sheet or a column is missing, or a value has the wrong shape (the message lists each problem)", "Read each line of the message and fix the sheet named", "The message (it names sheets and columns, never values)"),
    ErrorCode("J2D-CFG-012", 703, "No iteration", "The sheet Iterations is empty", "Fill Iteration-label, Start and End", "A screenshot of Iterations"),
    # --- JIRA: what Jira answered ----------------------------------------------------------------------------------------
    ErrorCode("J2D-JIRA-010", 601, "The search did not end", "More search pages than the safety limit", "Narrow the extraction with the Filter of Config", "The log lines END search page n"),
    ErrorCode("J2D-JIRA-011", 606, "Jira keeps sending the same page token", "A Jira paging fault", "Try again later; if it repeats, narrow the Filter", "The log (the page numbers) and the time of the run"),
    ErrorCode("J2D-JIRA-012", 505, "Jira keeps sending the same changelog page", "A Jira paging fault on the changelog call", "Try again later", "The log (the issue count of the last chunk)"),
    ErrorCode("J2D-JIRA-013", 621, "The project list did not end", "More project pages than the safety limit", "Narrow the Projects sheet", "The log"),
    ErrorCode("J2D-JIRA-020", 607, "Jira answered 200 with something that is not JSON", "A proxy or a login page answers instead of Jira", "Check ProxyServer and BaseUrl in Config; open BaseUrl in the browser on the same machine", "The start of the answer (the message shows 200 characters: hide anything private)"),
    ErrorCode("J2D-JIRA-400", None, "HTTP 400: Jira refused the request", "A JQL or parameter Jira cannot read", "Read the message after the code; check the Filter and the field names", "The message (without private values)"),
    ErrorCode("J2D-JIRA-401", None, "HTTP 401: Jira refused the account", "Wrong Email or ApiToken, or the token expired", "Check Email and ApiToken in Config; create a new API token if needed", "The date of the token and the Email used (the token itself is never sent)"),
    ErrorCode("J2D-JIRA-403", None, "HTTP 403: not allowed", "The account lacks a right (Browse projects; Create issues for the field list by project)", "Ask a Jira administrator for the right named in the message", "The project key and the right"),
    ErrorCode("J2D-JIRA-404", None, "HTTP 404: not found", "A wrong BaseUrl, a project that does not exist or is not visible", "Check BaseUrl and the project keys", "BaseUrl without the secret parts, the project key"),
    ErrorCode("J2D-JIRA-407", None, "HTTP 407: the proxy asks for authentication", "The corporate proxy wants the Windows session or credentials", "Check ProxyServer and ProxyWindowsAuth in Config", "Whether other programs reach Jira through the proxy"),
    ErrorCode("J2D-JIRA-429", None, "HTTP 429: too many requests", "Jira slows the account down; the run already retried", "Wait a few minutes and run again", "The time of the run"),
    ErrorCode("J2D-JIRA-500", None, "HTTP 5xx: Jira or the proxy failed (500, 502, 503, 504)", "A server side fault; the run already retried", "Try again later; look at the Atlassian status page", "The exact status in the code and the time"),
    # --- NET: no answer ---------------------------------------------------------------------------------------------------
    ErrorCode("J2D-NET-001", None, "No answer from Jira", "The connection failed with no WinHTTP error number", "Check the network, BaseUrl, ProxyServer", "The log of the run"),
    ErrorCode("J2D-NET-12002", None, "The request timed out", "A slow proxy or network", "Run again; check the proxy", "The log, the time of the run"),
    ErrorCode("J2D-NET-12007", None, "The server name could not be resolved", "A wrong BaseUrl host, or the DNS does not answer", "Check BaseUrl (the site name); try it in the browser", "BaseUrl without the secret parts"),
    ErrorCode("J2D-NET-12029", None, "Cannot connect to the server", "A firewall, a proxy that is down, or a wrong proxy address", "Check ProxyServer; ask the network team about the proxy", "The proxy address and whether the browser reaches Jira"),
    ErrorCode("J2D-NET-12030", None, "The connection was closed", "A proxy or an endpoint protection cut the connection", "Run again; ask the network team if it repeats", "The log, the time of the run"),
    ErrorCode("J2D-NET-12038", None, "The certificate name does not match", "An inspecting proxy with its own certificate", "Ask the network team to trust the proxy certificate on this machine", "The proxy product name"),
    ErrorCode("J2D-NET-12045", None, "The certificate authority is not trusted", "An inspecting proxy whose authority is not installed", "Ask the network team to install the authority", "The proxy product name"),
    ErrorCode("J2D-NET-12175", None, "A secure channel failure", "TLS blocked or downgraded by a proxy", "Ask the network team", "The proxy product name and the Windows version"),
    # --- DATA ---------------------------------------------------------------------------------------------------------------
    ErrorCode("J2D-DATA-001", 513, "The JSON cannot be read", "An answer cut or altered on the way (the text is not logged: it quotes Jira)", "Run again; check the proxy", "The log (the position of the error)"),
    ErrorCode("J2D-DATA-002", 514, "A value cannot be written as JSON", "An internal case", "None", "The log"),
    # --- FLOW ---------------------------------------------------------------------------------------------------------------
    ErrorCode("J2D-FLOW-001", 702, "A raw sheet is empty", "The extraction did not run or returned nothing", "Run Extract from Jira first", "The log of the run"),
    ErrorCode("J2D-FLOW-002", 705, "A sheet of the workbook is missing", "A sheet was deleted or renamed", "Use a fresh copy of the delivered workbook", "The workbook version"),
    ErrorCode("J2D-FLOW-003", 706, "A raw column differs from the expected one", "The sheet Issues or StatusIntervals was edited", "Run the extraction again", "A screenshot of the header row"),
    ErrorCode("J2D-FLOW-004", 707, "Excel refused a formula", "A column name or a Config value breaks a formula", "Read the table and column in the message", "The message and the Excel version"),
    ErrorCode("J2D-FLOW-005", 708, "Excel is too old", "The workbook needs LET, XLOOKUP and FILTER (Excel 365 or 2021)", "Use Excel 365 or 2021", "File > Account > About Excel"),
    ErrorCode("J2D-FLOW-006", 709, "The sheet Issues has no column Type or PI", "The sheet was edited", "Run the extraction again", "A screenshot of the header row"),
    # --- ROVO ---------------------------------------------------------------------------------------------------------------
    ErrorCode("J2D-ROVO-001", 630, "The Rovo prompt spec is empty", "The sheet RovoPromptSpec was cleared", "Use a fresh copy of the delivered workbook", "The workbook version"),
    ErrorCode("J2D-ROVO-002", 631, "The Rovo CSV file does not exist", "A wrong path or a file removed", "Click Load settings CSV again and choose the file in the dialog", "The folder content (names only)"),
    ErrorCode("J2D-ROVO-003", 632, "Retired on 2026-11-11: the Rovo CSV had to be in the folder of the workbook", "A workbook built before that date", "Use a workbook built since: the CSV may be in any folder", "The version on the RUN line of the log"),
    # --- SYS ----------------------------------------------------------------------------------------------------------------
    ErrorCode("J2D-SYS-001", 610, "Cancelled by the user", "The Cancel button was pressed", "Nothing: the results of the previous run are kept", "None"),
    ErrorCode("J2D-SYS-002", 611, "An extraction is already running", "A second click or a second macro during a run", "Wait for the run to end or cancel it", "None"),
    ErrorCode("J2D-SYS-010", None, "The results file is open in another program", "The .xlsx of the Python extraction is open in Excel", "Close the file and run again", "None"),
    ErrorCode("J2D-SYS-011", None, "The workbook cannot be read", "The path is wrong, or the file is not an .xlsx/.xlsm workbook, or is damaged", "Check the --workbook path; open the file in Excel", "The file size and where it came from"),
    ErrorCode("J2D-SYS-020", None, "diagnose: no workbook found", "No .xlsm next to the scripts and none given", "Build one (scripts\\run.ps1 build) or give one with --workbook", "None"),
    ErrorCode("J2D-SYS-021", None, "The workbook has the Mark of the Web", "The file came from the Internet or a mail: Office opens it in Protected View and blocks the macros", "Right-click > Properties > Unblock, or put it in a Trusted Location", "The ZoneId in the report"),
    ErrorCode("J2D-SYS-022", None, "The workbook is not in a Trusted Location of Excel", "Excel asks to enable macros for a file outside its trusted folders", "Put the file in a Trusted Location", "The number of trusted locations in the report"),
    ErrorCode("J2D-SYS-023", None, "A policy blocks the macros of files from the Internet", "BlockContentExecutionFromInternet = 1", "Unblock the file or use a Trusted Location; ask IT if it still blocks", "The policy value in the report"),
    ErrorCode("J2D-SYS-024", None, "Office disables unsigned macros", "VBAWarnings is 3 or 4 (user setting or policy)", "Use a Trusted Location, or ask IT to sign the macros or change the policy", "The VBAWarnings values in the report"),
    ErrorCode("J2D-SYS-025", None, "A security product other than Windows Defender is registered", "An endpoint protection can stop macros or the COM objects they create", "Ask the security team to read its alerts or to put a test machine in Audit mode", "The product names in the report"),
    ErrorCode("J2D-SYS-026", None, "No VBA log on this machine", "No macro has run here, or it was blocked before it started", "Open the workbook, enable macros and click a button", "None"),
    ErrorCode("J2D-SYS-027", None, "The workbook is not named jira2dashboard.xlsm", "An allow-list on the name or the path misses a renamed copy", "Rename the file to jira2dashboard.xlsm", "The file name in the report"),
    ErrorCode("J2D-SYS-028", None, "Nothing in Office or in the file explains a block", "The endpoint protection is the likely cause", "Run diagnose and ask the security team for the alert (class, process, rule)", "The report of diagnose"),
    ErrorCode("J2D-SYS-029", None, "The macros did not run even with Office's security relaxed", "Something outside Office stops them, most likely the endpoint protection", "Run diagnose and ask the security team for its Audit mode", "The report of diagnose and of check-machine"),
    ErrorCode("J2D-SYS-040", None, "A check of check-machine failed", "Excel, the VBA project access or a COM call failed on this machine", "Read the line FAIL: it names the check", "The output of check-machine"),
    ErrorCode("J2D-SYS-099", None, "Unexpected error in the Python extraction", "A case this tool does not know; the traceback is in python.log", "Send python.log", "python.log (it holds no token)"),
)

SCRIPT_ROWS = (
    ("J2D-SCR-001", "A script got an argument it does not know", "A typo in an option", "Run the script with --help", "The command line"),
    ("J2D-SCR-002", "run.ps1: unknown command", "A typo in the command", "Run .\\scripts\\run.ps1 --help for the list", "The command line"),
    ("J2D-SCR-003", "No .venv", "setup.ps1 was not run on this machine", "Run scripts\\setup.ps1 (-Dev for the tests)", "None"),
    ("J2D-SCR-004", "setup: the virtual environment was not created", "Python 3.14 is not installed or not found by py", "Install Python 3.14 (py -3.14 --version)", "The output of py -0"),
    ("J2D-SCR-005", "setup: pip could not upgrade itself", "No access to the package index (proxy)", "Run setup.ps1 -Proxy http://proxy:port", "setup.log"),
    ("J2D-SCR-006", "setup: pip could not install the requirements", "No access to the package index, or a wheel is missing", "Run setup.ps1 -Proxy http://proxy:port", "setup.log"),
    ("J2D-SCR-007", "setup: pywin32 cannot be imported", "pywin32 was not installed in .venv", "Run setup.ps1 again", "setup.log"),
    ("J2D-SCR-010", "release: the version is not vX.Y.Z", "A typo in -Version", "Use v1.0.0 as the form", "The value given"),
    ("J2D-SCR-011", "release: the working tree has changes", "Uncommitted files", "Commit or stash them first", "git status"),
    ("J2D-SCR-012", "release: gh is not logged in", "No GitHub session", "Run gh auth login", "None"),
    ("J2D-SCR-013", "release: git fetch failed", "No access to GitHub (network, proxy)", "Check the network, then run again", "release.log"),
    ("J2D-SCR-014", "release: HEAD is not origin/main", "A local commit is not pushed, or the branch is not main", "Push first", "git log origin/main..HEAD"),
    ("J2D-SCR-015", "release: the tag already exists on GitHub", "The version was published before", "Choose a new version number", "None"),
    ("J2D-SCR-016", "The build of the workbook failed", "Excel did not start, VBA project access is off, or the VBA does not compile", "Run run.ps1 build and read its output; run check-machine", "The output of the build"),
    ("J2D-SCR-017", "release: git archive failed", "git is missing or the repository is damaged", "Run git status", "release.log"),
    ("J2D-SCR-018", "release: the zip holds files it must leave out", ".gitattributes export-ignore lost a path", "Fix .gitattributes (test_release_contents.py pins the list)", "release.log"),
    ("J2D-SCR-019", "release: the zip misses a required file", "A required path was removed from the repository", "Restore the file", "release.log"),
    ("J2D-SCR-020", "release: gh release create failed", "No right on the repository or no network", "Check gh auth status", "release.log"),
    ("J2D-SCR-021", "manual-test: the copy of the workbook could not be prepared", "mock-workbook failed (Excel, a locked file)", "Close the copy in Excel and run again", "manual-test.log"),
    ("J2D-SCR-022", "manual-test: the fake Jira did not start", "The port is taken or Python cannot start", "Run start-mock.ps1 alone and read its message", "start-mock.log"),
    ("J2D-SCR-023", "manual-test: the Python extraction failed", "See the J2D code of python.log", "Read python.log for the J2D code", "python.log"),
    ("J2D-SCR-030", "start-mock: the port is held by another program", "Another program listens on the port", "Choose another one with -Port", "The port number"),
    ("J2D-SCR-031", "start-mock: the fake Jira stopped at once", "A fault when it starts", "Run run.ps1 mock-jira --port <port> to see why", "start-mock.log"),
    ("J2D-SCR-032", "start-mock: no answer on the port after 60 s", "The fake Jira is slow or blocked by a firewall or an endpoint protection", "Run run.ps1 mock-jira alone and watch it start", "start-mock.log"),
    ("J2D-SCR-040", "scenario: a step failed", "See the step marked FAILED in the summary", "Read scenario.log from the line of the failed step", "scenario.log"),
    ("J2D-SCR-041", "scenario: the workbook given does not exist", "A wrong -Workbook path", "Check the path", "The path"),
    ("J2D-SCR-042", "scenario: the extraction wrote no .xlsx", "The Python extraction failed silently", "Read python.log", "python.log"),
    ("J2D-SCR-043", "scenario: the .xlsx is empty", "The extraction stopped while writing", "Read python.log", "python.log"),
)

TABLE = TABLE + tuple(ErrorCode(code, None, meaning, cause, check, support) for code, meaning, cause, check, support in SCRIPT_ROWS)

BY_CODE: dict[str, ErrorCode] = {row.code: row for row in TABLE}
HEADER = ("Code", "Meaning", "Probable cause", "What to check", "What the support asks")


def rows() -> list[tuple[str, str, str, str, str]]:
    """The rows of the sheet ErrorCodes of the delivered workbook (Code | Meaning | Probable cause | What to check | Support asks)."""
    return [(row.code, row.meaning, row.cause, row.check, row.support) for row in TABLE]


def markdown() -> str:
    """The table as Markdown (python -m jira2dashboard.error_codes)."""
    lines = ["| " + " | ".join(HEADER) + " |", "|" + "---|" * len(HEADER)]
    lines += ["| " + " | ".join(cell.replace("|", "/") for cell in row) + " |" for row in rows()]
    return "\n".join(lines)


_BY_VBA = {row.vba: row.code for row in TABLE if row.vba is not None and row.vba != 602}


def code_of_vba(number: int, description: str = "") -> str:
    """The support code of an error raised as vbObjectError + number (the same rule as modLog.ErrorCode).

    602 has two causes: a description that starts with "HTTP <status>" is an answer of Jira (status 0: no answer, the WinHTTP error
    number follows "attempt(s). "), any other is the BaseUrl check."""
    if number == 602:
        status = re.match(r"HTTP (\d+)", description)
        if not status:
            return "J2D-CFG-005"
        if int(status.group(1)) > 0:
            return code_of_http(int(status.group(1)))
        net = re.search(r"attempt\(s\)\. (-?\d+)", description)
        return code_of_winhttp(int(net.group(1))) if net else "J2D-NET-001"
    return _BY_VBA.get(number, f"J2D-GEN-{number:03d}")


def code_of_http(status: int) -> str:
    """The code of an answer of Jira that is not 200: `J2D-JIRA-<status>` (status 0 = no answer: `J2D-NET-001`)."""
    return f"J2D-JIRA-{status:03d}" if status > 0 else "J2D-NET-001"


def code_of_winhttp(number: int) -> str:
    """The code of a WinHTTP error given as an HRESULT (0x80072EE7) or as 12007; `J2D-NET-001` when it is not one of WinHTTP."""
    low = number & 0xFFFF
    if (number == low or number & 0xFFFF0000 == 0x80070000) and 12000 <= low <= 12200:
        return f"J2D-NET-{low}"
    return "J2D-NET-001"


if __name__ == "__main__":
    print(markdown())
