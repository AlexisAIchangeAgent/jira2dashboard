# Config parameters

Rows of the sheet `Config`: the connection to Jira, including the proxy. The application writes this table into
the workbook and the VBA reads the sheet (`modJira.JiraConfigFromSheet`). **Fill the values in the workbook**:
nothing from the client may be kept in this repo, so `BaseUrl`, `Email`, `ProxyServer` and `ProxyBypass` stay
empty here, and **`ApiToken` must stay empty** (a test refuses a token in this file).

**Example** shows the expected form of the value. It is never read: only **Value** counts.

**Mandatory** says whether the macros work while the value is empty:
- `True`: the button **Extract from Jira and compute** stops before any call to Jira and names the empty ones. The
  button **Read the Jira configuration** checks only `BaseUrl`, `Email` and `ApiToken`. Without the value, a request
  would fail or a figure or a data check would be wrong.
- `False`: empty is allowed. A default is used, nothing is added (no filter, no proxy), or the data checks that need
  the value show "not computable".

| Parameter | Value | Example | Mandatory | Description |
|---|---|---|---|---|
| BaseUrl |  | https://your-site.atlassian.net | True | Jira site, for example https://your-site.atlassian.net (no trailing slash needed) |
| Email |  | first.last@company.com | True | Account e-mail (Basic authentication with the API token) |
| ApiToken |  | the token created at id.atlassian.com > Security > API tokens | True | Type the API token in the workbook only. Spaces and line breaks around it are removed |
| Filter |  | updated >= -90d | False | Extra JQL restriction added with AND to the project list (sheet Projects), for example updated >= -90d. Empty = none |
| WorkTypeField |  | customfield_10200 | False | Id of the custom field that holds the work type (TypeCategorie: Feature, Enabler, Defect, Autre), for example customfield_10200. Find it with GET /rest/api/3/field. Empty = every item counts as Unmapped in Flow Distribution |
| ProxyServer |  | proxy.company.com:8080, or DIRECT | False | Proxy as host:port or http://host:port. Empty = the proxy of the Windows settings, found as the browser finds it (automatic detection WPAD, proxy script PAC, then the fixed proxy). DIRECT = no proxy |
| ProxyBypass |  | *.company.local;<local> | False | Hosts that skip the proxy, for example *.client.local;<local>. Empty with an empty ProxyServer = the list of the Windows settings |
| ProxyWindowsAuth | Yes | Yes | False | Yes = answer the proxy's Windows authentication with the Windows session. No = never send it |
| IsUserProjectAdmin | True | True | False | True = you are an administrator of the Jira projects: the button Propose values from Jira reads the full workflows (statuses, transitions, initial status). False = it reads only the statuses and their categories (the right to browse the projects is enough). Anything but True, Yes or 1 counts as False |
| FlowLoadMode | A | A | True | A = a ticket counts from its first start until it is done or removed, except while in a Backlog status. B = only from its last start (after the last return to backlog) |
| AsOf |  | 2026-11-03 18:00 | True | UTC date and time of the extraction, for example 2026-11-03 18:00. Open statuses run up to it. Required to extract: the macro stops with a message if it is empty |
| TimeoutSeconds | 60 | 60 | False | Connect, send and receive time limit of one request, in seconds |
| MaxRetries | 4 | 4 | False | Retries after a 429 or a server error (5xx), so 5 attempts at most |
| ParentRequiredTypes | Story,Task | Story,Task | False | Issue types of the team projects that must have a parent (data check "Story or task without parent"), separated by commas |
| FeatureTypes | Feature | Feature | True | Issue types that are features (data checks "Parent is not a feature (wrong level)"; "No Team value"; "Feature without PI value"; "Feature PI inconsistent with its stories' sprints"; "Feature without fixed dates (Target start / Target end)"; "Target end before Target start, or child dates outside parent dates"; "Business value on a feature"; "Feature without children"; "Missing fix version on features"), separated by commas |
| SubTaskTypes | Sub-task | Sub-task | False | Issue types that are sub-tasks (data checks "Sub-tasks of a team have parents of different types"; "Story points on sub-tasks"), separated by commas |
| PointsRequiredTypes | Story | Story | False | Issue types that must have story points in an active or future sprint (data check "Story without story points in an active or future sprint"), separated by commas |
| StoryPointsField |  | customfield_10016 | False | Id of the story points field, for example customfield_10016 (data checks "Story without story points in an active or future sprint"; "Story points on sub-tasks"). Empty = those checks are not computable |
| BusinessValueField |  | customfield_10301 | False | Id of the Business Value field (data check "Business value on a feature"). Empty = not computable |
| PIField |  | customfield_10300 | False | Id of the PI field holding PI#XX-YYYY (data checks "Feature without PI value"; "Feature PI inconsistent with its stories' sprints"; "PI field not empty but not in the format PI#XX-YYYY"). Empty = not computable |
| TeamField |  | customfield_10400 | False | Id of the Team field (data checks "No Team value"; "Team value does not match the team project"). Empty = not computable |
| SprintField |  | customfield_10020 | False | Id of the Sprint field (data checks "Story without story points in an active or future sprint"; "In progress without sprint"; "Carried over 2+ sprints"; "Feature PI inconsistent with its stories' sprints"; "Dependency risk: blocker ends after the blocked item starts"; "Cross-team dependency without both items planned"). Empty = not computable |
| TargetStartField |  | customfield_10401 | False | Id of the Target start date field (data checks "Feature without fixed dates (Target start / Target end)"; "Target end before Target start, or child dates outside parent dates"). Empty = not computable |
| TargetEndField |  | customfield_10402 | False | Id of the Target end date field (data checks "Feature without fixed dates (Target start / Target end)"; "Target end before Target start, or child dates outside parent dates"). Empty = not computable |
| BlocksLinkType | Blocks | Blocks | True | Name of the issue link type that means blocks (data checks "Dependency risk: blocker ends after the blocked item starts"; "Cross-team dependency without both items planned") |
| FillRateThreshold | 90 | 90 | True | Minimum percentage of filled values per field, team and issue type (data check "Field fill rate below threshold") |
| LogLevel | Steps | Steps | False | Detail of the local log of the runs (the folder logs next to the workbook): Steps = each step (BEGIN/END), Errors = only the start, the end and the errors |
| OpenLogOnError | Yes | Yes | False | After an error of a button, open the local log in Notepad (Yes or No). The status of the Dashboard names the log file in both cases |
