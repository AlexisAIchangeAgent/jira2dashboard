# Config parameters

Rows of the sheet `Config`: the connection to Jira, including the proxy. The application writes this table into
the workbook and the VBA reads the sheet (`modJira.JiraConfigFromSheet`). **Fill the values in the workbook**:
nothing from the client may be kept in this repo, so `BaseUrl`, `Email`, `ProxyServer` and `ProxyBypass` stay
empty here, and **`ApiToken` must stay empty** (a test refuses a token in this file).

**Mandatory** says whether the macros work while the value is empty:
- `True`: the button **Extract from Jira and compute** stops before any call to Jira and names the empty ones. The
  button **Read the Jira configuration** checks only `BaseUrl`, `Email` and `ApiToken`. Without the value, a request
  would fail or a figure or a data check would be wrong.
- `False`: empty is allowed. A default is used, nothing is added (no filter, no proxy), or the data checks that need
  the value show "not computable".

| Parameter | Value | Mandatory | Description |
|---|---|---|---|
| BaseUrl |  | True | Jira site, for example https://your-site.atlassian.net (no trailing slash needed) |
| Email |  | True | Account e-mail (Basic authentication with the API token) |
| ApiToken |  | True | Type the API token in the workbook only. Spaces and line breaks around it are removed |
| Filter |  | False | Extra JQL restriction added with AND to the project list (sheet Projects), for example updated >= -90d. Empty = none |
| WorkTypeField |  | False | Id of the custom field that holds the work type (TypeCategorie: Feature, Enabler, Defect, Autre), for example customfield_10200. Find it with GET /rest/api/3/field. Empty = every item counts as Unmapped in Flow Distribution |
| ProxyServer |  | False | Proxy as host:port or http://host:port. Empty = the proxy of the Windows settings, found as the browser finds it (automatic detection WPAD, proxy script PAC, then the fixed proxy). DIRECT = no proxy |
| ProxyBypass |  | False | Hosts that skip the proxy, for example *.client.local;<local>. Empty with an empty ProxyServer = the list of the Windows settings |
| ProxyWindowsAuth | Yes | False | Yes = answer the proxy's Windows authentication with the Windows session. No = never send it |
| FlowLoadMode | A | True | A = a ticket counts from its first start until it is done or removed, except while in a Backlog status. B = only from its last start (after the last return to backlog) |
| AsOf |  | True | UTC date and time of the extraction, for example 2026-11-03 18:00. Open statuses run up to it. Required to extract: the macro stops with a message if it is empty |
| TimeoutSeconds | 60 | False | Connect, send and receive time limit of one request, in seconds |
| MaxRetries | 4 | False | Retries after a 429 or a server error (5xx), so 5 attempts at most |
| ParentRequiredTypes | Story,Task | False | Issue types of the team projects that must have a parent (data check DC01), separated by commas |
| FeatureTypes | Feature | True | Issue types that are features (data checks DC02, DC13, DC17-DC21, DC24, DC26), separated by commas |
| SubTaskTypes | Sub-task | False | Issue types that are sub-tasks (data checks DC03, DC12), separated by commas |
| PointsRequiredTypes | Story | False | Issue types that must have story points in an active or future sprint (data check DC11), separated by commas |
| StoryPointsField |  | False | Id of the story points field, for example customfield_10016 (data checks DC11, DC12). Empty = those checks are not computable |
| BusinessValueField |  | False | Id of the Business Value field (data check DC21). Empty = not computable |
| PIField |  | False | Id of the PI field holding PI#XX-YYYY (data checks DC17, DC18, DC29). Empty = not computable |
| TeamField |  | False | Id of the Team field (data checks DC13, DC14). Empty = not computable |
| SprintField |  | False | Id of the Sprint field (data checks DC11, DC15, DC16, DC18, DC22, DC23). Empty = not computable |
| TargetStartField |  | False | Id of the Target start date field (data checks DC19, DC20). Empty = not computable |
| TargetEndField |  | False | Id of the Target end date field (data checks DC19, DC20). Empty = not computable |
| BlocksLinkType | Blocks | True | Name of the issue link type that means blocks (data checks DC22, DC23) |
| FillRateThreshold | 90 | True | Minimum percentage of filled values per field, team and issue type (data check DC25) |
| LogLevel | Steps | False | Detail of the local log of the runs (%LOCALAPPDATA%\jira2dashboard\logs): Steps = each step (BEGIN/END), Errors = only the start, the end and the errors |
