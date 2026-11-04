# Offline reference: Jira Cloud REST API from Excel VBA

This file tells you **how the macro-enabled workbook talks to Jira Cloud**: endpoints, limits, auth, proxy,
pagination, rate limits, and the workbook layout that receives the data.
The VBA code that implements it is in `vba-modules-tested.md` (tested, 143 checks OK, 100 % coverage);
the Python tools to extract and test it are in `cookbook_python.md`.
Python generation rules are in `pywin32-excel-vba-reference.md`.

Labels:
- **Documented**: from the official Jira Cloud REST API v3 / Jira Software API specification, and the
  Anthropic Jira API plugin reference (both checked 2026-11-03).
- **Tested**: run in Excel 16.0 (French UI) on 2026-11-03, with fake Jira responses or httpbin.org.
- **Verify at client**: cannot be tested outside the client network (real Jira, real proxy).

---

## 1. Endpoints used

| Need | Endpoint | Key facts (Documented) |
|---|---|---|
| Issues with all fields | `GET /rest/api/3/search/jql` (or `POST` with a JSON body) | The **only** supported search (`/rest/api/3/search` and `/search/id` return **410 Gone**: never retry, never use). Default returns **ids only**: always pass `fields=*all`. `expand=names` adds field id -> display name. Pagination: pass `nextPageToken`; the token is **absent on the last page**; `isLast` = true. The JQL must be **bounded** (`project in (...)`), max 7 `ORDER BY` fields. Page size can be smaller than `maxResults` when many fields are requested: read what came back. **No total.** `*all` is slow. POST body: `{"jql","fields":[...],"maxResults","nextPageToken","expand":"names"}` (`expand` is a comma-separated **string**, not an array). Use POST when the JQL is long. |
| Number of issues (progress, reconciliation) | `POST /rest/api/3/search/approximate-count` | Body `{"jql":"..."}`, response `{"count":N}`. **Approximate**: use it for a progress bar and to warn when the extracted count differs, never as the exact total. |
| Full change history | `POST /rest/api/3/changelog/bulkfetch` | Body `{"issueIdsOrKeys":[...], "maxResults":1000, "nextPageToken":"..."}`. **Max 1000 issues per call** (400 error above), up to 10 `fieldIds` filters. Paginated: `nextPageToken` is null on the last page. Sorted from oldest change. The response gives **`issueId`, not the key**: keep an id -> key map from the search. |
| Do not use for history | `expand=changelog` | Returns at most **40** changes per issue. Not reliable for flow metrics. |
| Field catalogue | `GET /rest/api/3/field` | All system and custom fields visible to the user: `id`, `name`, `custom`, `schema.type`, `schema.custom`. Use it to find the real ids of Team, Sprint, Story points, Target start/end, PI, Business Value. |
| Sprints of a board | `GET /rest/agile/1.0/board/{boardId}/sprint` | `startAt` / `maxResults` / `isLast` pagination (offset style), `state=future,active,closed`. Gives sprint start, end and complete dates. |
| Statuses | `GET /rest/api/3/status`, `GET /rest/api/3/statuscategory` | Status name, id and `statusCategory.key` (`new`, `indeterminate`, `done`). |
| Board estimation field | `GET /rest/agile/1.0/board/{id}/configuration` | Column mapping and the **estimation field** of the board: tells which field is "story points" for each team (DC11, DC12). Find boards with `GET /rest/agile/1.0/board?projectKeyOrId=KEY` (offset pagination, `.values`). |
| Where a custom field exists | `GET /rest/api/3/field/{fieldId}/context` | Projects and issue types of each context. Distinguish "**empty**" (fillable, not filled) from "**not available**" in that project (DC25). |
| Link types | `GET /rest/api/3/issueLinkType` | Names, inward and outward labels (`Blocks` / `is blocked by`), for DC22 and Plans. |
| Workflows | `GET /rest/api/3/workflowscheme/project?projectId=...`, `GET /rest/api/3/workflows/search` | Which workflow each project uses, to fill the `Workflows` and `Transitions` sheets. Admin rights may be needed. |
| Sprint issues | do not use `GET /rest/agile/1.0/sprint/{id}/issue` | Agile issue-list endpoints are **deprecated**; read sprint membership from the issues (Sprint field) found by `/search/jql`. |

With 1200 issues in one project: search = 12 pages of 100; changelogs = 2 calls (1000 + 200) plus their
extra pages. `JiraChangelogsAll` does the chunking (Tested with 1200 fake keys: 1000 then 200).

## 2. Authentication (Documented)

- Jira Cloud basic auth: header `Authorization: Basic base64(email:api_token)`, UTF-8 encoded.
  `BasicAuthHeader(email, token)` builds it (Tested against Python `base64`, accents included).
- The token is typed by each user in the **Config** sheet. Each user keeps their own copy of the workbook.
  Rules: never write the token in code, logs, test results or file names; never send the workbook to someone
  else with the token inside; give the Config cell a clear warning label; clear it with a "Clear token" button.
- `401` = wrong email/token (the body can be plain text, not JSON). `403` = no permission.
  `404` on an issue = not visible to this user (Jira answers 404, not 403). `410` = removed endpoint.
- Error bodies are `{"errorMessages":[...],"errors":{"field":"reason"}}`. Show both in the `Log` sheet
  (`CheckStatus` keeps the first 300 characters of the body; Copilot may add a parser for this envelope).
- Users are identified by `accountId`. `emailAddress` is often hidden by privacy settings: never use it as a key.

## 3. Proxy (Verify at client)

The client already has a **working PowerShell extraction** that goes through the proxy.
Copilot must ask the developer to open that script **on the client machine** and copy its proxy settings.
Nothing leaves the client: Copilot reads it there. Do not copy the token from it.

Mapping from PowerShell to the VBA settings in `modHttp` (`NewHttpConfig` keys):

| In the PowerShell script | VBA setting | WinHttp call made by modHttp |
|---|---|---|
| `-Proxy http://host:port` | `cfg("ProxyServer") = "host:port"` | `SetProxy 2, "host:port", bypass` |
| `-ProxyUseDefaultCredentials` | `cfg("ProxyWindowsAuth") = True` | `SetAutoLogonPolicy 0` (always send Windows credentials) |
| A bypass list (`-NoProxy`, `ProxyBypass`) | `cfg("ProxyBypass") = "*.client.local;<local>"` | 3rd argument of `SetProxy` |
| `-ProxyCredential (Get-Credential)` | not supported as is | needs `SetCredentials user, pwd, 1` (1 = for proxy). Ask before adding: never store a password. |
| `[System.Net.WebRequest]::DefaultWebProxy` (system proxy) | leave `ProxyServer` empty and try first | WinHTTP does **not** read the browser proxy. If it fails, set the proxy explicitly. |

If calls fail with status 0 and an error text, show `resp("Error")` in the status bar and in a log sheet.
Common causes: wrong proxy, TLS inspection certificate not trusted by WinHTTP, proxy needs other credentials.

## 4. Rate limits and retries (Documented, retry logic Tested with fakes)

- Jira Cloud uses a **points-based** hourly budget plus per-second burst limits, shared by all requests
  to the site with the same credential. Running the extraction twice in parallel halves the speed.
- `X-RateLimit-NearLimit: true` = less than 20 % of a budget left: slow down **before** the 429 (for
  example wait 1 s between pages). `RateLimit-Reason` says which limit was hit.
- On `429`: wait **at least `Retry-After` seconds**; if absent, exponential backoff with full jitter
  (reference client: base 2 s, max 30 s, 4 retries). Retry `503` only when it has `Retry-After`.
- Retry only safe methods. Our two POST calls (search is GET; bulkfetch is read-only) are safe to retry.
  Never auto-retry a POST that creates or changes data.
- `HttpRequest` does all this. Defaults: `MaxRetries=4`, `RetryBaseMs=2000`, `RetryMaxMs=30000`.

## 5. Not blocking Excel: two options (both kept, choose at the client)

**Option A - VBA calls Jira itself (Tested with httpbin.org).**
`WinHttp` is opened in asynchronous mode; `WaitForResponse(0)` is polled with `DoEvents` + `Sleep 20`.
Excel repaints and the status bar shows progress (`Jira search: 300 issues read (page 3)`).
Limits: VBA still runs one thing at a time. While the macro runs, the user should not edit the workbook.
Protect against double clicks with a module flag (`If gBusy Then Exit Sub`). Set
`Application.Cursor = xlWait` during the run and `xlDefault` in the error handler and at the end.

**Option B - the user runs the PowerShell export, Excel watches the folder (Tested without a process).**
- **Tested: an antivirus blocked `powershell.exe` started from VBA** (Office starting PowerShell is a
  classic malware pattern; ASR rules block it). So Excel must **never** start PowerShell, `cmd`, `WScript.Shell`
  or any process. Do not generate such code.
- The user starts the existing script as today. The script writes `page-0001.json`, `page-0002.json` ... (each
  one = the raw `/search/jql` response) and `changelog-0001.json` ... then **`done.flag` last**.
- Button "Watch export folder" -> `WatchExportFolder(outDir)`: `Application.OnTime` checks every 5 s.
  Excel stays fully usable. When `done.flag` appears: status `DONE`, button "Import" -> `ImportSearchPages`.
- Ask the developer to add the `page-*.json` / `done.flag` output to the PowerShell script if it does not
  write them yet. Keep file names sorted (`0001`) so `ListFiles` reads pages in order.

## 6. Workbook layout (agreed with the developer)

| Sheet | Content | Written by |
|---|---|---|
| `Config` | Base URL, email, API token (warning label), proxy server / bypass / Windows auth, project list (key, team name, role `Team` or `Program`), parameter `FlowLoadMode` (`A` or `B`), field id mapping (logical name -> custom field id), export folder for option B, "as of" date | user |
| One sheet **per team** (`T_<team>`) | All issues of that team project, one row per issue, **all fields**: `Key`, `Id`, then one column per field. Header row 1 = display name, row 2 = field id. **Join tables on `Id`**, not on `Key`: a key changes when an issue moves to another project, the id never changes | Extract |
| `Workflows`, `Transitions` | Statuses with category, start, end, free flag, and allowed moves per workflow (`Team`, `Feature`); see `safe-flow-metrics-and-data-checks.md` section 1 | user (Config data) |
| `Features` | Issues of the Program project (features), same layout | Extract |
| `Changelog` | One row per changed item: IssueKey, IssueId, ChangeId, CreatedUtc, Author, AuthorId, FieldId, Field, FieldType, From, FromString, To, ToString | Extract |
| `StatusIntervals` | One row per status visit (see `modFlow`): Key, Seq, Status, Category, EnteredUtc, LeftUtc, DurationDays, IsOpen, ChangeId, Check | Compute |
| `FieldCatalog` | From `GET /rest/api/3/field`: id, name, custom, schema type, fill rate per team and issue type | Extract + Checks |
| `DataChecks` | One row per anomaly: CheckId, Severity, IssueKey, Team, Field, Found value, Expected, Rule, Fix hint | Checks |
| `MissingData` | Fields or objects missing for flow metrics and for Plans (see `safe-flow-metrics-and-data-checks.md`) | Checks |
| `RawFields` | Narrow raw values for the checks: Key, FieldId, JSON (cut at 32 767), only for the fields the checks need | Extract |
| `RovoCompare` | CheckId, Excel count, saved filter count, Rovo count, exact/estimated, difference, comment (counts typed by the user; see `rovo-guide.md` section 9) | Checks + user |
| `ItemFlow`, `DailyLoad`, `Metrics` | Intermediate and final flow metric tables, with formulas visible | Compute |
| `Tests_<module>` | Inputs and outputs of the tests of one module (fixtures visible to the user) | Tests |
| `TestResults` | TestId, Area, Description, Input, Expected, Actual, Result OK/KO, RunAt | Tests |
| `CoverageMap`, `CoverageHits` | Coverage markers, hits, and the coverage % (`E2`) | Generator + Tests |
| `Log` | One row per run: start, end, user, projects, issues, changes, HTTP calls, retries, errors | all macros |

Buttons (Form buttons, `add_macro_button` in the pywin32 reference; set captions in English or French as the
developer prefers):
- `Config`: **Extract from Jira** (option A), **Watch export folder** and **Import export files** (option B),
  **Clear token**.
- `DataChecks`: **Run data checks**. `Metrics`: **Compute flow metrics**.
- `TestResults`: **Run all tests** (offline), **Run network test** (optional), **Clear results**.

## 7. Extraction flow (button "Extract from Jira")

1. Read Config. Validate: URL starts with `https://` and ends with `.atlassian.net`; email and token not empty;
   at least one project. Stop with a clear message otherwise.
2. `cfg = JiraConfig(url, email, token, proxy, bypass, winAuth)`.
3. For each project: `issues = JiraSearchAll(cfg, ProjectsJql(oneProject), names)`.
   Build rows with `IssueFieldIds` + `IssueRow`; write with `WriteTable` (one block, fast).
   Keep `id -> key` and `key -> Array(createdUtc, currentStatus)` maps.
4. `rows = JiraChangelogsAll(cfg, allKeys)` (chunks of 1000). Map ids to keys. Write `Changelog`.
5. `StatusChangesFromChangelog` + `BuildStatusIntervals` -> `StatusIntervals` (with the Check column).
6. Write one `Log` row. Reset `Application.StatusBar = False` in all cases (use an error handler).

During writes: `Application.ScreenUpdating = False`, `Application.Calculation = xlCalculationManual`,
restore both in the error handler. Never use `Select` or `Activate`.

## 8. VBA traps found by testing (Tested, French Excel)

| Trap | What happened | Rule |
|---|---|---|
| `True & ""`, `CStr(True)`, `Str$(True)` | Gave **"Vrai"** | Never turn a Boolean into text. Use `IIf(b, "True", "False")` or write the Boolean to the cell. |
| `CDbl("1.5")`, `CStr(1.5)` | Use the Windows locale (comma) | Parse JSON numbers with `Val()`; write numbers to cells as numbers. |
| `On Error Resume Next` in a helper | `Err` stayed set after the helper returned; the caller saw a fake error | After an expected error: `Err.Clear` before leaving. |
| `req.ResponseText` | Charset depends on headers | Decode `ResponseBody` as UTF-8 (`Utf8Decode`). |
| Any text written with `.Value` | Converted by Excel: `true` -> TRUE, `1.10` -> 1.1, `2026-10-01` and `1/2` -> dates, `007` -> 7, `50%` -> 0.5, `#N/A` -> error, `=1+1` -> formula (tested, 11 cases) | `SafeCell()` gives **every** text a leading apostrophe; `.Value` reads it back unchanged. |
| `NumberFormat` set from **Python** | Read in French | From Python use `put_en_us()`. In VBA, `.NumberFormat` is fine (English codes). |
| PowerShell started from VBA | Blocked by the antivirus | Option B: the user starts the script; Excel only watches files. |
| Unsaved one-sheet workbook | Named `Feuil1`, `Application.Run` failed | Save before `Application.Run`. |
| JSON parser speed | 635 KB parsed in 3.7 s (about 170 KB/s) | Expect about 1 minute for 10 MB. Show progress per page. Consider option B for big projects. |
| Macro entry point is a `Function` or has parameters | Not shown in Alt+F8 and cannot be put on a button (tested) | Buttons call `Public Sub Btn...()` without arguments (see `modButtons`). |
| Non-ASCII in VBA source | Becomes `?` or breaks the generator | Keep VBA source ASCII. Use `ChrW()` (accents in status names: `"Termin" & ChrW(233)`). Read status names from the Config sheet instead of writing them in code. |

## 9. Dates and time zones

- Jira returns `2026-10-01T09:12:33.123+0200`. `ParseJiraDate` returns the **UTC** date (Tested: `+0200`,
  `+02:00`, `Z`, `-0500`, date only, null, invalid -> Empty).
- Store UTC in all sheets. Convert to local only for display, in a separate column if needed.
- Durations in days = difference of two UTC dates. Weekends are not removed (document it in `Metrics`).

## 10. Field values in cells

`FieldToCell` (Tested): user -> displayName; option -> value; named objects -> name / key / id;
arrays -> values joined with `"; "` (sprints -> `PI3-S1; PI3-S2`); rich text (ADF) -> plain text;
null -> empty; unknown objects -> compact JSON; text cut at 32 767 characters (Excel cell limit).
For data checks you sometimes need the raw value (for example all sprint ids): keep a second, narrow
sheet `RawFields` with Key, FieldId, JSON (cut at 32 767) only for the fields the checks need.

## 11. Pagination schemes (Documented)

| Endpoint family | Scheme | Stop when |
|---|---|---|
| `/rest/api/3/search/jql` | token: send `nextPageToken` back | token absent or `isLast = true` |
| `/rest/api/3/changelog/bulkfetch` | token in the POST body | `nextPageToken` null or absent |
| `/rest/agile/1.0/...` lists, `/project/search`, `/field/search` | offset: `startAt += maxResults` | `isLast = true` (or `startAt >= total` when present) |

Always bound loops with a maximum page count and stop on an error body (no `issues` / `values` key).
`JiraSearchAll` does both (`maxPages`, `CheckStatus`).

## 12. What Copilot must still write (not in the tested modules)

- `modConfig`: read the Config sheet into a Dictionary; validation messages.
- `modExtract`: the flow of section 7, the `Log` sheet, the busy flag.
- `GET /rest/api/3/field` and `/rest/agile/1.0/board/{id}/sprint` readers (same pattern as `JiraSearchAll`;
  the board sprint endpoint uses `startAt` + `isLast`, not `nextPageToken`).
- `POST /rest/api/3/search/approximate-count` before each project: show "x / about N issues" and write
  `Extracted` vs `Approximate count` in the `Log` sheet.
- Board configuration reader (estimation field per board) and field context reader (for DC11, DC12, DC25).
- `modChecks` and `modMetrics` from `safe-flow-metrics-and-data-checks.md`.
- Tests for each new procedure in `modTests` style, with fake responses (`FakeAdd`) and fixtures in
  `Tests_<module>` sheets. Keep coverage >= 90 % of markers.

Sources: Anthropic Jira API plugin (`jira:jira-api`) reference; Jira Cloud platform REST API v3 specification (search/jql, changelog/bulkfetch, issue/bulkfetch,
field), Jira Software REST API (board sprints), Atlassian "Build an app that respects rate limits".
