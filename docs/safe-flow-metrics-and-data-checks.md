# SAFe flow metrics, traceability, data checks and Jira Plans readiness

Purpose: compute the six SAFe flow metrics **from the extracted Jira data**, show **every intermediate step**
so a human can check that the right data and the right formulas are used, and **list every anomaly or
missing field** that makes the metrics or Jira Plans wrong.

The data today is known to be wrong and inhomogeneous. So: **detect and report, never "fix" silently.**
No row is dropped without a `DataChecks` row explaining why.

Labels: **SAFe** = definition from the SAFe framework site; **Documented** = Atlassian documentation;
**Decision** = to be decided by the developer and the client (do not decide in code);
**Proposal** = suggested by Claude, to be validated.

---

## 1. Context agreed with the developer

- One **Program** project (the train) holds **Features**. Several **Team** projects hold **Stories** and
  **Tasks**, with **Sub-tasks** under stories *or* tasks (mixed today; the extraction keeps both).
- Stories are linked to features through the **parent** field (as far as known; the extraction keeps
  parent, issue links and sub-tasks as they are, so the analysis can show the real situation).
- Metrics focus: Features (Program), Stories and Tasks (teams). Sub-tasks are extracted for integrity checks.
- All teams share **one workflow** (`Team`). Features have **one workflow** of their own (`Feature`). Statuses
  and moves (Jira UI, French) are defined **in the workbook, never in code**, on two sheets.

**Sheet `Workflows`** (one row per status and workflow): `Workflow | Order | Status | Category | IsStart | IsEnd | IsFree`

| Column | Meaning |
|---|---|
| Workflow | `Team` or `Feature` |
| Order | position in the normal flow (1 to 8) |
| Status | exact Jira name (with accents, read from the sheet) |
| Category | `Active`, `Wait`, `Done`, `Removed`, `Backlog`, `Unclassified` |
| IsStart | 1 on `En analyse` (start of the flow time) |
| IsEnd | 1 on `Terminé` only (end of the flow time) |
| IsFree | 1 on `Dépriorisé`: reachable from any status |

| Order | Team status | Feature status | Category |
|---|---|---|---|
| 1 | Nouveau | Nouveau | Backlog (**Decided**) |
| 2 | En analyse | En analyse | Active (**Decided**) |
| 3 | Prêt | Prêt | Wait (**Decided**) |
| 4 | Planifié | Planifié | Wait (**Decided**) |
| 5 | En développement | En développement | Active (**Decided**) |
| 6 | En test | En test / recette | Active (**Decided**) |
| 7 | À déployer | À déployer | Wait (**Decided**) |
| 8 | Terminé | Terminé | Done (**Decided**) |
| - | Annulé | Annulé | Removed (**Decided**) |
| - | Dépriorisé | Dépriorisé | Backlog (**Decided**), `IsFree = 1` |

**Sheet `Transitions`** (one row per allowed move, forward and backward): `Workflow | From | To | Type`.
The same 13 rows exist for `Team` and for `Feature` (with `En test / recette` in `Feature`):

| From | To | Type |
|---|---|---|
| Nouveau | En analyse | forward |
| En analyse | Prêt | forward |
| Prêt | Planifié | forward |
| Prêt | En analyse | back |
| Planifié | En développement | forward |
| Planifié | Prêt | back |
| En développement | En test | forward |
| En test | À déployer | forward |
| En test | En développement | back |
| À déployer | Terminé | forward |
| À déployer | En test | back |
| Dépriorisé | Annulé | exit |
| Dépriorisé | Nouveau | back to backlog |

`Dépriorisé` can be reached from **any** status: this is carried by `IsFree = 1`, not by extra rows.
`Terminé` and `Annulé` are terminal.

- **Decided:** flow time starts at the first entry in `En analyse` (`IsStart`) and ends at `Terminé` only
  (`À déployer` does not count as delivered). Moving back (for example `Prêt` to `En analyse`) does **not**
  reset the flow time. Passing through a `Backlog` status (`Dépriorisé`, then `Nouveau`) **does**: the flow
  time restarts at the next entry in `En analyse`.
- **Decided:** statuses 2 to 7 are *active* (work is done) or *waiting* (queue): `En analyse`, `En développement`
  and `En test` (`En test / recette` for Feature) are `Active`; `Prêt`, `Planifié` and `À déployer` are `Wait`.
  The `Category` column of the `Workflows` sheet **is** the status mapping: the workbook reads it, and the user
  can change a category there without touching the macros. A status missing from the sheet stays
  `Unclassified`: the checks report it (`DC08`) and Flow Efficiency shows "not computable".
- Status names contain accents: read them from the `Workflows` sheet, never write them in VBA source.

### Fields of the existing ticket extract (given by the developer)

The extract that exists today has these 22 columns (names as in the extract, French, without accents):

`Projet, Cle, Type, TypeCategorie, Resume, Status, Categorie, Livre, Resolution, Assigne, Priorite, Cree,
MisAjour, Resolu, Echeance, Label, FixVersion, PIsource, Parent, BusinessValue, StoryPoint, URL`

| Group | Fields |
|---|---|
| Always have a value (as stated) | `Projet, Cle, Type, TypeCategorie, Resume, Status, Categorie, Livre, Priorite, Cree, MisAjour, URL` |
| Can be empty | `Resolution, Assigne, Resolu, Echeance, Label, FixVersion, PIsource, Parent, BusinessValue, StoryPoint` |

| Field | Meaning (stated by the developer) |
|---|---|
| `Categorie` | Jira **status category** (`À faire` / `En cours` / `Terminé`), not our Active / Wait / Done categories |
| `Livre` | `True` / `False`: the item is **released**, that is available to the users |
| `BusinessValue` | integer |
| `Parent` | can be empty (so DC01 is meaningful; a feature has no parent) |
| `TypeCategorie` | the **work type** (**Decided**): `Feature`, `Enabler`, `Autre` or `Defect`. Used directly by Flow Distribution |
| `PIsource` | the PI field of the Jira tickets (`PI`). Values look like `PI#XX-YYYY`: `XX` = PI number, `YYYY` = year. `XX` has 2 digits (maximum `PI#99-2046`). Can be empty, or contain the text `Champ_customefield_11339`: a placeholder whose origin is unknown and not needed, reported by DC29 |

- **To confirm:** the first list had an empty slot (`MisAjour, , Parent`): a field may be missing from the
  "always have a value" list.
  `BusinessValue` can be empty, so DC21 only reports features where it is set.
- **Not in this extract:** Team, Sprint, Target start, Target end, issue links and the changelog. Checks that
  need them (DC13-DC16, DC18-DC20, DC22-DC24, and the status history for `StatusIntervals`) cannot run on this
  extract alone: list them in `MissingData` until they are extracted.

### Sheet `FieldMapping` (extract column -> logical name -> metric or check)

One row per column of the extract. It shows, for each field, which metric or check uses it and what is
wrong or not computable when it is missing. Columns: `Extract | Logical | Type | CanBeEmpty | UsedBy | Role | IfMissing | TypeConfirmed`.

- `Role`: `Required` (the metric or check cannot run without it), `Optional`, `Display` (not used in a calculation).
- `TypeConfirmed`: `Yes` = type and emptiness given by the developer (this says nothing about the other
  columns). **`UsedBy`, `Role` and `IfMissing` were validated by the developer on 2026-11-03.** Types of dates (date or date-time) are to be
  checked on the real extract.
- The sheet is data, not code: the VBA reads the logical names from it. A row can be edited without changing macros.

| Extract | Logical | Type | CanBeEmpty | UsedBy | Role | IfMissing | Status |
|---|---|---|---|---|---|---|---|
| Projet | Project | text | no | team mapping (Config), every metric per team; DC14 (needs `Team`, not in the extract) | Required | no team split; DC14 not computable | Yes |
| Cle | Key | text | no | join key of all tables, traceability | Required | nothing computable | Yes |
| Type | IssueType | text | no | DC01, DC02, Flow Velocity per type | Required | no per-type metric | Yes |
| TypeCategorie | WorkType | text (`Feature`, `Enabler`, `Autre`, `Defect`) | no | Flow Distribution | Required | Flow Distribution not computable | Yes |
| Resume | Summary | text | no | readable lists of keys | Display | none | Yes |
| Status | Status | text | no | `StatusIntervals` (CURRENT_MISMATCH), DC08, DC28 | Required | no intervals check | Yes |
| Categorie | JiraStatusCategory | text (`À faire`, `En cours`, `Terminé`) | no | DC04, DC05; DC15 (needs `Sprint`, not in the extract) | Required | DC04, DC05, DC15 not computable | Yes |
| Livre | Released | True/False (released = available to the users) | no | **usage to decide** (see below) | Optional | none | Yes |
| Resolution | Resolution | text | yes | none today | Optional | none | Yes |
| Assigne | Assignee | text | yes | none today | Display | none | Yes |
| Priorite | Priority | text | no | Plans ranking (display) | Display | none | Yes |
| Cree | CreatedUtc | date-time | no | Lead Time, BEFORE_CREATED | Required | no lead time, no interval start | Yes |
| MisAjour | UpdatedUtc | date-time | no | "as of" control | Optional | none | Yes |
| Resolu | ResolvedUtc | date-time | yes | DC04, DC05 | Required | DC04, DC05 not computable | Yes |
| Echeance | DueDate | date | yes | DC20 if the plan uses due dates | Optional | none | Yes |
| Label | Labels | text | yes | none today | Optional | none | Yes |
| FixVersion | FixVersion | text | yes | DC26 | Optional | DC26 not computable | Yes |
| PIsource | PI | text (`PI#XX-YYYY`, `XX` 2 digits) | yes | DC17, DC18, DC29 | Required for PI views | DC17, DC18 not computable | Yes |
| Parent | ParentKey | text (issue key) | yes | DC01, DC02, DC24, hierarchy | Required | DC01, DC02, DC24 not computable | Yes |
| BusinessValue | BusinessValue | integer | yes | DC21 only (Flow Predictability needs the business value of PI objectives, which are not in the extract) | Optional | DC21 not computable | Yes |
| StoryPoint | StoryPoints | number | yes | DC12, Flow Velocity in points, fill rate; DC11 (needs `Sprint`, not in the extract) | Optional | points view "not reliable"; DC11 not computable | Yes |
| URL | Url | text | no | link from a metric to the issue (traceability rule 3) | Display | none | Yes |

`Released` is not the end of the flow time: the end stays `Terminé` (`IsEnd`). `Released = True` is a
different fact (the item reached the users). Possible uses, **none decided**: a check "`Terminé` but not
released" (information), or "released but not `Terminé`" (inconsistency). Until decided, `Released` is only
extracted and shown.

Every missing column of the extract (Team, Sprint, Target start, Target end, issue links, changelog) gets a
row too, with `TypeConfirmed = Missing`, so `MissingData` can list what is not available.

## 2. The six flow metrics (SAFe definitions) and how we compute them

SAFe definitions (framework.scaledagile.com, "Accelerating Flow with SAFe"):
Flow Distribution = proportion of each backlog item type in the workflow; Flow Velocity = number of items
completed in a given time; Flow Time = time from when an item enters the workflow to its release;
Flow Load = total work in process across all steps; Flow Efficiency = time actively worked as a percentage
of total elapsed time; Flow Predictability = overall planned versus actual business value.

The formulas below are our implementation choices on top of these definitions. Each one names its inputs,
so a reviewer can check them. All dates are UTC. Durations are calendar days (weekends included).

### 2.1 Base table `ItemFlow` (one row per issue, all formulas visible)

| Column | Formula / source |
|---|---|
| Key, Team, Type, Project role | issue fields + Config project table |
| CreatedUtc | `created` |
| StartUtc | first `EnteredUtc` in `StatusIntervals` of a status with `IsStart = 1`, **after the last `EnteredUtc` of a status with category `Backlog`** (the flow time restarts after a return to backlog) |
| BacklogReturns | count of entries in a `Backlog` status after a first start (makes the restart visible) |
| DoneUtc | last `EnteredUtc` of a status with `IsEnd = 1`, **only if the issue is in that status now** |
| RemovedUtc | same for category `Removed` (`Annulé`; excluded from velocity and flow time) |
| FlowTimeDays | `DoneUtc - StartUtc` (empty if not done or never started) |
| LeadTimeDays | `DoneUtc - CreatedUtc` |
| ActiveDays | `SUMIFS(StatusIntervals[DurationDays], [Key], key, [Category], "Active", [EnteredUtc], ">="&StartUtc, [EnteredUtc], "<"&DoneUtc)` |
| WaitDays | same with `"Wait"` |
| UnclassifiedDays | same with `"Unclassified"` (must be 0 for a valid efficiency) |
| Reopened | count of transitions from a Done status to a non-Done status (from `StatusIntervals`) |
| SkippedStart | Y if done without any interval in a start status after its last `Backlog` entry (flow time would be wrong) |
| Checks | concatenated `Check` values of its intervals + its DataChecks ids |

Write these as **Excel formulas** (structured references on tables), not VBA-computed constants, so the user
can click a cell and see how it is computed. VBA only writes the raw tables (`StatusIntervals`, issues).

### 2.2 Metric formulas (per team, per type, per period: iteration or PI from Config)

| Metric | Formula | Inputs to show next to it |
|---|---|---|
| **Flow Velocity** | `COUNTIFS(ItemFlow[DoneUtc], ">="&PeriodStart, ItemFlow[DoneUtc], "<"&PeriodEnd, ItemFlow[Team], team, ItemFlow[Type], type)`. Optional second view: sum of story points of the same items (only if the points fill rate is above a threshold, else show "not reliable") | list of the counted keys (FILTER formula), fill rate of points |
| **Flow Time** | median and 85th percentile of `FlowTimeDays` for items done in the period: `MEDIAN(FILTER(...))`, `PERCENTILE.INC(FILTER(...),0.85)` | count of items, items excluded (SkippedStart, Reopened) |
| **Flow Load** | for each day D, count of items in work in process. Parameter `FlowLoadMode` in Config: **`A` (default)** = item counted from its first start until it is done or removed, except while it sits in a `Backlog` status (use the `StatusIntervals` rows of D: not `Backlog`, not terminal, and first start `<= D`); **`B`** = item counted only from its last `StartUtc` (after the last return to backlog) until `DoneUtc` or `RemovedUtc`. Table `DailyLoad` (Date x Team) built from `ItemFlow` and `StatusIntervals`. Report average and end-of-period value, and show the mode used | the `DailyLoad` rows of the period, the mode `A` or `B` |
| **Flow Efficiency** | `SUM(ActiveDays) / SUM(ActiveDays + WaitDays)` over items done in the period. Show **"not computable"** while any status of these items is `Unclassified` (for example a status missing from `Workflows`) | ActiveDays, WaitDays, UnclassifiedDays sums |
| **Flow Distribution** | share of items by work type in the period: done items (and, separately, WIP at period end). Work type = the `TypeCategorie` column (**Decided**): `Feature`, `Enabler`, `Defect`, `Autre`. `Risk` and `Debt` (SAFe) do not exist in the extract; an empty or unknown value is counted as `Unmapped` | counts per type, `Unmapped` count |
| **Flow Predictability** | `SUM(actual business value) / SUM(planned business value)` of **committed PI objectives** per team and for the ART. Uncommitted objectives add to actual, not to planned. Reliable trains operate in the **80-100 %** range (SAFe) | per objective: planned BV, actual BV, committed Y/N |

**Predictability today is not computable:** there are no PI objective items in Jira, and business value is
sometimes put on features. The workbook must show "not computable" and list what is missing (section 4).

### 2.3 Traceability rules

1. Every metric cell has, next to it, the count of items used and a link/filter to the exact rows.
2. Every excluded item appears in `DataChecks` with the reason. `Included + Excluded = Extracted` is checked
   by a formula on the `Metrics` sheet (show OK/KO).
3. Every interval keeps the `ChangeId` that ended it; every issue row keeps its Jira key: a reviewer can
   go from a metric to the issue history in Jira.
4. The `Metrics` sheet states the "as of" date, the period table, and the status mapping version used.
5. Tests: for each metric, a `Tests_Metrics` sheet with a small fixture (5-10 issues with hand-computed
   expected values) and the OK/KO row in `TestResults`.

## 3. Data checks catalogue (same ids in Excel, Rovo prompts, saved filters and automations)

Severity: **Blocker** (metric or Plans wrong), **Major** (metric biased), **Minor** (hygiene).
The JQL is a starting point: field names differ per site; take the real ids from `FieldCatalog`
(in JQL, a custom field can be written `cf[10016]`).

| Id | Check | Sev. | Why it matters | Excel rule (from extracted data) | JQL to find them |
|---|---|---|---|---|---|
| DC01 | Story or task without parent feature | Blocker | Feature progress, Plans roll-up, distribution | team issue of type Story/Task with empty `parent` | `project in (TEAMS) AND issuetype in (Story, Task) AND parent is EMPTY` |
| DC02 | Parent is not a feature (wrong level) | Major | Broken hierarchy | parent type not in Config "feature types" | (Excel only) |
| DC03 | Sub-task parent type mixed (story vs task) | Minor | Inhomogeneous model | distribution of sub-task parent types per team | `issuetype in subTaskIssueTypes()` then group by parent type (Excel) |
| DC04 | Done without resolution date | Major | Velocity / reports using resolution | status category Done and `resolutiondate` empty | `statusCategory = Done AND resolutiondate is EMPTY` |
| DC05 | Resolution date but not done | Major | Same | `resolutiondate` set and category not Done | `statusCategory != Done AND resolutiondate is not EMPTY` |
| DC06 | Changelog `FROM_MISMATCH` | Major | Status history inconsistent | `StatusIntervals[Check]` contains it | (Excel only) |
| DC07 | `BEFORE_CREATED` change | Major | Negative or wrong durations | idem | (Excel only) |
| DC08 | Status not in `Workflows` or category `Unclassified` | Blocker for efficiency | Efficiency not computable | category Unclassified | (Excel only) |
| DC09 | Done without ever being in a start status | Major | Flow time 0 or empty | `SkippedStart = Y` | (Excel only) |
| DC10 | Reopened after Done | Major | Velocity counted twice, flow time wrong | `Reopened > 0` | (Excel only, from changelog) |
| DC11 | Story without story points in an active or future sprint | Blocker for capacity | Plans capacity, velocity in points | Story, sprint state active/future, points empty | `issuetype = Story AND sprint in (openSprints(), futureSprints()) AND cf[POINTS] is EMPTY` |
| DC12 | Story points on sub-tasks | Major | Double counting | sub-task with points | `issuetype in subTaskIssueTypes() AND cf[POINTS] is not EMPTY` |
| DC13 | No Team value | Blocker for Plans | Capacity per team, filters | `Team` empty on Story/Task/Feature | `project in (TEAMS) AND cf[TEAM] is EMPTY` (check the real JQL name of the Team field) |
| DC14 | Team value does not match the team project | Major | Wrong capacity | Team field vs Config project -> team | (Excel only) |
| DC15 | In progress without sprint | Major | Sprint-based scheduling | status category In progress, sprint empty | `project in (TEAMS) AND statusCategory = "In Progress" AND sprint is EMPTY` |
| DC16 | Carried over 2+ sprints | Minor (signal) | Predictability, overload | count of sprint values in changelog per issue >= 3 | `sprint in closedSprints() AND statusCategory != Done` (approximate) |
| DC17 | Feature without PI value | Blocker for PI views | PI planning, predictability | Feature with PI custom field empty | `issuetype = Feature AND cf[PI] is EMPTY` |
| DC18 | Feature PI inconsistent with its stories' sprints | Major | Plan vs reality | child sprint dates outside the PI dates (Config PI table) | (Excel only) |
| DC19 | Feature without fixed dates (Target start / Target end) | Blocker for Plans timeline | Plans schedules with target dates by default | empty Target start or Target end on features in status >= Planifié | `issuetype = Feature AND ("Target start" is EMPTY OR "Target end" is EMPTY)` |
| DC20 | Target end before Target start, or child dates outside parent dates | Major | Wrong timeline | date comparisons | (Excel only) |
| DC21 | Business value on a feature | Major | BV belongs to PI objectives | Feature with Business Value set | `issuetype = Feature AND cf[BV] is not EMPTY` |
| DC22 | Dependency risk: blocker ends after the blocked item starts | Major | Plans dependency warnings | `issuelinks` type Blocks, compare dates / sprints | `issueLinkType in ("blocks", "is blocked by")` |
| DC23 | Cross-team dependency without both items planned | Minor | Hidden risk | link between two team projects, one side without sprint | (Excel only) |
| DC24 | Feature without children | Minor | Empty scope | no child in team projects | (Excel only: standard JQL cannot test "has no children") |
| DC25 | Field fill rate below threshold | Major | Global data quality | matrix field x team x type, % filled | (Excel only; `FieldCatalog`) |
| DC26 | Missing release / fix version on features (if the client uses releases in Plans) | Minor | Release view in Plans | `fixVersions` empty | `issuetype = Feature AND fixVersion is EMPTY` |
| DC27 | Issue moved between projects (key changed) | Minor | Team history and metrics split across keys | changelog rows with field `Key` or `project`; join on `Id` | (Excel only) |
| DC28 | Status transition not in `Transitions` | Major | Workflow bypassed, intervals unreliable | an interval change (From, To) with no matching `Transitions` row (same workflow) and `To` not `IsFree` | (Excel only) |
| DC29 | PI field not empty but not in the format `PI#XX-YYYY` (includes the placeholder `Champ_customefield_...`) | Major | The PI looks filled but is not a real value; PI views, DC17 and DC18 are wrong | `PIsource` not empty and not matching `PI#` + 2 digits + `-` + 4 digits (for example `LIKE "PI#??-????"` with digit test); treated as empty for DC17 and DC18 | (Excel only) |

Each check writes rows to `DataChecks` with the same Id, so the Excel results, the Rovo agents and the Jira
saved filters (`DQ DC01 ...`) can be compared.

## 4. "What is missing" report (sheet `MissingData`)

Copilot must produce, and Rovo prompts must ask for, a list of what is missing to steer the train:

1. **Missing objects**: PI Objective items (none today), PI / iteration calendar (Config table if not in Jira),
   active/wait status categories (statuses 2 to 7), feature types list.
2. **Missing fields** (per team and type, with fill rate): Team, Sprint, Story points, Parent, PI,
   Target start, Target end, Business value (on objectives), Fix version, Resolution date.
3. **Not computable metrics** and why (for example: efficiency -> statuses unclassified; predictability ->
   no objectives with planned and actual BV).
4. **Plans blockers** (section 5).

Each line: Item, Scope (team / type), Count, % of scope, Impact (metric or Plans view), Suggested owner
(team, RTE, product management), Check id.

## 5. Jira Plans (Advanced Roadmaps) readiness

What Plans uses (Documented, Atlassian Plans / Advanced Roadmaps docs; verify the settings of the client plan):

| Plans feature | Data it needs | Our check |
|---|---|---|
| Timeline scheduling | By default **Target start** and **Target end**; the plan can be set to use due dates, sprint dates or custom dates instead (plan settings > Scheduling > Dates) | DC19, DC20 |
| Hierarchy (program board view) | Parent links consistent from Feature down to Story/Task; hierarchy levels configured (Feature above Epic/Story level as set by the admin) | DC01, DC02, DC24 |
| Teams and capacity | Every item has a **Team** (Atlassian team or plan-only team); each team has velocity (story points per iteration, default 30) or hours per week; Scrum teams have **sprints**; issues are **estimated** | DC11, DC13, DC14, DC15, DC12 |
| Auto-scheduler | Scrum/Kanban method, sprint assignments, remaining capacity, start/end dates, ranking, **dependencies**, estimates, team members | all of the above + DC22 |
| Dependencies | Issue links of type **Blocks** by default (other link types must be added in plan settings) | DC22, DC23 |
| Releases | Fix versions (if releases are used) | DC26 |
| Capacity view | Columns priority, sprint, story points, team, release | DC11, DC13 |

**Proposal (cleaning order):** 1) status categories (`Workflows` sheet); 2) hierarchy (DC01-DC03);
3) teams (DC13-DC14); 4) sprints (DC15-DC16); 5) estimates (DC11-DC12); 6) PI and fixed dates (DC17-DC20);
7) business value moved to objectives (DC21); 8) dependencies (DC22-DC23); 9) re-extract, recompute
metrics, compare; 10) switch on the workflow validators and the Jira automations that prevent regressions (see `rovo-guide.md` section 6).

## 6. Tests required for this part

- One fixture per check: an issue that must trigger it and one that must not (OK/KO each).
- One fixture per metric with hand-computed values (write the calculation in the Input column).
- A reconciliation test: Included + Excluded = Extracted.
- Coverage >= 90 % of markers in `modChecks` and `modMetrics`.

Sources: SAFe "Accelerating Flow with SAFe" (flow metric definitions); SAFe Inspect & Adapt / PI objectives
(predictability, 80-100 % range); Atlassian Advanced Roadmaps / Plans docs (dates, capacity and velocity,
auto-scheduler, dependencies with the Blocks link type, sprint capacity view); Jira JQL reference.
