# Rovo guide: prompts, agents, skills and automations for Jira data quality

> **Updated copy, 2026-11-05.** The original is `docs/rovo-guide.md` (frozen: its MD5 is tracked in
> `claude-code-handoff.md`); use this copy. Sections 1, 3 (coverage), 4, 5, 6 and 7 were revised with the research
> report `C:\dev\reports\Agents Rovo Studio et métriques.md` (outside the repo). The agents of this folder:
> - `rovo_settings_analyser.md`: the Jira configuration compared with the workbook.
> - `rovo_flow_metric.md`: flow metrics per work item type, each one labelled.
> - `rovo_train_steering.md`: the steering cycle. It runs the data checks, explains the figures published with
>   `confluence_flow_metrics_page.md`, and drafts a steering note. Its skills are in `skills/`. It replaces the Flow
>   Data Steward of section 4.
>
> The Jira dashboard of saved filters is in `jira_dashboard_data_quality.md`.

For: **you** (Jira admin, using Rovo Chat and Rovo Studio) and **Copilot** (which builds the Excel side).
Goal: find and fix the data anomalies that make the SAFe flow metrics wrong and keep the Jira **Plans**
program board unusable, then **prevent** them from coming back.

The check ids `DC01`...`DC29` are the same as in `safe-flow-metrics-and-data-checks.md` and in the Excel
`DataChecks` sheet. Always use them: it lets you compare Rovo, saved filters and Excel counts.

Labels: **Documented** = Atlassian support docs (checked 2026-11-03, revised 2026-11-05); **Community** = Atlassian
Community posts, staff replies or the public tracker (jira.atlassian.com); **Beta** = marked beta by Atlassian;
**Verify in UI** = exact label not confirmed in the docs; **Proposal** = suggested, to validate.

---

## 1. What Rovo can and cannot do (revised 2026-11-05)

**Building an agent.** Studio > Agents > Create (or from Rovo Chat, the Jira assignee picker, the JSM "AI agents" page,
or *Duplicate*), then a guided chat or **Skip to manual setup** (Documented). Since 2026-04-28 an agent has two layers
(Community, staff post; older "scenarios" were migrated to subagents by 2026-06-15):
- **Identity:** Name, Description, **Behavior** (rules applied in every turn), **Conversation starters** (at most
  three; left blank, three generic ones are shown).
- **Base configuration:** Instructions, Skills, Tools, Knowledge, Reasoning. **Subagents** repeat these fields
  (name, trigger also called routing description, instructions, tools, skills, knowledge) and **inherit nothing** from
  the main agent; each has its own context window. Routing uses only the main instructions, the user's query and the
  triggers. Response mode: *Combine subagent outputs* (default: the main agent rewrites them, which broke JSON in an
  Atlassian KB case) or *Forward the subagent's output as-is* (one subagent per turn; keeps tables and JQL unchanged).
  Limit 100 subagents; staff advise starting with fewer than five (Community, staff reply).
- **Tools** vs **skills:** a tool is one action with inputs and outputs; a skill is reusable instructions that combine
  tools (Documented). Atlassian recommends **five tools at most**. Consequential tools ask for confirmation. Jira bulk
  actions: **20 work items** at a time (Documented). The full tool list shows only in Studio; the docs give examples
  (Search with JQL, Create work item, Update status). No documented tool reads the change history or the workflows of a
  custom agent; Atlassian's own Jira Delivery agent has a "get changelog" tool, one work item per call (Documented):
  check the picker on the site.
- **Knowledge:** All organizational knowledge (default), **Custom knowledge** (chosen spaces, projects, page trees,
  connector filters) or none; a **Web search** toggle if the admin allows it. Answers are always trimmed to what the
  user may see (Documented). Deep research mode exists on subagents with Custom knowledge only (15 minutes, 30 requests
  a day).
- **Reasoning** (Details > Reasoning, Documented): **Quick answers** (10 loops) or **Think deeper** (50 loops). Think
  deeper shows a model list with a **Default** model; Atlassian names no models and gives 30 days' notice before
  retiring one. "The same agent configuration does not guarantee the same response across models" (Documented, "Model
  behavior guidelines"). The new picker reached 50 % of organisations on 2026-10-09 (Community, staff post).
- **Access and identity:** the agent runs with the **User's account** or its own **Agent's account** (recommended by
  Atlassian for automation; admin grants need Guard Premium) (Documented).
- **Users and permissions:** Owner, Editor, Manager; *Open to all users* on by default; no restriction by group or
  team. Anyone who can use an agent can read its instructions (Community): keep nothing secret in them.
- **Drafts and versions:** an Unpublished tab, a version picker and View History (Community, staff post, 2026-08).
  Import and export are not settled: keep the master text of every agent in this repo folder.
- **Evaluations:** an Evaluation tab takes a CSV of up to **100 prompts** with optional expected answers, scored by an
  LLM judge (Response accuracy, Resolution rate, Manual) (Documented). The `{ }` button under an answer shows the
  subagent, agent, request and **Language Model ID**.
- **Who may create agents:** an admin setting: all users (default), up to 10 groups, or nobody (Documented).

**Automations** (Documented). Action **Use agent** gives `{{agentResponse}}`; **Use Rovo** (no agent) gives
`{{rovoResponse}}`. Triggers include work item events and scheduled runs. Agents taking actions in automations started
rolling out on 2026-10-06 (tracker ROVO-261, closed Fixed): the older doc sentence "agents in automations can't use
their tools" is outdated, so give a reporting agent no write tool at all. Forge actions called from automation run
only if they are GET. ROVO-1148 (empty `{{agentResponse}}` when the agent called a tool) is fixed, but test that the
response is not empty before using it.

**Credits** (Documented, rate card from 2026-08-31): per paid user and month, 25 / 70 / 150 credits for Jira or
Confluence (Standard / Premium / Enterprise), pooled per organisation, no roll-over. Basic agents: 10 credits per call;
agents with a chosen model and Think deeper: variable. Extra usage is billed from **2026-12-03** at $0.01 per credit,
on by default (admins can cap it). Measure the cost of a scheduled agent on the site before enabling it.

**Counts and limits.** Atlassian: Rovo search is "not a good fit" to "analyze the work items to create charts,
summaries, or other representations of your data" (Documented). An agent reasons over **at most about 50 work items**
of a JQL result (Atlassian Support, relayed on the Community, 2026-01; not in the docs). Open suggestions: ROVO-396
(wrong results from JQL lists), ROVO-535 (cannot count or sum), ROVO-509 (cannot aggregate number fields). The Jira
search API returns **no total** (Documented, REST v3). JQL never returns a duration, and history operators work only on
Status, Resolution, Assignee, Fix Version, Priority and Reporter (Documented).

Consequence (unchanged, now with sources): use Rovo to **find, explain and propose**; take every number from JQL in Jira
issue search, from a Jira-computed count in an automation, or from the Excel workbook; prevent with workflow validators
and automation. **No flow-metric figure comes from a Rovo agent.** Every prompt asks for the JQL, the number of work
items received and the key list, so a cut result is visible.

## 2. Prompt pattern (use it for every prompt)

```text
ROLE: You are a Jira data-quality auditor for a SAFe Agile Release Train.
GOAL: <one sentence>.
SCOPE: Use this JQL only: <JQL>. Projects: Program = <PROGRAM_KEY>, Teams = <TEAM_KEYS>.
FIELDS: <list of fields with their exact names on this site>.
RULES: <numbered rules, each with its check id DCxx>.
OUTPUT:
  1. A summary line per rule: check id, count, "exact" or "estimated", % of scope, the JQL to check it.
  2. A table: Check id | Work item key | Team | Field | Found value | Expected | Fix hint.
  3. If 50 or more items come back, the list may be cut: give the keys received, write "at least N", and the JQL
     (the exact count is in Jira issue search).
CONSTRAINTS: Read only. Do not edit anything. Use only data you can see. If a field does not exist or is
  hidden, say so instead of guessing. Quote work item keys exactly.
```

Placeholders to fill once (keep them in a Confluence page "ART data rules", see section 7):
`<PROGRAM_KEY>`, `<TEAM_KEYS>` (comma list), `<POINTS_FIELD>` (for example `Story point estimate`),
`<TEAM_FIELD>` (`Team`), `<PI_FIELD>` (`PI`, values like `PI#XX-YYYY`: `XX` = 2-digit PI number, `YYYY` = year), `<BV_FIELD>` (`Business Value`), `<TARGET_START>` / `<TARGET_END>`
(`Target start` / `Target end`), `<FEATURE_TYPE>` (`Feature` or its French name on the site).
Field names in JQL differ per site: get the exact names with prompt P00 or from the Excel `FieldCatalog`.

## 3. Prompt library (copy, fill the placeholders, run in Rovo Chat)

**Coverage (revised 2026-11-05, checked against the rule code `src/jira2dashboard/flow/checks.py`):** 10 checks have
a JQL that should return exactly the workbook's `DataChecks` list, 9 only partly, 1 depends on a syntax test, 9 need
the workbook.
- **Exact (10):** DC04, DC05, DC11, DC12, DC13, DC14, DC15, DC19, DC21, DC26. They read only the current state, with no
  date: a difference with the workbook on the same as-of date is a real signal (scope, field id, renamed status,
  permissions).
- **Partly (9):** DC08 and DC09 (current status, or `WAS NOT IN`, which ignores the restart after a return to
  Backlog), DC10 (exact only with one Done status per workflow), DC16 (the number of sprints cannot be counted), DC17
  (a malformed PI counts as empty in the workbook), DC22 and DC23 (candidates only), DC25 (links, not the ratio), DC28
  (chosen forbidden pairs only).
- **Depends on a test (1):** DC01: `parent IS EMPTY` is not documented for `parent`, community answers say it works.
- **Workbook only (9):** DC02, DC03, DC06, DC07, DC18, DC20, DC24, DC27, DC29 (parent types, history chains, dates
  across items, regular expressions).

Exact JQL (prefix each with the workbook scope `project in (<Projects>) AND (<Filter>)`; placeholders are the
workbook parameters, filled at the client only):

| DC | JQL |
|---|---|
| DC04 | `statusCategory = Done AND resolutiondate IS EMPTY` |
| DC05 | `statusCategory != Done AND resolutiondate IS NOT EMPTY` |
| DC11 | `issuetype in (<PointsRequiredTypes>) AND sprint in (openSprints(), futureSprints()) AND "<StoryPointsField>" IS EMPTY` |
| DC12 | `issuetype in (<SubTaskTypes>) AND "<StoryPointsField>" IS NOT EMPTY` |
| DC13 | `issuetype in (<FeatureTypes>, <ParentRequiredTypes>) AND "<TeamField>" IS EMPTY` |
| DC14 | per team project P: `project = <P> AND "<TeamField>" IS NOT EMPTY AND "<TeamField>" != "<TeamOfP>"` (test how the Team field compares) |
| DC15 | `project in (<TeamProjects>) AND statusCategory = "In Progress" AND sprint IS EMPTY` |
| DC19 | `issuetype in (<FeatureTypes>) AND status not in (<BacklogStatuses>, <RemovedStatuses>) AND ("<TargetStartField>" IS EMPTY OR "<TargetEndField>" IS EMPTY)` |
| DC21 | `issuetype in (<FeatureTypes>) AND "<BusinessValueField>" IS NOT EMPTY` |
| DC26 | `issuetype in (<FeatureTypes>) AND fixVersion IS EMPTY` |

The prompts below predate this table: where they differ (P03 uses Story instead of `PointsRequiredTypes`; P04 limits
DC13 to some types; P05 DC19 names one status), the table wins.

**P00 - Field discovery (run first)**

```text
ROLE: Jira configuration analyst.
GOAL: List the exact field names and types used by projects <PROGRAM_KEY>, <TEAM_KEYS>.
OUTPUT: a table Field name | Field type | Used in projects | Example value, for these fields:
  parent, Team, Sprint, story points (any variant), PI or Program Increment, Target start, Target end,
  Business Value, Fix version, Resolution, Due date. Also list any custom field whose name contains
  "PI", "Objective", "Value", "Point", "Team", "Iteration".
CONSTRAINTS: Read only. Say "not found" when a field does not exist.
```

**P01 - Hierarchy (DC01, DC02, DC03, DC24)**

```text
ROLE: Jira data-quality auditor for a SAFe ART.
GOAL: Find broken links between features and team work.
SCOPE: project in (<TEAM_KEYS>) AND issuetype in (Story, Task) AND statusCategory != Done
FIELDS: parent, issuetype, status, Team, sprint.
RULES:
  DC01: the item has no parent.
  DC02: the parent is not of type <FEATURE_TYPE> or is not in project <PROGRAM_KEY>.
  DC03: for sub-tasks of these items, report how many are under a Story and how many under a Task, per project.
  DC24: features in <PROGRAM_KEY> (statusCategory != Done) with no child item in <TEAM_KEYS>.
OUTPUT and CONSTRAINTS: as in the standard pattern.
```

**P02 - Status and resolution (DC04, DC05, DC10)**

```text
GOAL: Find items whose status, resolution and history disagree.
SCOPE: project in (<PROGRAM_KEY>, <TEAM_KEYS>) AND updated >= -180d
RULES:
  DC04: statusCategory = Done and Resolved date is empty.
  DC05: Resolved date is set but statusCategory != Done.
  DC10: the item was Done, then moved back to a non-done status (look at its history).
        Best effort: you may not see the full history. Say so. The Excel workbook is the reference for DC10.
```

**P03 - Estimation (DC11, DC12)**

```text
GOAL: Find estimation gaps that break capacity in Plans and velocity in points.
SCOPE: project in (<TEAM_KEYS>) AND sprint in (openSprints(), futureSprints())
FIELDS: issuetype, <POINTS_FIELD>, sprint, Team.
RULES:
  DC11: Story with empty <POINTS_FIELD>.
  DC12: sub-task with a value in <POINTS_FIELD> (double counting).
OUTPUT: add a line per team: stories in scope, stories estimated, % estimated.
```

**P04 - Team and sprint (DC13, DC14, DC15, DC16)**

```text
GOAL: Check that every planned item has a team and a sprint.
SCOPE: project in (<PROGRAM_KEY>, <TEAM_KEYS>) AND statusCategory != Done
RULES:
  DC13: <TEAM_FIELD> is empty (features, stories, tasks).
  DC14: <TEAM_FIELD> does not match the team that owns the project (mapping: <PROJECT=TEAM list>).
  DC15: statusCategory = "In Progress" and sprint is empty (team projects only).
  DC16: the item has been in 3 or more sprints (sprint history) and is not done.
        Best effort, as for DC10: the Excel workbook is the reference.
```

**P05 - PI and fixed dates (DC17, DC18, DC19, DC20)**

```text
GOAL: Check that features are planned in a PI and have landing dates for Plans.
SCOPE: project = <PROGRAM_KEY> AND issuetype = <FEATURE_TYPE> AND statusCategory != Done
FIELDS: <PI_FIELD>, <TARGET_START>, <TARGET_END>, status, child items and their sprints.
RULES:
  DC17: <PI_FIELD> is empty. (The format PI#XX-YYYY is checked by DC29 in Excel only.)
  DC18: the sprints of its children fall outside the PI dates (PI calendar: <PI list with start-end dates>).
  DC19: status is Planifié or later and <TARGET_START> or <TARGET_END> is empty.
  DC20: <TARGET_END> is before <TARGET_START>, or a child has dates outside the feature dates.
```

**P06 - Business value (DC21)**

```text
GOAL: Find business value set on features instead of PI objectives.
SCOPE: project = <PROGRAM_KEY> AND issuetype = <FEATURE_TYPE> AND "<BV_FIELD>" is not EMPTY
OUTPUT: table key | PI | business value | status. Then explain in 3 sentences why business value belongs to
  PI objectives in SAFe (planned and actual BV feed the predictability measure) and what is missing in this
  Jira to store objectives.
```

**P07 - Dependencies (DC22, DC23)**

```text
GOAL: Find dependency risks Plans will flag.
SCOPE: project in (<PROGRAM_KEY>, <TEAM_KEYS>) AND issueLinkType in ("blocks", "is blocked by") AND statusCategory != Done
RULES:
  DC22: the blocking item is planned (sprint end or target end) after the blocked item starts.
  DC23: the link crosses two team projects and one side has no sprint or no target dates.
OUTPUT: table blocking key | blocked key | teams | blocking end | blocked start | risk.
```

**P08 - Fill rate and missing data report (DC25 + section 4 of the metrics file)**

```text
GOAL: Tell me what data is missing to compute the six SAFe flow metrics and to run the Plans program board.
SCOPE: project in (<PROGRAM_KEY>, <TEAM_KEYS>) AND created >= -365d
OUTPUT:
  1. Fill rate table: field x team x issue type (Team, sprint, <POINTS_FIELD>, parent, <PI_FIELD>,
     <TARGET_START>, <TARGET_END>, Resolved, Fix version).
  2. For each metric (velocity, flow time, load, efficiency, distribution, predictability): the data it
     needs, whether it exists here, and what is missing.
  3. Objects that do not exist yet (for example PI objective items).
  4. A prioritised list of 10 actions, each with the check id it fixes.
CONSTRAINTS: Read only. Give counts, not guesses. Say when you only looked at a sample.
```

**P09 - Plans readiness**

```text
GOAL: Decide if the plan "<PLAN NAME>" can show a reliable program board, capacity and dependencies.
RULES: check DC01, DC11, DC13, DC15, DC19, DC20, DC22 on the plan scope (<JQL of the plan issue sources>).
OUTPUT: Ready / Not ready per Plans feature (timeline, hierarchy, capacity, dependencies, releases), the
  blocking check ids with counts, and the 5 most impactful fixes.
```

**P10 - Fix proposal (only after a human reviewed P01-P09)**

```text
GOAL: Propose corrections for check <DCxx> on these items: <keys, max 20>.
OUTPUT: table key | field | current value | proposed value | reason | confidence (high/medium/low).
CONSTRAINTS: Do not change anything. Wait for my explicit "apply" on a list of keys. Never more than 20
  items per change.
```

## 4. Agent to create in Rovo Studio: Flow Data Steward (revised 2026-11-05)

> **Replaced, 2026-11-05.** For the steering cycle, create the agent ART Flow Steering (`rovo_train_steering.md`)
> instead. It has one agent with three jobs and no subagents, and it reads the Confluence page "Flow metrics <PI>".
> The procedure of subagent 1 below became its skill `skills/steering-readiness/SKILL.md`. This section stays for its
> reasoning and for subagent 3 (Item investigator), which can come back as a fourth job if a changelog tool exists.

This replaces the "ART Data Quality" sketch of the original guide, which predates subagent isolation, the
forward-as-is mode and the evidence on counts. One rule above the others: **it never states a total it computed, and
never computes a metric.** It gives three kinds of answer: JQL links and key lists for the checks, explanations of the
workbook's figures, and the history of one work item. Create it with *Skip to manual setup*, as a draft; keep this text
as the master copy.

| Setting | Value | Why |
|---|---|---|
| Name / Description | Flow Data Steward. "Explains the jira2dashboard flow metrics and finds the Jira data problems (DC01-DC29) behind them. Gives JQL links, never computes a metric. Read only." | the description sets expectations |
| Behavior | the role and the audience, plus: "Numbers come from the workbook or from Jira, never from you. Saying what you could not see is a good answer." | applies in every turn |
| Base configuration | one line of scope and a polite refusal for anything else; **no tools, no organisational knowledge** (Proposal: a community trick against drift, to retest after the 2026-04 refresh) | keeps the agent on its job |
| Subagent 1: Check finder | Trigger: "the user names a DCxx check, or asks which work items lack a parent, team, sprint, points, PI, dates or resolution, or asks for a filter". Tools: Search with JQL (plus get work item if offered). Knowledge: Custom: the Confluence page "ART data rules" with a table DC, Name, JQL match (exact / partial / none), JQL template, the threshold T, and the ART projects | section 3 decides what it may claim |
| Subagent 2: Metrics explainer | Trigger: "the user asks what a flow metric means, why its value is what it is, why it is not computable, or which data it needs". Tools: Search with JQL, for drill-down links only. Knowledge: a definitions page and a page "Flow metrics <PI>" where the RTE pastes the workbook's `Metrics` and `DataChecks` summary (a client policy decision) | quotes workbook values with their as-of date, never recomputes them |
| Subagent 3: Item investigator | Trigger: "the user gives one work item key and asks about its history, flow time or flags". Tools: get work item and get changelog **only if the picker offers them; otherwise drop this subagent** | explains one item; durations still come from the `ItemFlow` row |
| Response mode | Forward the subagent's output as-is | keeps tables and JQL strings unchanged |
| Reasoning | Think deeper on Default; then evaluate any Claude model the list offers against it (section 7) and keep the one that passes | answers with several searches outgrow 10 loops |
| Starters | "Run check DC13 on <project>", "Why is Flow Efficiency not computable?", "Explain the history of <KEY>" | one per subagent, three at most |
| Identity | User's account in chat. In automation, the Agent's account with read-only access to the ART projects, if the admins can grant it (Guard Premium); otherwise the rule owner's account | Documented, agent accounts |
| Automation | see section 6, rule R05 | the count comes from Jira |

Instructions of subagent 1 (the rules live in the Confluence table, so the text stays short; client values such as
the site URL, field ids and project keys go on that page at the client, never in the repo):

```text
<role>
You find the Jira work items that fail the jira2dashboard data checks (DC01-DC29) for this ART. People fix data from
your answer before the workbook computes the SAFe flow metrics, so a wrong key or an invented count costs them time.
Saying what you could not see is a good answer.
</role>
<procedure>
1. Find each requested check in the table "DC checks" on the page "ART data rules": Name, JQL match, JQL template.
2. If JQL match is "none", explain the rule in two sentences, say that the workbook sheet DataChecks is the only
   source, and stop for that check.
3. Otherwise fill the placeholders from the same page and run the JQL once with Search with JQL.
4. List the keys you received. Read the truncation threshold T from the same page (it was measured on this site).
   Fewer than T keys: the list is complete. T or more: show the keys and write "truncated".
5. Build the link https://<site>/issues/?jql=<the JQL, URL-encoded>.
Stop when every requested check has its block.
</procedure>
<output_format>
One block per check: "DCxx - <Name> - JQL match: exact | partial"
| JQL | Link | Items received | List complete | Keys |
Last line: "Exact count: open the link, or read DataChecks in the workbook."
</output_format>
<rules>
- Read only: you have no edit tool, and an edit would change the data the workbook measures.
- State no total you computed; the number people use comes from Jira or the workbook.
- Use only values from tool results and the rules page; write "not visible" when a field or item is missing.
- Quote keys, statuses and field names exactly, accents included.
</rules>
<fallback>
If the JQL fails, show it with the error, name the placeholder that is probably wrong, and go on to the next check.
</fallback>
```

**Plans readiness** (prompt P09) and **Backlog fixer** (P10, only after the rules are validated, with an edit tool,
never more than 20 items, always with confirmation) stay possible later as their own agents; keep them out of the
ART Flow Steering (`rovo_train_steering.md`), which has no write tool.

## 5. Skills (Beta) - one `SKILL.md` per family of checks

Revised 2026-11-05 (Documented, "Create and manage Rovo skills", first published 2026-10-25, beta): custom skills
are **personal** (shared only as copies through a link), made with `/build-skill` in Rovo Chat, in Studio, or by
uploading **one `SKILL.md`** (`name` and `description`; no scripts, zip or extra files). About 25 Atlassian skills
exist too. The tools a skill needs are added to the agent automatically and hidden from its Tools list. **A skill can
encode a method; it cannot add computation or API access** beyond the tools it lists: it does not make counts reliable.

Create each with Studio > *Create* > *Upload a skill* (one file at a time), then *Publish*. Template:

```text
---
name: dq-hierarchy-check
description: Check SAFe hierarchy links between features and team work items (DC01, DC02, DC03, DC24). Use when asked about missing parents, wrong parent level or orphan features.
---

# Hierarchy check

1. Read the project keys and field names from the Confluence page "ART data rules".
2. Search: project in (<TEAM_KEYS>) AND issuetype in (Story, Task) AND statusCategory != Done.
3. Apply rules DC01, DC02, DC03 exactly as written on the page. Then search features for DC24.
4. Output: one line per rule (id, items received, "complete" or "truncated", JQL link), then the table
   Check id | Work item key | Team | Field | Found value | Expected | Fix hint.
5. T or more items received (threshold T on the same page): the keys received, "truncated", and the JQL link.
   Never a total of your own: the exact count is in Jira issue search or in DataChecks.
6. Read only. Never edit. Never guess.
```

Make the same for: `dq-status-resolution-check` (DC04, DC05, DC10), `dq-estimation-check` (DC11, DC12),
`dq-team-sprint-check` (DC13-DC16), `dq-pi-dates-check` (DC17-DC20), `dq-business-value-check` (DC21),
`dq-dependency-check` (DC22, DC23), `dq-missing-data-report` (P08), `dq-plans-readiness` (P09).
Beta limits: personal skills, shared as copies; a copy does not follow later changes of the original.
So: keep the master text in Confluence (section 8), and make each agent work **without** the skills (its
Behavior points to the Confluence rules). Skills only make answers more consistent; they must not be required.

## 6. Prevent and watch: workflow rules, filters, automations (not beta)

Audit correction: an automation that **adds labels** to flag anomalies edits the work items. Each edit
changes `updated`, adds a history entry, can notify watchers and can start other rules. For checks that JQL
can already find, a **saved filter** gives the same list without touching the data. So:

1. **Detect with saved filters, not labels.** One saved filter per JQL-detectable check, named
   `DQ DC01 Story without feature`, ... (the JQL of section 3 of the metrics file and of the prompts above).
   Dashboard: "Filter results" gadgets, and a "Two dimensional filter statistics" gadget per filter
   (for example Project x Status). Exact counts, history-free, no Rovo needed. The full list of filters and the
   layout are in `jira_dashboard_data_quality.md`.
2. **Prevent in the workflow** (you are Jira admin; this is the most reliable layer):
   - transition to *En développement*: **validator** "field required" for story points and Sprint (DC11, DC15);
   - transition to *Planifié* (features): validator for Target start, Target end, PI (DC17, DC19);
   - transitions to *Terminé* / *Annulé*: set the resolution in a **post function** (DC04, DC05).
     *Dépriorisé* is not terminal (it goes to *Annulé* or back to *Nouveau*): it sets **no** resolution; if
     Jira sets one today, clear it on the move back to *Nouveau*.
   These validators are a **Proposal, to validate with the teams** before any is switched on: a blocked
   transition is visible friction.
3. **Warn with automations** only where a validator is too strict (event triggers, comment only, no label):

| Rule | Trigger | Condition | Action |
|---|---|---|---|
| R01 DC11 | Sprint field changed | issuetype = Story and `<POINTS_FIELD>` is empty | Comment to the assignee: "Please estimate before sprint start (DC11)" |
| R02 DC01 | Work item created in a team project | issuetype in (Story, Task) and parent is empty | Comment: "Link this item to its feature (DC01)" |
| R03 DC12 | `<POINTS_FIELD>` changed | issuetype is a sub-task | Comment: "Points on sub-tasks are counted twice (DC12)" |
| R04 DC21 | `<BV_FIELD>` changed | issuetype = `<FEATURE_TYPE>` | Comment: "Business value belongs to PI objectives (DC21)" |
| R05 weekly summary with Rovo | Scheduled, Monday 08:00 | Lookup work items with the JQL of the DC filters (one rule per family, to stay under the lookup size limit); continue only if the Jira-computed size is above 0 | **Use agent** (ART Flow Steering, `rovo_train_steering.md`), prompt below with the size and at most 50 keys; continue only if `{{agentResponse}}` is not empty; then email it to the RTE |

Settings for every rule: rule actor and notifications **Verify in UI** (avoid notifying on every comment);
multi-project scope needs a Jira admin; check the site's automation usage limit and the Rovo usage of
scheduled agent runs with your admin before enabling R05 weekly.

Prompt for R05 (smart values: **Verify in UI**):

```text
These {{lookupIssues.size}} work items match the data-quality filter "<FILTER NAME>" (rule <DCxx>):
{{#lookupIssues}}{{key}}; {{/}}
Group them by team. Compare with the rule <DCxx> on the page "ART data rules".
List the 5 most important fixes for the RTE, with the work item keys. Read only. Under 20 lines.
Exact list: open the filter "<FILTER NAME>" in Jira.
```

Revised 2026-11-05: the count in R05 comes from Jira (`{{lookupIssues.size}}`), never from the agent ("The count
then comes from Jira, not the model", Community). A Champion reports that putting the JQL in the scheduled trigger
instead of a lookup allows up to 999 work items (Community). Agents now act in automations unless blocked (ROVO-261):
give the ART Flow Steering no write tool rather than relying on a toggle (the toggle **Let this agent take actions**
was not found in the 2026-11 docs: Verify in UI). Run R05 with the Agent's account if the admins can grant it. Use
**Debug response** (the automation debug log shows subagent, agent and request, not the model) when an answer looks
wrong.

**Apps (Beta):** a generated app could show the same dashboard; keep it optional until it leaves beta.

## 7. Test your prompts and agents (Decided)

Rovo answers change when the model, the data or your text changes. Treat prompts like code:
1. Create a small **sandbox**: a Program project `<SBX_PROGRAM>` and a team project `<SBX_TEAM>` (or use a
   test filter) with these 20 work items (Decided). Each one carries one known anomaly, or none:

| # | Project | Type | Setup | Expected check |
|---|---|---|---|---|
| 1 | `<SBX_PROGRAM>` | Feature | clean: PI `PI#01-2026`, Target start and end, children below | none |
| 2 | `<SBX_PROGRAM>` | Feature | PI empty | DC17 |
| 3 | `<SBX_PROGRAM>` | Feature | status Planifié, Target end empty | DC19 |
| 4 | `<SBX_PROGRAM>` | Feature | Business Value set, has a child | DC21 |
| 5 | `<SBX_PROGRAM>` | Feature | clean but no child | DC24 |
| 6 | `<SBX_TEAM>` | Story | clean: parent #1, Team, sprint, points | none |
| 7 | `<SBX_TEAM>` | Task | clean: parent #1, Team, sprint | none |
| 8 | `<SBX_TEAM>` | Story | clean: parent #1, Team, sprint, points | none |
| 9 | `<SBX_TEAM>` | Story | no parent | DC01 |
| 10 | `<SBX_TEAM>` | Task | no parent | DC01 |
| 11 | `<SBX_TEAM>` | Story | parent is not a feature (only if the hierarchy allows it) | DC02 |
| 12 | `<SBX_TEAM>` | Story | Done, resolution date empty (may need a migration or a rule to create) | DC04 |
| 13 | `<SBX_TEAM>` | Story | resolution date set, status not Done (same remark) | DC05 |
| 14 | `<SBX_TEAM>` | Story | in an active sprint, points empty | DC11 |
| 15 | `<SBX_TEAM>` | Sub-task | under #6, points set | DC12 |
| 16 | `<SBX_TEAM>` | Story | Team empty | DC13 |
| 17 | `<SBX_TEAM>` | Story | In Progress, no sprint | DC15 |
| 18 | `<SBX_TEAM>` | Story | blocks #19, sprint ends after #19 starts | DC22 |
| 19 | `<SBX_TEAM>` | Story | blocked by #18 | DC22 |
| 20 | `<SBX_TEAM>` | Sub-task | under #7, no points, clean | none |

   Never put real client data in the sandbox. Checks that need the full history (DC06 to DC10, DC16, DC27 to
   DC29) are Excel only: they are not in this table.
2. Write the expected result per prompt: which keys, which check ids. Keep it on a Confluence page.
3. Run each prompt / agent on the sandbox after every change of the rules, the agent or a skill.
   Score: found / missed / false alarms. Keep the scores in a table (date, prompt version, score).
4. Use **Debug response** for automation runs, and the Studio audit log when a flow fails.
5. Load the prompts and their expected keys as a **Studio Evaluation** (Evaluation tab: a CSV of up to 100 prompts,
   mode Response accuracy, Documented). Add one **calibration filter** of 25 to 45 known work items: if the agent's
   list comes back cut, the measured size is the threshold T written on the "ART data rules" page. Check every number
   by hand against JQL (an LLM judge accepts a count that is merely close). Keep the CSV, the scores, the date and the
   **Language Model ID** (the `{ }` button) in `rovo/`, and run it again after any change of model, reasoning tier or
   instructions, one change at a time (Documented, "Model behavior guidelines").
6. Each month, compare the 10 exact checks of section 3 with the workbook's `DataChecks` on the same as-of date.
7. Only then run on the real projects.

## 8. Capitalising on all this

1. **One source of truth** (Decided): a Confluence page "ART data rules" with the DC table (copy section 3 of
   `safe-flow-metrics-and-data-checks.md`), the field names of this site, the status categories and allowed transitions, the PI
   calendar and the project/team list. Agents and skills read it as knowledge.
2. **Version the rules**: change log at the bottom of the page (date, rule, change, who).
3. **Same ids everywhere**: Excel `DataChecks`, Rovo outputs, saved filter names, automation comments,
   Confluence page.
4. **Weekly loop during the clean-up**: run the Excel extraction and checks -> compare counts with the
   saved filters and with the ART Flow Steering agent -> fix the gaps in the rules -> track counts per DC per week.
5. **Then per iteration**: R05 summary + Excel metrics before the iteration review; per PI: P08 and P09
   before PI planning.
6. **Keep a prompt log**: prompts that gave wrong or vague answers, with the fix. Promote good prompts to skills.

## 9. For Copilot: how the Excel side aligns with Rovo

- Use exactly the check ids `DC01`...`DC29` in `DataChecks[CheckId]`.
- Add a column `JiraJql` with `key in (A-1, A-2, ...)` (at most 100 keys per cell) so the user can paste it
  into Jira to see the same items.
- Add a sheet `RovoCompare` (Decided): CheckId | Excel count | Saved filter count (typed by the user) |
  Rovo count (typed by the user) | Rovo exact/estimated | Difference | Comment.
- Do not call Rovo from Excel. In this setup Rovo stays a human tool (Chat, Studio, automations).

Sources added 2026-11-05 (links in the research report): "Set a reasoning tier for a Rovo agent", "Model behavior
guidelines for Rovo Agents", "Knowledge sources for agents", "Evaluate the performance of your Rovo agent", "Rovo agent
permissions and governance", "Understand Rovo agent accounts", "How Rovo credits work", "Work with Jira Delivery
agent", "Use Rovo to search for work items", "JQL operators", "JQL fields", Forge `rovo:action` and `rovo:mcp`; staff
posts "A simpler way to build Rovo Agents in Studio", "Choose how Rovo Agents respond with subagents", "What's New in
Rovo Agents" (2026-06 to 2026-10); tracker ROVO-261, ROVO-396, ROVO-509, ROVO-535, ROVO-1148.

Sources (Atlassian support, 2026-11-03): "Create and edit Rovo agents", "Add skills to Rovo agents",
"Create and manage Rovo skills" (beta), "Automate Rovo agents" / "Automating Rovo agents", "Use Rovo in an
automation rule", "How to configure the Use Rovo agent action", "The new, unified Rovo Studio"
(Team '26 announcement: Apps in beta).
