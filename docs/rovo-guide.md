# Rovo guide: prompts, agents, skills and automations for Jira data quality

For: **you** (Jira admin, using Rovo Chat and Rovo Studio) and **Copilot** (which builds the Excel side).
Goal: find and fix the data anomalies that make the SAFe flow metrics wrong and keep the Jira **Plans**
program board unusable, then **prevent** them from coming back.

The check ids `DC01`...`DC29` are the same as in `safe-flow-metrics-and-data-checks.md` and in the Excel
`DataChecks` sheet. Always use them: it lets you compare Rovo, saved filters and Excel counts.

Labels: **Documented** = Atlassian support docs (checked 2026-11-03); **Beta** = marked beta by Atlassian;
**Verify in UI** = exact label not confirmed in the docs; **Proposal** = suggested, to validate.

---

## 1. What Rovo can and cannot do (Documented)

- **Agents** (Studio > Agents) - available, not beta. Setup: *Describe it* (guided chat) or *Skip to manual
  setup*. Fields: **Name**, **Description**, **Behavior** (instructions applied in every situation),
  **Conversation starters**, then **Subagents**, each with **Instructions**, **Skills**, **Tools**
  (Atlassian recommends no more than five), **Knowledge** (Confluence spaces, Jira projects...) and
  **Reasoning** (depth vs speed and cost). Select **Activate**. Only the creator or an org admin can edit.
- **Skills** - **Beta**. Reusable multi-step instructions that orchestrate tools. Create in Studio
  (*Start from scratch* > *Skills*: Name, Summary, Description, Instructions, Tools > **Publish**), in Rovo
  Chat with `/build-skill`, or **upload one `SKILL.md`** (Agent Skills format: needs `name` and `description`;
  no scripts, no zip). Skills are **personal**: share a *copy* with *Copy link*; the receiver selects
  *Duplicate skill*. Add skills to an agent: agent *Details* > *Edit* > *Skills* > *Add skills*.
- **Automations** - available, not beta. Action **Use agent** (select an agent + write a prompt); the
  output is the smart value `{{agentResponse}}`. Action **Use Rovo** (no agent) gives `{{rovoResponse}}`.
  Triggers include work item created and **scheduled** runs. Out-of-the-box agents **do not take actions**
  in automations: use `{{agentResponse}}` in the next action (comment, email, page).
- **Use agent in a Jira automation rule** (Documented, "How to configure the Use Rovo agent action"):
  the toggle **Let this agent take actions** is **off by default**; the agent runs either as the
  **requesting user** (their permissions) or as its own **agent account** (permissions managed by admins);
  smart values such as `{{issue.description}}` and `{{issue.url}}` in the prompt give more predictable
  answers; `{{agentResponse}}` is Markdown by default; **Debug response** (on the action, More actions)
  shows the tools used and the changes made.
- **Apps** - **Beta**: custom apps generated from a prompt.
- **Permissions**: agents and skills always act with the permissions of the person using them.
  Bulk actions in Jira are limited to **20 work items at a time**. The agent asks for confirmation before
  consequential actions.

Consequence: use Rovo to **find, explain and propose**; use JQL (issue search) + Excel to **count exactly**;
use workflow validators and Jira automation to **prevent**. Do not ask Rovo to compute metrics.
**Counts from Rovo are not reliable**: its answers come from search results it reads, not from a database
count. Every prompt asks it to say whether a count is exact or estimated, and to give the JQL: open that
JQL in Jira issue search to get the exact number.

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
  3. If more than 50 items match a rule: give the count, the first 50 keys, and the JQL to see all.
CONSTRAINTS: Read only. Do not edit anything. Use only data you can see. If a field does not exist or is
  hidden, say so instead of guessing. Quote work item keys exactly.
```

Placeholders to fill once (keep them in a Confluence page "ART data rules", see section 7):
`<PROGRAM_KEY>`, `<TEAM_KEYS>` (comma list), `<POINTS_FIELD>` (for example `Story point estimate`),
`<TEAM_FIELD>` (`Team`), `<PI_FIELD>` (`PI`, values like `PI#XX-YYYY`: `XX` = 2-digit PI number, `YYYY` = year), `<BV_FIELD>` (`Business Value`), `<TARGET_START>` / `<TARGET_END>`
(`Target start` / `Target end`), `<FEATURE_TYPE>` (`Feature` or its French name on the site).
Field names in JQL differ per site: get the exact names with prompt P00 or from the Excel `FieldCatalog`.

## 3. Prompt library (copy, fill the placeholders, run in Rovo Chat)

Not covered by Rovo: **DC06, DC07, DC08, DC09** (history mismatch, change before creation, status not in
the workflow sheet, done without a start status), **DC27** (moved issues), **DC28** (transition not allowed) and **DC29** (PI format). They need the complete change log
and the workflow sheets:
only the Excel workbook checks them. Rovo can explain them, not count them.

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

## 4. Agent to create in Rovo Studio (Decided: one agent with subagents)

Create it with *Skip to manual setup*. Start **read-only**: no edit tools until the rules are stable.
One agent "ART Data Quality" with a **default subagent** (the auditor) and subagents for Plans readiness and
metrics coaching. Studio uses a subagent only when its situation matches, else the default. Keep each
Behavior text short: the rules live on the Confluence page, not in the agent.

### Agent: ART Data Quality

- **Name:** ART Data Quality
- **Description:** Finds Jira data anomalies (DC01-DC29) that break SAFe flow metrics and the Plans program board. Read only.
- **Behavior:**

```text
You audit Jira data for the SAFe Agile Release Train "<ART NAME>".
Program project: <PROGRAM_KEY> (features). Team projects: <TEAM_KEYS> (stories, tasks, sub-tasks).
Always follow the rules in the Confluence page "ART data rules" (knowledge). Each rule has an id DCxx.
Always answer with: a summary line per rule (id, count, % of scope), then a table
Check id | Work item key | Team | Field | Found value | Expected | Fix hint.
If more than 50 items match, give the count, the first 50 keys and the JQL to see all.
Never edit work items. Never guess a field value: say "not visible" or "field not found".
Use exact field names from the "ART data rules" page. Quote keys exactly. State the date of the check.
```

- **Conversation starters:** "Run the hierarchy checks (DC01-DC03)", "Which data is missing for flow
  metrics?", "Is the plan ready?", "Show estimation gaps per team".
- **Knowledge:** Jira projects `<PROGRAM_KEY>`, `<TEAM_KEYS>`; Confluence page "ART data rules".
- **Skills:** the skills of section 5 once created. **Tools:** Jira search / read only (max five tools).
- **Reasoning:** higher depth (accuracy matters more than speed).

### Subagent 1 (default): Data Quality Auditor

- **Instructions:** the Behavior above, plus the prompts P01 to P08 as the way to answer.

### Subagent 2: Plans Readiness Checker

- **Situation:** the user asks if a Jira plan can show a reliable timeline, hierarchy, capacity and dependencies.
- **Instructions:** prompt P09 rules, plus: "Explain each blocker in one sentence a team member understands."
- **Knowledge:** same projects + the "ART data rules" page + a page describing the plan settings
  (issue sources, hierarchy levels, date fields used, teams and their velocity).

### Subagent 3: Flow Metrics Data Coach

- **Situation:** the user asks about a SAFe flow metric or the data it needs.
- **Instructions:**

```text
You explain SAFe flow metrics (distribution, velocity, time, load, efficiency, predictability) using the
definitions on the "ART data rules" page. For a question about a metric: say which Jira data it needs,
which checks (DCxx) protect it, and whether the data exists today (run the matching checks).
Never give a metric value: the official values come from the Excel workbook. Point to it.
```

### Subagent 4 (later, optional): Backlog Fixer

- Only after the rules are validated. Instructions = prompt P10. Add the Jira edit tool.
- Always asks for confirmation; never more than 20 items per run (Atlassian limit for bulk actions).

## 5. Skills (Beta) - one `SKILL.md` per family of checks

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
4. Output: summary per rule (id, count, %), then the table
   Check id | Work item key | Team | Field | Found value | Expected | Fix hint.
5. More than 50 matches: count + first 50 keys + JQL.
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
   (for example Project x Status). Exact counts, history-free, no Rovo needed.
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
| R05 weekly summary with Rovo | Scheduled, Monday 08:00 | Lookup work items with the JQL of the DC filters (one rule per family, to stay under the lookup size limit) | **Use agent** (ART Data Quality), prompt below, then email `{{agentResponse}}` to the RTE |

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

Keep the toggle **Let this agent take actions** off for R05. Use **Debug response** when an answer looks wrong.

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
5. Only then run on the real projects.

## 8. Capitalising on all this

1. **One source of truth** (Decided): a Confluence page "ART data rules" with the DC table (copy section 3 of
   `safe-flow-metrics-and-data-checks.md`), the field names of this site, the status categories and allowed transitions, the PI
   calendar and the project/team list. Agents and skills read it as knowledge.
2. **Version the rules**: change log at the bottom of the page (date, rule, change, who).
3. **Same ids everywhere**: Excel `DataChecks`, Rovo outputs, saved filter names, automation comments,
   Confluence page.
4. **Weekly loop during the clean-up**: run the Excel extraction and checks -> compare counts with the
   saved filters and with the ART Data Quality agent -> fix the gaps in the rules -> track counts per DC per week.
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

Sources (Atlassian support, 2026-11-03): "Create and edit Rovo agents", "Add skills to Rovo agents",
"Create and manage Rovo skills" (beta), "Automate Rovo agents" / "Automating Rovo agents", "Use Rovo in an
automation rule", "How to configure the Use Rovo agent action", "The new, unified Rovo Studio"
(Team '26 announcement: Apps in beta).
