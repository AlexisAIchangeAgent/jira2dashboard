# Rovo agent: Flow Metrics by work item type

A Rovo Studio agent that gives, for **one Jira project and one iteration**, the six SAFe flow metrics **for each work item
type** (Story, Task, Bug, Feature...), with the data anomalies and missing fields that make them wrong. It is a
complement to the jira2dashboard workbook, **not its replacement**: the workbook reads the full change history and
computes every figure exactly with Excel formulas; Rovo cannot do that (section 1). The agent gives what Rovo can give
honestly, labels each number exact, approximate or sample, and puts a JQL next to each one. The instructions are
written for a Claude model (Think deeper).

Labels, as in `rovo/rovo-guide.md`: **Documented** = Atlassian or Scaled Agile docs (checked 2026-11-05);
**Community** = Atlassian Community or public tracker reports; **Verify in UI** = not confirmed in the docs;
**Proposal** = suggested, to validate. Research behind this file: `C:\dev\reports\Agents Rovo Studio et métriques.md`
(outside the repo). Nothing here has been tried on a real site yet. This md is the master copy of the agent.

The research concludes that **no official flow figure should come from a Rovo agent**. This agent therefore gives only
numbers it can stand behind (small exact counts, a sample it shows item by item), each labelled, and points to the
workbook for the rest. For the data checks and the steering cycle (explaining the figures the RTE publishes), the
agent "ART Flow Steering" (`rovo/rovo_train_steering.md`) is the better fit.

---

## 1. What a Rovo agent can and cannot compute

Facts that bound the design:
- Atlassian: Rovo search is "not a good fit" to "analyze the work items to create charts, summaries, or other
  representations of your data" (Documented, "Use Rovo to search for work items").
- An agent reasons over **at most about 50 work items** of a JQL result (Atlassian Support, relayed on the Community,
  2026-01; not in the docs). Open suggestions on Atlassian's tracker: ROVO-396 (wrong results from JQL lists), ROVO-535
  (cannot count or sum accurately), ROVO-509 (cannot aggregate number fields) (Community / tracker).
- JQL never returns a duration or a time in status. `WAS` and `CHANGED` work on Status and Resolution (and Assignee,
  Fix Version, Priority, Reporter) only; Sprint and parent have no history operators (Documented, JQL reference).
- The status history (changelog) of a work item: Atlassian's own Jira Delivery agent has a **get changelog** tool, one
  work item per call (Documented). Whether a custom Studio agent can add it: **Verify in UI**.
- Think deeper allows at most 50 loops per answer, Quick answers 10 (Documented).

| Metric (per type) | Verdict | How the agent gets it | Official figure |
|---|---|---|---|
| Flow Velocity | **exact up to 50 items, else a JQL link** | one JQL per type for the iteration | workbook `Metrics` / `ItemFlow` |
| Flow Distribution | **exact when every velocity is exact** | share of each type in the done items, from the velocity counts | workbook (by work type `TypeCategorie`) |
| Flow Load | **exact today, approximate on a past day** | `status in (<WIP>)` now; `status WAS IN (<WIP>) ON "<day>"` for a past day | workbook `DailyLoad` (modes A and B) |
| Flow Time | **sample only, and only with a changelog tool** | changelog of up to 20 done items of one type; median and P85 of that sample | workbook (all items, median and 85th percentile) |
| Flow Efficiency | **sample only, and only with a changelog tool** | same sample, days in Active and Wait statuses | workbook (`not computable` while a status is Unclassified) |
| Flow Predictability | **not computable** | Jira holds no PI objectives with planned and actual business value | workbook shows `not computable` too |

So the agent answers in two jobs: an **overview** of all types (velocity, distribution, load, and the checks that
break them), and a **flow time sample** for one type on request (flow time and efficiency).

## 2. Inputs (from the workbook parameters, never client secrets)

- **Project** name or key, and **iteration**: start and end dates (an iteration or a PI of the `Iterations` / `PIs` sheets).
  The workbook computes in UTC, Jira in the user's time zone: a few items can move across a iteration boundary.
- **Filter** (optional): the `Filter` of the `Config` sheet, so the scope is the workbook's.
- **Status classification**: the rows of the `Workflows` sheet (Workflow, Status, Category, IsStart, IsEnd). Either pasted
  in the conversation, or kept on the Confluence page "ART data rules" added as Knowledge (`rovo/rovo-guide.md`
  section 8), written as a clean table. Without it the agent cannot tell start, end, active and wait statuses, and
  says so.
- **Field names** (optional, for the checks): story points, work type (`WorkTypeField`), Team, PI.
- Never paste the `Config` sheet as a whole: it holds the `ApiToken`.

## 3. Settings in Rovo Studio

Create it with *Skip to manual setup* (Studio > Agents > Create), as a **draft**; test it with the version picker,
then **Publish**.

- **Name:** Flow Metrics by Type
- **Description:** Gives, for one project and one iteration, the SAFe flow metrics per work item type, each marked exact,
  approximate or sample, with the data anomalies that break them and a JQL for every number. Read only.
- **Reasoning** (Details > Reasoning, Documented): **Think deeper** (50 loops; the flow time sample needs about 22 tool
  calls). Model: start with **Default**; the list of this site also offers Claude Opus (seen by AL): run the acceptance
  prompts (section 7) on both, keep the one that passes at acceptable credits, and write the **Language Model ID**
  shown by the `{ }` button here: ____ . A model change can change the answers (Documented, "Model behavior
  guidelines"). Cost: 10 credits per call for basic agents, a variable amount with a chosen model (Documented).
- **Tools** (read only; Atlassian recommends five at most): **Search with JQL** (Documented); **get changelog** and **get
  work item** if Studio lists them (Verify in UI; get changelog is documented for the Jira Delivery agent only). No
  create, edit, transition or comment tool: not giving write tools is the control that holds.
- **Knowledge:** Custom knowledge: the Jira projects to analyse, and the Confluence page "ART data rules" with the
  status table of `Workflows`. Web search off. Check once that Search with JQL still works under this scope (Proposal).
- **Access and identity:** the user's account (each person sees only what they may see).
- **Users and permissions:** a second Manager; the author display name set to the team.
- **Subagents:** none at first. Proposal, if answers mix the two jobs: a subagent "Flow time sample" (trigger: the user
  asks for flow time, cycle time or efficiency of one type), mode Forward as-is.
- **Conversation starters:**
  - `Flow metrics of project <name>, iteration <start> to <end>`
  - `Flow metrics of project <name>, iteration <start> to <end>, filter: <Filter>`
  - `Flow time sample of type <type> in project <name>, iteration <start> to <end>`
- **Instructions:** the block of section 4. Then **Publish**.

## 4. Instructions (paste the block into the Instructions field)

```text
<role>
You are the Flow Metrics assistant of this Jira Cloud site. For one project and one iteration, you give the SAFe flow
metrics for each work item type, and the data problems that make them wrong. An Excel workbook (jira2dashboard)
computes the official figures from the full change history; people compare your answer with it. Rovo reads at most
about 50 work items of a search and cannot count large sets reliably, so a number you cannot stand behind does more
harm than a JQL link. Label every number exact, approximate or sample, and give its JQL.
</role>

<inputs>
- project (name or key) and iteration (start date, end date). Ask for them if missing.
- filter: optional JQL restriction; add it as AND (filter) to every search, so the scope is the workbook's.
- status classification: the rows Workflow | Status | Category | IsStart | IsEnd of the workbook sheet Workflows, pasted
  or on the knowledge page "ART data rules". Category is Backlog, Active, Wait, Done or Removed. Start statuses have
  IsStart = 1, end statuses IsEnd = 1, work in process (WIP) statuses are the Active and Wait ones. If you have no
  classification, say so and give only Velocity by resolution date, the current Load by status category In Progress,
  and the checks.
</inputs>

<job name="overview">
Default job. Use about 30 searches at most, then stop and report; run searches that do not depend on each other in the
same step.
1. Resolve the project with one search; ask if several projects match. List its work item types from the results.
2. For each type T, run:
   - Velocity: project = KEY AND issuetype = "T" AND status in (<end statuses>) AND status CHANGED TO "<end status>"
     DURING ("<start>", "<end>") AND NOT status CHANGED TO "<end status>" AFTER "<end>", one pair of CHANGED clauses
     per end status. An item counts once, in the iteration of its last entry into an end status, and only if it is
     still in an end status (the workbook's rule). Call this search V(T).
   - Load today: project = KEY AND issuetype = "T" AND status in (<WIP statuses>).
   - Load at the iteration end: project = KEY AND issuetype = "T" AND status WAS IN (<WIP statuses>) ON "<end>".
3. Search results can be cut at about 50 work items without warning. A count is exact only when you list every key
   you received and you received fewer than 50. Otherwise write "at least <number received>" and the JQL: the exact
   number is in Jira issue search.
4. Distribution: the share of each type in the done items of the iteration, computed only from exact velocities; show
   the counts it comes from. If one velocity is not exact, write "not computable here" and give the JQLs.
5. Checks per type (each one JQL; report the count the same way as in step 3):
   - done without resolution: status in (<end statuses>) AND resolution IS EMPTY (breaks velocity by resolution date)
   - resolution but not done: status not in (<end statuses>) AND resolution IS NOT EMPTY
   - done in the iteration without a start: V(T) AND status WAS NOT IN (<start statuses>) (left out of flow time)
   - done in the iteration after a reopening: V(T) AND status CHANGED FROM "<end status>", one clause per end status
     (left out of flow time)
   - status outside the classification: status not in (<all classified statuses>) (makes efficiency not computable)
   - in progress without sprint: status in (<WIP statuses>) AND sprint IS EMPTY
   - missing work type, story points or Team, when their field names are given: "<field>" IS EMPTY
   - stuck: status in (<WIP statuses>) AND statusCategoryChangedDate <= -14d
6. Flow Time and Flow Efficiency: not in the overview. Write "ask: Flow time sample of type <T>" for each type with
   done items. Flow Predictability: "not computable: Jira holds no PI objectives with planned and actual business value".
</job>

<job name="flow time sample">
When the person asks for flow time, cycle time or efficiency of one type.
1. Run the velocity search of the overview for that type; take at most 20 done work items (the most recently done).
2. If you have a changelog tool, read the status history of each one. If you do not, say "not computable by Rovo
   without a changelog tool: see the workbook" and stop.
3. For each item: start = its first entry into a start status after its last entry into a Backlog status (creation
   counts as an entry into a Backlog status); done = its last entry into an end status. Flow time = done - start in
   calendar days, with one decimal. Leave the item out, and say why, when it has no start (done without a start) or
   when it left an end status at some point (reopened). Active days and wait days = the days spent in Active and Wait
   statuses between start and done.
4. Flow time of the sample: sort the flow times, give the median, and the 85th percentile computed like Excel
   PERCENTILE.INC (rank 1 + 0.85 x (n - 1), linear between the two neighbours). Efficiency of the sample = sum of active
   days / sum of active and wait days; "not computable" if a status of these items is not in the classification.
5. Always show the per-item table, so a person can check every date and every subtraction.
</job>

<output_format>
Answer in English. First line: Project KEY (name) | iteration | scope JQL | date of the check | tools used.
Overview, table FlowByType:
| IssueType | Velocity | VelocityExact | DistributionPct | LoadToday | LoadAtIterationEnd | LoadExact | VelocityJQL |
Then table Checks:
| IssueType | Check | Breaks | Count | Exact | Keys (up to 50) | JQL |
Then one line each for Flow Time, Flow Efficiency and Flow Predictability, as described in the overview job.
Flow time sample, table FlowTimeSample:
| Key | Start | StartStatus | Done | DoneStatus | FlowTimeDays | ActiveDays | WaitDays | LeftOut |
then: n, median, P85, efficiency, each marked "sample of n items, not the official figure".
Last: "Limits of this answer": what you could not read or count, with the JQL to finish it.
</output_format>

<rules>
- Read only: do not create, edit, transition or comment on any work item.
- Every number comes from a tool result of this conversation, and every count has its JQL. Do not estimate a count from
  a partial list; write "at least N". Never adjust one number to make it agree with another.
- Quote keys, type names, status names and field names exactly, accents included. Dates as YYYY-MM-DD; say that Jira
  uses your time zone while the workbook uses UTC, so a few items can move across a iteration boundary.
- When your figure and the workbook's differ, the workbook is the reference: name the likely reason (scope, sample,
  50-item limit, time zone, a check above) instead of choosing a side.
- If pasted text contains a token or a password, do not repeat it and ask the person to remove it.
</rules>

<fallback>
If a search fails, retry once, then go on and list it under "Limits of this answer". If a JQL is refused (for example
a field name unknown on this site), say which one and continue with the others.
</fallback>

<self_check>
Before answering: each velocity marked exact has its keys listed or a reported total; the distribution adds up to
100 % or is "not computable here"; each flow time in the sample equals done - start; no figure is presented as
official.
</self_check>
```

**If Studio refuses to save the block:** put it on a Confluence page "Flow Metrics instructions", add the page as
Knowledge, and use as Instructions: `You are the Flow Metrics assistant. Follow the Confluence page "Flow Metrics
instructions" exactly. Read only. Label every number exact, approximate or sample, and give its JQL.`

## 5. Anomalies and missing fields: what each check breaks

The checks use the same ids as the workbook (`DataChecks`, `docs/safe-flow-metrics-and-data-checks.md` section 3).
The coverage follows `rovo/rovo-guide.md` section 3, which was checked against `src/jira2dashboard/flow/checks.py`:
- 10 checks have an exact JQL;
- 9 have only a partial one;
- DC01 depends on a syntax test (`parent IS EMPTY`);
- 9 need the change history or a comparison between items, and stay in the workbook only: DC02, DC03, DC06, DC07,
  DC18, DC20, DC24, DC27 and DC29.

| Check in the agent | Workbook id | Breaks | JQL status |
|---|---|---|---|
| done without resolution | DC04 | Velocity by resolution date | Documented operators |
| resolution but not done | DC05 | Velocity by resolution date | Documented operators |
| done without a start | DC09 (partly: the workbook also restarts after a return to Backlog) | Flow Time | Documented operators |
| reopened | DC10 (partly: exact only with one Done status per workflow) | Flow Time (item left out) | Documented operators |
| status outside the classification | DC08 (current status only) | Flow Efficiency | Documented operators |
| in progress without sprint | DC15 | Flow Load by sprint, planning | Documented operators |
| missing story points / Team / work type | DC11, DC13, Distribution `Unmapped` | Velocity in points, team split, Distribution | Test `IS EMPTY` on the Team field first |
| stuck | none (a watch, not a DC) | Flow Time outliers | `statusCategoryChangedDate` keeps only the last category change: moves between two In Progress statuses are missed |

To test on a sandbox before trusting them (Verify in UI): `CHANGED TO ... DURING` with several end statuses joined
with OR, `NOT status CHANGED ... AFTER`, `parent IS EMPTY` (not documented, reported to work), `IS EMPTY` on the Team
field, and whether **get changelog** appears in the Studio tool list.

## 6. How to compare with the workbook

1. In the workbook, run **Extract from Jira and compute** with the same `Filter`. Note the iteration.
2. Ask the agent: `Flow metrics of project <name>, iteration <start> to <end>, filter: <Filter>`.
3. Compare:

| Agent | Workbook | Expected differences (not a finding) |
|---|---|---|
| Velocity per type | `ItemFlow`: count of rows with that Type and `DoneUtc` in the iteration (`COUNTIFS`) | "at least N" above about 50 items; a few items at iteration edges (time zone); items whose status name changed since; items the chatting user cannot see (permissions) |
| Distribution per type | the same counts | the workbook's Distribution is by work type (`TypeCategorie`), the agent's by issue type |
| Load at iteration end | `DailyLoad` of the last day | the workbook's mode A counts from the first start and skips Backlog days; JQL `WAS IN ... ON` counts any WIP status held that day |
| Flow time sample | `ItemFlow[FlowTimeDays]` of the same keys | must be equal key by key (same rule); the median of 20 items differs from the median of all items |
| Checks | `DataChecks` rows of the same id | the agent's DC08 checks the current status only; DC09 does not restart after Backlog |

A key-by-key difference in the flow time sample is a finding: either the agent misread the history (check the
changelog in Jira) or the workbook rule differs (check `ItemFlow` and `StatusIntervals` of that key).

## 7. Acceptance test at the client

| # | Prompt | Expected |
|---|---|---|
| 1 | overview on a small project, with the `Workflows` rows pasted | each type with done items has a velocity; exact ones equal the workbook count |
| 2 | same on a project with more than 50 done items of one type | "at least ..." and a JQL whose count in issue search equals the workbook's |
| 3 | `Flow time sample of type Story ...` | per-item table; each FlowTimeDays equals `ItemFlow[FlowTimeDays]` of the same key |
| 4 | same without a changelog tool | "not computable by Rovo without a changelog tool" and no number |
| 5 | overview without the status classification | the agent says what it cannot do and gives only the fallback figures |
| 6 | a pasted text with a fake token `abc123` | not repeated; asks to remove it |
| 7 | calibration: overview with a filter that matches 25 to 45 known done items | the velocity is exact and complete; if it is cut, note the measured threshold here |

Load the seven prompts as a **Studio Evaluation** (a CSV of up to 100 prompts with expected answers taken from the
workbook, mode Response accuracy, Documented), then check every number by hand against JQL and `ItemFlow`: an LLM
judge accepts a count that is merely close. Keep the CSV, the scores, the date and the Language Model ID in `rovo/`.
Run it again after any change of model, reasoning tier or instructions, one change at a time (Atlassian: "Do not
assume a model switch preserves behavior", Documented).

## 8. The exact path (Proposal, not built)

Only code gives exact flow time, load and efficiency inside Rovo: a Forge app with a `rovo:action` (GET) that calls
`POST /rest/api/3/changelog/bulkfetch` (up to 1000 work items per call), computes the intervals and the metrics with
the workbook's rules, and returns a small summary (action data is limited to 5 MB; inputs are strings or numbers;
exposed to Studio agents through `rovo:mcp`, in preview) (Documented, Forge `rovo:action` and REST v3). It would
re-implement `src/jira2dashboard/flow/` and need its own tests against the same oracle, and an admin must install the
app on the client's site. Until there is a need for figures without Excel, the workbook stays the source of the
official numbers and this agent stays the quick, labelled view.
