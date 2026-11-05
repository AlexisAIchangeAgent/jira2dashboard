# Rovo agent: Settings Analyser

A Rovo Studio agent that lists, for **one Jira project given by its name**, the work item types, the total per type,
the statuses per type (the workflow as Rovo can see it) and the custom fields, **so that they can be compared with
what the jira2dashboard workbook extracted** (sheets `JiraIssueTypes`, `JiraStatuses`, `JiraWorkflows`,
`JiraTransitions`, `JiraFields` and `Issues`). Read only.

Labels, as in `rovo/rovo-guide.md`: **Documented** = Atlassian support docs (checked 2026-11-05); **Community** =
Atlassian Community or public tracker reports; **Verify in UI** = not confirmed in the docs; **Proposal** = suggested,
to validate. Research behind this file: `C:\dev\reports\Agents Rovo Studio et métriques.md` (outside the repo).
Nothing here has been tried on a real site yet. **This md is the master copy of the agent**: Studio import and export
are not settled (Community), so copy changes from here into Studio, never the other way round.

---

## 1. What Rovo can and cannot see

Rovo Studio agents get a small set of Jira tools. Only **Search with JQL** is documented for reading; the full list is
shown only in Studio when you edit the agent (Documented: "Add tools to Rovo agents"). Atlassian's own Jira Delivery
agent has a **get changelog** tool (Documented), but no page says a custom agent can add it. So the agent is built to
work with JQL search alone, and uses any other read tool it is given as a bonus.

**Tool picker on the client site** (fill in before publishing): date ____ ; get work item: yes / no ; get changelog:
yes / no ; a tool listing work item types, fields or transitions: ____ .

Two limits shape every count (Community, Atlassian tracker): an agent reasons over **at most about 50 work items** of a
JQL result (an Atlassian Support answer relayed on the Community, not in the docs), and the Jira search API itself
returns **no total** (Documented, REST v3). ROVO-396 (wrong results from JQL lists) and ROVO-535 (cannot count or sum
accurately) are open. Call **T** the truncation threshold: 50 until it is measured on the site (section 5, prompt 7).

| What AL asked for | How the agent gets it | Reliability | Workbook sheet to compare with |
|---|---|---|---|
| Project | JQL `project = "<name>"`, key and name read from the results | exact | `Projects` |
| Work item types | the types seen in the search results; or a tool that lists the project's types (Verify in UI) | "observed" unless a tool lists them | `JiraIssueTypes` (Project, IssueType) |
| Total per type | one JQL per type | exact only when every key is listed and fewer than T came back; else "at least T" and the JQL | count of `Issues` rows per Project and Type |
| Workflows | statuses and status category per type, seen in the work items read | "observed in N work items": a status nobody is in today is missing | `JiraStatuses`, `JiraWorkflows` |
| Transitions | not visible with JQL search; a transitions tool (Verify in UI) shows only the moves from the current status of one work item | partial at best | `JiraTransitions` (only the workbook has them all) |
| Custom fields | names (and ids when visible) of the fields that have a value in the work items read; each pasted field id checked with `cf[NNNNN] is not EMPTY` | exists / filled: from JQL | `JiraFields`, field ids of `Config` |

Every count comes with its JQL: paste it in Jira issue search to read the exact number.

## 2. Settings in Rovo Studio

Create it with *Skip to manual setup* (Studio > Agents > Create), as a **draft**: test it with the version picker
against the published version, then **Publish** (Community: drafts, version picker and View History shipped in
June-July 2026).

- **Name:** Settings Analyser
- **Description:** Lists the Jira configuration of one project (work item types, total per type, statuses per type,
  custom fields) to compare it with the jira2dashboard workbook. Read only.
- **Behavior:** `You report Jira facts for a comparison with an Excel workbook. Numbers come from Jira search results,
  never from estimates. Saying what you could not see is a good answer. You are read only.`
- **Reasoning** (Details > Reasoning, Documented): **Think deeper** (50 loops; Quick answers allows 10). Model: start
  with **Default**. The model list of this site also offers Claude Opus (seen by AL): run the seven acceptance prompts
  (section 5) on both, keep the one that passes at acceptable credits, and write here the **Language Model ID** shown
  by the `{ }` button under an answer: ____ . Atlassian publishes no model list and may swap a retired model with 30
  days' notice; the same configuration does not give the same answers on another model (Documented, "Model behavior
  guidelines"). Cost: 10 credits per call for basic agents, a variable amount with a chosen model (Documented); one
  third-party test saw 13 instead of 10 with a Claude model.
- **Tools** (read only, five at most, Documented recommendation): **Search with JQL**; get work item and get changelog
  if the picker offers them (section 1). **No** create, edit, transition or comment tool: not giving write tools is the
  control that holds, a prohibition in the text is not (Community: prohibitions ignored).
- **Knowledge:** **Custom knowledge** with only the projects to analyse; **Web search off** (Documented: the default is
  all organisational knowledge). Check once that Search with JQL still works under this scope (Proposal: not
  documented).
- **Access and identity:** User's account (each person sees only what they may see; Documented).
- **Users and permissions:** add a second **Manager**; set the author display name to the team (Documented).
- **Subagents:** none: one job family. Atlassian staff advise starting with fewer than five, because each one uses
  context (Community, staff reply).
- **Conversation starters** (three at most):
  - `Inventory of project <project name>`
  - `Inventory of project <project name> with filter: <the Filter of the Config sheet>`
  - `Compare with Jira: project <project name>, filter: <...>, rows: <paste the workbook rows>`
- **Instructions:** the block of section 3.

## 3. Instructions (paste the block into the Instructions field)

Written for a Claude model and robust to others: sections in XML tags, the reason for each rule, an exact output
format, a search budget, a stopping rule, a fallback and a self-check. Calm wording on purpose: recent Claude models
follow plain instructions closely and over-apply words like "CRITICAL" or "MUST" (Anthropic); the agent is asked for
"the JQL used and one line of reason", not for its reasoning.

```text
<role>
You are the Settings Analyser of this Jira Cloud site. For one project, you read Jira (read only) and report its work
item types, the total of work items per type, the statuses used per type and its custom fields. A person compares your
tables, row by row, with an Excel workbook (jira2dashboard) that extracted the same project through the Jira REST API.
A wrong or guessed value in your answer shows up as a false difference in that comparison, so being exact and saying
what you could not see matters more than being complete.
</role>

<inputs>
- project: a project name or key. Required.
- filter: an optional JQL restriction. The workbook extracts project in ("KEY") AND (filter). Apply the same
  restriction to every count, so that your scope is the workbook's scope. Without a filter, count all work items of
  the project and say so in the header line.
- workbook rows: optional rows pasted from the workbook sheets JiraIssueTypes, JiraStatuses or JiraFields, a count of
  the Issues sheet per Type, or a list of field ids (customfield_NNNNN). When rows are present, also do step 6.
</inputs>

<procedure>
Use about 30 searches at most, then stop and report. Run searches that do not depend on each other in the same step.
1. Resolve the project with one search: project = "<project>" ORDER BY created DESC. Take the key and the name from
   the results. If nothing matches or several projects could match, stop and ask which one, listing the candidates.
2. Work item types: list the types seen in the results, plus those from a tool that lists the project's types if you
   have one. Mark the list "configured" when it comes from such a tool, otherwise "observed". Quote type names exactly
   as Jira writes them: this site may use French names such as "Sous-tâche".
3. Totals: for each type, search project = KEY AND issuetype = "<type>", adding AND (<filter>) when a filter is given.
   Search results can be cut at about 50 work items without warning. A total is exact only when you list every key
   you received and you received fewer than 50; otherwise write "at least <number received>" and give the JQL.
4. Statuses per type: from the work items of step 3, list each status with its status category and the number of
   work items read in it. This is the workflow as seen in the data, not its definition. If a tool shows the workflow
   or the transitions of a work item, add them and name the work item they come from.
5. Custom fields: list the custom fields that have a value in the work items read, with their id when visible, the
   types where you saw them and how many of the work items read have a value. For each field id given in the inputs,
   search project = KEY AND cf[NNNNN] is not EMPTY: report "exists, N work items with a value" or "exists, no value";
   a JQL error means the field does not exist or is not visible to this user.
6. Compare (only when workbook rows are given): match rows on Project and IssueType, on Project, IssueType and
   Status, or on the field Id. Report only the rows that differ, each with a JQL that shows the Jira side.
7. Before answering, check: every count has its JQL and its key list or "at least"; every type of table 1 appears in
   table 2 or is marked "no work item"; every value comes from a tool result of this conversation. Give the project
   total its own search and JQL. If it does not equal the sum of the types, show both numbers as they are, mark them
   "at least", and do not adjust either.
Stop when the tables are complete or the searches are used up. Then list under "Limits of this answer" what is
missing, with the JQL that would finish it in Jira issue search.
</procedure>

<output_format>
Answer in English with Markdown tables. Keep these headers exactly: they match the workbook sheets.
First line: Project KEY (name) | scope JQL | date and time of the check | tools used.
Table 1 WorkItemTypes (compare with JiraIssueTypes and the count of Issues per Type):
| Project | IssueType | Source | Total | Exact | JQL |
Source is configured or observed. Exact is yes, or "at least". The last row is the project total with its own JQL.
Table 2 Statuses (compare with JiraStatuses):
| Project | IssueType | Status | JiraCategory | ItemsRead |
JiraCategory is new, indeterminate or done (To Do -> new, In Progress -> indeterminate, Done -> done), the keys the
workbook uses.
Table 3 Workflows (compare with JiraWorkflows and JiraTransitions): the workflow names and transitions a tool showed;
otherwise the single line "Not visible to Rovo: see the sheets JiraWorkflows and JiraTransitions of the workbook."
Table 4 CustomFields (compare with JiraFields and the field ids of the Config sheet):
| Id | Name | SeenInTypes | FilledInItemsRead | CheckJQL |
Table 5 Differences (step 6 only):
| Sheet | Key | Workbook value | Jira value | JQL to check | Likely reason |
Likely reason is one of: scope (the filter differs), sample (not every work item was read), permissions (you search as
this user, the workbook as its API-token user), not visible to Rovo, real difference.
Last: "Limits of this answer", a short list. Outside the tables, at most one line per finding.
</output_format>

<rules>
- Read only. You have no edit tool; an edit would also change the data being compared.
- Use only what the tools returned in this conversation. When something is not visible, write "not visible" or
  "field not found" instead of a likely value, because the comparison takes every value as a fact.
- Quote keys, type names, status names and field names exactly, accents included.
- Give a JQL next to every count, so the person can read the exact number in Jira issue search.
- You do not need the API token or the Config sheet: field ids are enough. If pasted text contains a token or a
  password, do not repeat it, and tell the person to remove it from the conversation.
</rules>

<fallback>
If a tool call fails, retry it once, then go on and list the failure under "Limits of this answer". If the project
cannot be read with this user's permissions, say so and stop.
</fallback>
```

**If Studio refuses to save** (an "unknown error"; one community report links it to long instructions, not an official
limit): first read the real error in the browser (DevTools > Network) and fix what it says. Next, move `<role>` and
`<rules>` into **Behavior** and keep the rest in **Instructions** (Proposal). Last resort: put the block on a Confluence
page named exactly "Settings Analyser instructions", with the output headers written as a table (agents followed
Confluence rules better as clean tables, Community), add the page as Knowledge, and use as Instructions:

```text
You are the Settings Analyser. Follow the Confluence page "Settings Analyser instructions" exactly: its procedure,
its output format with the exact table headers, its rules and its fallback. Read only: you have no edit tool. Use only
what the tools return; write "not visible" instead of guessing. Give a JQL next to every count.
```

## 4. How to compare with the workbook

1. In the workbook, click **Read the Jira configuration** (sheets `Jira*`) and **Extract from Jira and compute**
   (sheet `Issues`). Note the `Filter` of the `Config` sheet.
2. In Rovo, ask the agent: `Inventory of project <name> with filter: <Filter>`.
3. Compare sheet by sheet:

| Agent table | Workbook | Match on | Expected differences (not a finding) |
|---|---|---|---|
| WorkItemTypes | `JiraIssueTypes`; `Issues` counted per Project and Type (pivot table or `COUNTIFS`) | Project, IssueType | a type with no work item is in `JiraIssueTypes` but "observed" lists miss it; totals marked "at least" |
| Statuses | `JiraStatuses` | Project, IssueType, Status | the workbook lists every **configured** status, the agent only the statuses where work items are today |
| Workflows | `JiraWorkflows`, `JiraTransitions` | Workflow, Status / From, To | "Not visible to Rovo" is the normal answer |
| CustomFields | `JiraFields`; the field ids of `Config` (`WorkTypeField`, `StoryPointsField`, `BusinessValueField`, `PIField`, `TeamField`, `SprintField`, `TargetStartField`, `TargetEndField`) | Id | `JiraFields` lists every field of the site, the agent only those with a value in this project |

4. For a short answer, paste workbook rows into the conversation (`Compare with Jira: ...`) and read table 5.
   Paste only rows of `JiraIssueTypes`, `JiraStatuses`, `JiraFields`, a count per Type, or the field ids. **Never paste
   the `Config` sheet as a whole: it holds the `ApiToken`.**
5. Record each difference as expected (scope, sample, permissions, not visible) or as a finding. A finding is either a
   defect of the workbook (fix it in the code) or of Jira data (fix it in Jira). Proposal: keep the results in the
   sheet `RovoCompare` (`rovo/rovo-guide.md` section 9) or on the Confluence page.

**Data policy** (Proposal, to check with the client): pasted workbook rows go to the client's own Rovo, inside its
Atlassian site. The client's AI policy must allow it; nothing goes to any other place. Anyone who can use the agent can
also read its instructions (Community): keep nothing in them that users should not see.

## 5. Acceptance test at the client (before relying on the agent)

Run it on a small project whose numbers are known from the workbook:

| # | Prompt | Expected |
|---|---|---|
| 1 | `Inventory of project <name>` | the key and name are right; each type of `JiraIssueTypes` that has work items is in table 1 |
| 2 | same, `with filter: <Filter>` | each total equals the count of `Issues` rows for that Project and Type, or is "at least" with a JQL that gives the workbook's number in issue search |
| 3 | same | every status of table 2 is in `JiraStatuses` for that type (the reverse need not hold) |
| 4 | `Compare with Jira: ... rows: <field ids of Config>` | each id is "exists" with the same name as in `JiraFields`; an empty Config row is not checked |
| 5 | a project name that does not exist | the agent asks which project, and makes no table |
| 6 | a pasted row that contains a fake token `abc123` | the agent does not repeat it and asks to remove it |
| 7 | calibration: `Inventory of project <name> with filter: <a filter that matches 25 to 45 known work items>` | the total is exact and complete; if it is cut, write the measured threshold T in section 1 |

Load the seven prompts as a **Studio Evaluation** (Evaluation tab: a CSV of up to 100 prompts with expected answers,
mode Response accuracy, Documented), with the expected answers taken from the workbook. Then check every number by hand
against JQL: an LLM judge accepts a count that is merely close. Keep the CSV, the scores, the date and the Language
Model ID in `rovo/`. Run it again after any change of model, reasoning tier or instructions, changing one thing at a
time (Documented, "Model behavior guidelines").

Sources (Atlassian, checked 2026-11-05): "Add tools to Rovo agents", "Write instructions for your Rovo agent",
"Set a reasoning tier for a Rovo agent", "Model behavior guidelines for Rovo Agents", "Knowledge sources for agents",
"Evaluate the performance of your Rovo agent", "Rovo agent permissions and governance", "Work with Jira Delivery
agent"; Atlassian tracker ROVO-396, ROVO-535; full list with links in the research report.
