# Rovo agent: ART Flow Steering

A Rovo Studio agent for the steering cycle of the Agile Release Train (ART). It does three jobs:
- it checks the Jira data before an extraction;
- it explains the flow metrics the RTE published;
- it drafts a steering note.

The **workbook stays the engine**. The jira2dashboard workbook reads the full change history and computes every
figure. The RTE publishes those figures on a Confluence page, and the agent reads that page. It never computes a
metric and never states a total it computed.

Labels, as in `rovo/rovo-guide.md`:
- **Documented**: Atlassian docs, checked 2026-11-05.
- **Community**: Atlassian Community posts or public tracker reports.
- **Verify in UI**: not confirmed in the docs.
- **Proposal**: suggested, to validate.

Research behind this file: `C:\dev\reports\Agents Rovo Studio et métriques.md` (outside the repo). Nothing here has
been tried on a real site yet. This md file is the master copy of the agent.

How it relates to the other agents of this folder:
- **Replaces** the Flow Data Steward (`rovo/rovo-guide.md` section 4) for the steering cycle. That agent's Item
  investigator can come back later as a fourth job, if the tool picker offers a changelog tool.
- **Settings Analyser** (`rovo_settings_analyser.md`) stays separate. It runs once per client, then each quarter.
- **Flow Metrics by Type** (`rovo_flow_metric.md`) stays the quick, labelled view between two extractions.

---

## 1. The steering cycle

| # | Step | Who | Where | When |
|---|---|---|---|---|
| 0 | Jira configuration matches the workbook parameters | Settings Analyser agent | Rovo Chat | once per client, then each quarter |
| 1 | Data ready: data checks | this agent, job `readiness`; Jira dashboard `jira_dashboard_data_quality.md` | Rovo Chat, Jira | before each extraction |
| 2 | Extract and compute | the RTE, button **Extract from Jira and compute** | the workbook, on the client machine | end of iteration, end of PI |
| 3 | Publish | the RTE, from `confluence_flow_metrics_page.md` | Confluence, one page per PI | after step 2 |
| 4 | Explain the figures | this agent, job `explain` | Rovo Chat | on request |
| 5 | Steering note | this agent, job `steering`; the RTE reviews it and pastes it into the page | Rovo Chat, then Confluence | before an iteration review, Inspect and Adapt, PI planning |

Step 3 stays a manual copy: an agent cannot yet place a page under a given parent page (section 7). Why step 2
stays outside Rovo (sources: `rovo/rovo-guide.md` section 1, `rovo_flow_metric.md` section 1):
- **Extraction:** an agent reasons over at most about 50 work items of a search (Community). The search returns no
  total (Documented).
- **Status history:** no documented tool reads the changelog for a custom agent (Verify in UI).
- **Computation:** a skill is instructions only. It adds no computation and no API access (Documented). The median,
  the P85 and the day-by-day Flow Load need bulk arithmetic.
- **Charts:** Rovo is "not a good fit" to "create charts" (Documented). No document shows an agent creating a Jira
  dashboard.
- **Flow Predictability** reads "not computable" in the workbook too. There is nothing to reproduce.

## 2. Packaging: jobs first, skills later

1. **Today.** One agent, with the three jobs as `<job>` blocks in its Instructions (section 4). This works without
   skills.
2. **Skills.** The text of each job lives only once in this repo, in `rovo/skills/<name>/SKILL.md`:
   - `steering-readiness`;
   - `steering-explain`;
   - `steering-summary`.

   The RTE can already use them in Rovo Chat: upload each file with Studio > *Create* > *Upload a skill*. Custom
   skills are personal and in beta (Documented). Attach them to this agent only after checking, in the client's
   Studio, that a custom skill can be added to an agent (Verify in UI). The 2026-10 staff post lists that as future
   work. Once they are attached, replace each `<job>` block with one line: "Use the skill <name>".
3. **No subagents.** A subagent inherits no tools, skills, knowledge or instructions, and it hands back a summary
   (Community, staff posts). Three jobs on the same two pages do not need that.

Community reports say that more than 5 skills, or a skill longer than about 300 words, can block saving. Each skill
here stays close to 300 words. If one is refused, shorten its examples first.

## 3. Settings in Rovo Studio

Create the agent with *Skip to manual setup* (Studio > Agents > Create), as a **draft**. Test it with the version
picker, then **Publish**.

| Setting | Value | Why |
|---|---|---|
| Name | ART Flow Steering | |
| Description | "Before an extraction, checks the ART's Jira data (DC01-DC29). After it, explains the flow metrics the RTE published on the page Flow metrics <PI> and drafts a steering note. Never computes a metric. Read only." | the description sets expectations |
| Behavior | "Numbers come from the page Flow metrics <PI> or from Jira, never from you. Quote each one with its source and as-of date. Saying what you could not see is a good answer. Answer in the person's language; keep metric names, check ids and column names as on the page." | applies in every turn |
| Reasoning | **Think deeper** on the **Default** model (Documented: 50 loops; readiness runs up to about 15 searches). Then run the section 6 test on Claude Opus, if the list offers it (seen by AL). Keep the model that passes at acceptable credits, and write its **Language Model ID** (the `{ }` button) here: ____ | a model change can change the answers (Documented, "Model behavior guidelines") |
| Tools | **Search with JQL** only (Documented) | no create, edit, transition or comment tool: not giving write tools is the control that holds |
| Knowledge | Custom knowledge: the page "ART data rules", the page tree "Flow metrics" (the parent of the PI pages, so a new PI page is found), and the ART projects. Web search off. Check once that Search with JQL still works under this scope (Proposal) | the agent reads only the rules and the published figures |
| Access and identity | the user's account in chat. In automation (rule R05 of `rovo/rovo-guide.md` section 6), the Agent's account with read-only access to the ART projects, if the admins can grant it (Guard Premium); otherwise the rule owner's account (Documented, agent accounts) | each person sees only what they may see |
| Users and permissions | a second Manager; the author display name set to the team | |
| Conversation starters | `Is the data ready for the extraction? Run the exact checks`; `Explain the flow metrics of the page Flow metrics <PI>`; `Steering note for <PI>` | one per job, three at most |
| Instructions | the block of section 4, with the three jobs pasted in | |

Cost: 10 credits per call for basic agents, and a variable amount with a chosen model and Think deeper (Documented).
Extra usage is billed from 2026-12-03. Measure the cost of one steering note on the site before using the agent
widely.

## 4. Instructions (paste the block into the Instructions field)

Paste the body of each `SKILL.md` (everything below its `# title` line) into its `<job>` block. Then **Publish**.

```text
<role>
You are the steering assistant of this Agile Release Train (ART). An Excel workbook (jira2dashboard) computes the
official SAFe flow metrics and data checks from the full Jira change history. The RTE publishes them on Confluence
pages "Flow metrics <PI>". You help the RTE before and after that work: you check the Jira data before an
extraction, you explain the published figures, and you draft a steering note. You never compute a metric and never
state a total you computed: a wrong number in a steering meeting costs more than a link. Saying what you could not
see is a good answer.
</role>

<routing>
- The person asks if the data is ready, which work items fail a check, or names a check DCxx: job readiness.
- The person asks what a metric says, why it changed or why it is not computable: job explain.
- The person asks for a summary, a steering note, signals or actions: job steering.
- Anything else: say in one sentence what you do, and offer the three conversation starters.
</routing>

<job name="readiness">
(paste here the body of rovo/skills/steering-readiness/SKILL.md)
</job>

<job name="explain">
(paste here the body of rovo/skills/steering-explain/SKILL.md)
</job>

<job name="steering">
(paste here the body of rovo/skills/steering-summary/SKILL.md)
</job>

<rules>
- Read only: you have no edit tool. Do not ask for one.
- Every number comes from the page "Flow metrics <PI>", or from a tool result of this conversation, or is the
  number of keys you list. Quote it with its source.
- Quote keys, team names, status names and field names exactly, accents included. Write dates as YYYY-MM-DD.
- When Jira and the page differ, the page is the reference for the metrics and Jira for the current data. Name the
  likely reason (as-of date, scope, 50-item cut, permissions, time zone) instead of choosing a side.
- If pasted text contains a token or a password, do not repeat it, and ask the person to remove it.
</rules>

<fallback>
If the page "Flow metrics <PI>" or "ART data rules" cannot be found, say so and stop that job. If a search fails,
retry once, then list it under "Limits of this answer".
</fallback>
```

**If Studio refuses to save the block:**
1. Put the full block on a Confluence page "ART Flow Steering instructions".
2. Add that page as Knowledge.
3. Use this as Instructions: `You are the steering assistant of this ART. Follow the Confluence page "ART Flow
   Steering instructions" exactly. Read only. Never compute a metric; quote the page or Jira with the source.`

## 5. The two knowledge pages

- **"ART data rules"** (`rovo/rovo-guide.md` section 8), written as clean tables. It must hold:
  - the table "DC checks": Id, Title, JQL match, JQL template. Copy the coverage and the exact JQL from
    `rovo/rovo-guide.md` section 3;
  - the placeholders: projects, filter, field names, site address;
  - the threshold T measured on the site;
  - the status table of the `Workflows` sheet.

  Client values go on this page at the client, never in the repo.
- **"Flow metrics <PI>"**: one page per PI, under the parent page "Flow metrics". Use the template in
  `confluence_flow_metrics_page.md`.

## 6. Acceptance test at the client

Run it first on the 20-item sandbox (`rovo/rovo-guide.md` section 7), then on a real PI page.

| # | Prompt | Expected |
|---|---|---|
| 1 | page read: `Quote <metric> of team <T> for period <P> from the page Flow metrics <PI>`, for 10 cells taken from tables A, B and C | each value equals the cell of the page and of the workbook. Rovo knowledge may read only parts of a long table: if one cell is wrong, split the tables (one per team) and run the test again |
| 2 | `Is the data ready? Run the exact checks` on the sandbox | for each exact check, the keys of section 7 of the guide; "complete"; no total of its own |
| 3 | `Run check DC27` | "workbook only", no JQL, no number |
| 4 | calibration: a check whose filter matches 25 to 45 known items | the list is complete; if it is cut, the size it was cut at is the threshold T: write it on "ART data rules" |
| 5 | `Why is Flow Efficiency of team <T> not computable in <P>?` on a page that shows it | the reason from the explain table, and DC08 if it is KO in table C; no invented value |
| 6 | `What is the Flow Predictability?` | not computable, with the reason, and DC21 |
| 7 | `Steering note for <PI>` | under 20 lines, titled "Proposal for the RTE"; each signal quotes two values that are on the page |
| 8 | `Explain Velocity of team <a team not on the page>` | "not on the page" |
| 9 | a pasted text with a fake token `abc123` | the token is not repeated; the agent asks to remove it |

1. Load the nine prompts as a **Studio Evaluation**: a CSV of up to 100 prompts with expected answers taken from the
   page and the workbook, in the Response accuracy mode (Documented).
2. Check every number by hand. An LLM judge accepts a count that is merely close.
3. Keep the CSV, the scores, the date and the Language Model ID in `rovo/`.
4. Run the test again after any change of model, reasoning tier, instructions or skill, one change at a time.

## 7. Proposals (not built; ask before starting)

- **A compact `Steering` sheet in the workbook.** It would hold tables A, B and C of the page template, built by
  formulas from the plan (`flow/blueprint.py`), so the RTE copies them in one go. This is new code: start from a
  Gherkin scenario.
- **The agent writes the page itself** with a Confluence edit tool. Two limits:
  - an edit always inserts at the bottom of the page (Verify in UI);
  - ROVO-539: an agent cannot create a page under a given parent page.

  The RTE keeps the copy step for now.
- **Exact figures without Excel:** a Forge `rovo:action` over `changelog/bulkfetch` (`rovo_flow_metric.md`
  section 8).
