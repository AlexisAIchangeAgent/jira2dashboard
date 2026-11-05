# Jira dashboard: ART data quality (built by hand, no Rovo)

A Jira dashboard of saved filters, one per data check that JQL can find. **Jira computes the counts, not Rovo**, so
they are exact for the checks marked exact.

**Who:** the Jira admin builds it once; the RTE and the teams use it before each extraction.
**Where it fits:** step 1 of the steering cycle (`rovo_train_steering.md` section 1), next to the job `readiness` of
the agent.

Labels as in `rovo/rovo-guide.md` (**Verify in UI** = not confirmed in the docs; **Proposal** = to validate). The
coverage comes from `rovo/rovo-guide.md` section 3, which was checked against `src/jira2dashboard/flow/checks.py`.

---

## 1. Saved filters

**Name:** `DQ DCxx <Title>`. The Title is the one in the workbook's `DataChecks` sheet, for example `DQ DC13 No Team
value`. The same ids appear everywhere: `DataChecks`, the agent's answers, the filters and the Confluence pages.

**Scope:** prefix each JQL with `project in (<Projects>) AND (<Filter>) AND `, so the scope is the workbook's.

**Placeholders:** the placeholders are workbook parameters. Fill them at the client only, never in this repo.

**Exact** (they should list exactly the items of `DataChecks` on the same day):

| DC | JQL |
|---|---|
| DC04 | `statusCategory = Done AND resolutiondate IS EMPTY` |
| DC05 | `statusCategory != Done AND resolutiondate IS NOT EMPTY` |
| DC11 | `issuetype in (<PointsRequiredTypes>) AND sprint in (openSprints(), futureSprints()) AND "<StoryPointsField>" IS EMPTY` |
| DC12 | `issuetype in (<SubTaskTypes>) AND "<StoryPointsField>" IS NOT EMPTY` |
| DC13 | `issuetype in (<FeatureTypes>, <ParentRequiredTypes>) AND "<TeamField>" IS EMPTY` |
| DC14 | one filter per team project P: `project = <P> AND "<TeamField>" IS NOT EMPTY AND "<TeamField>" != "<TeamOfP>"` (test how the Team field compares) |
| DC15 | `project in (<TeamProjects>) AND statusCategory = "In Progress" AND sprint IS EMPTY` |
| DC19 | `issuetype in (<FeatureTypes>) AND status not in (<BacklogStatuses>, <RemovedStatuses>) AND ("<TargetStartField>" IS EMPTY OR "<TargetEndField>" IS EMPTY)` |
| DC21 | `issuetype in (<FeatureTypes>) AND "<BusinessValueField>" IS NOT EMPTY` |
| DC26 | `issuetype in (<FeatureTypes>) AND fixVersion IS EMPTY` |

**To test first:** `parent IS EMPTY` is not documented, though community answers say it works. Test it on the
sandbox (`rovo/rovo-guide.md` section 7, items 9 and 10).

| DC | JQL |
|---|---|
| DC01 | `project in (<TeamProjects>) AND issuetype in (<ParentRequiredTypes>) AND parent IS EMPTY` |

**Partial:** add "partial" to the filter name. Each one lists candidates, not the workbook's exact list.

| DC | JQL | What it misses |
|---|---|---|
| DC08 | `status not in (<AllWorkflowStatuses>)` | it sees the current status only; the workbook also flags past statuses |
| DC09 | `status in (<EndStatuses>) AND status WAS NOT IN (<StartStatuses>)` | it ignores the restart of the clock after a return to Backlog |
| DC10 | `status CHANGED FROM "<EndStatus>"`, one clause per end status, joined with OR | exact only with one Done status per workflow |
| DC17 | `issuetype in (<FeatureTypes>) AND "<PIField>" IS EMPTY` | the workbook also counts a malformed PI as empty |

**Workbook only** (no filter):
- DC02, DC03, DC06, DC07, DC18, DC20, DC24, DC27, DC29;
- DC16, DC22, DC23, DC25, DC28: their JQL gives only loose candidates (`rovo/rovo-guide.md` section 3).

Read these in the `DataChecks` sheet.

**Sharing:** share each filter with the ART group, in view mode. The owner is the Jira admin.

## 2. Dashboard

**Name:** "ART data quality". Share it with the ART group.

**Gadgets** (names from `rovo/rovo-guide.md` section 6; exact labels: Verify in UI):
- **Filter results**, one per check with a Blocker severity: DC01, DC08, DC11, DC13, DC17, DC19. Show the
  columns Key, Summary, Team, Status and Assignee. The gadget shows how many items match.
- **Two dimensional filter statistics**, one per other exact filter. Put the Team field on one axis and Status on
  the other, so each team sees its own share.
- Put the Blocker checks at the top, partial ones included (DC08 and DC17 keep "partial" in their name). Put the
  other partial filters (DC09, DC10) in a separate column, titled "partial: candidates".

**Counts depend on the viewer.** A viewer sees only the items they may see. A count can differ from the workbook's
when the workbook's Jira account sees more.

## 3. Use in the cycle

1. **Before each extraction:** open the dashboard. Fix the Blocker checks first, then extract.
2. **Each month:** compare the exact filters with the `Items` column of `DataChecks`, on the same day. A difference
   is a real signal: scope, field id, renamed status or permissions. Note it in the `RovoCompare` sheet, column
   "Saved filter count" (`rovo/rovo-guide.md` section 9).
3. **Flow metrics:** the workbook draws the flow metric charts, and the page "Flow metrics <PI>" shows them as
   images. No native gadget computes flow time, load or efficiency (research notes, Community).

   Proposal: a Marketplace time-in-status gadget. It uses its own definitions, so compare it with the workbook
   before showing it.
