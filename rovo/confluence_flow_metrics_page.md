# Confluence page template: "Flow metrics <PI>"

The page where the RTE publishes the workbook's figures, one page per PI. The agent ART Flow Steering
(`rovo_train_steering.md`) reads it to explain the figures and draft a steering note. People read it in meetings.

**Who:** the RTE.
**When:** after each **Extract from Jira and compute**, at least at the end of each iteration and of each PI.

Publishing these figures on Confluence is a **client policy decision** (`rovo/rovo-guide.md` section 4):
- the page holds client data;
- it lives on the client's Confluence only, in the ART space;
- nothing from it comes back to this repo.

Labels as in `rovo/rovo-guide.md` (**Verify in UI** = not confirmed in the docs).

---

## 1. Page skeleton

Create a parent page "Flow metrics" in the ART space once. Add it, as a page tree, to the agent's Custom knowledge.
Then for each PI, create a child page titled `Flow metrics <PI>`, with these sections in this order.

**Header** (a two-column table):

| Field | Value |
|---|---|
| As of (UTC) | the as-of date of the `Metrics` sheet |
| Flow Load mode | A or B (the `FlowLoadMode` of `Config`) |
| Scope | the projects and the `Filter` of `Config`, in words or as JQL |
| Workbook | the file name and the date of the extraction |
| Published by | <RTE>, <date> |

**Table A - Flow metrics** (one row per period and team, from the sheet `Metrics`):

| Period | Team | Velocity | FlowTimeMedian | FlowTimeP85 | FlowTimeExcluded | FlowLoadAverage | FlowLoadEnd | FlowEfficiency | FlowPredictability |
|---|---|---|---|---|---|---|---|---|---|

**Table B - Distribution of done items** (one row per period and team, from the sheet `Metrics`):

| Period | Team | DistDoneFeature | DistDoneEnabler | DistDoneDefect | DistDoneAutre | DistDoneUnmapped |
|---|---|---|---|---|---|---|

**Table C - Data checks** (one row per check, from the sheet `DataChecks`):

| Id | Title | Severity | Items | Result |
|---|---|---|---|---|

**Charts:** the 8 charts of the `Dashboard` sheet as images. Set the team picker to "All teams" first. Put each chart
under its title, as in the workbook:
1. Velocity - items done per period
2. Flow Time - median, in days
3. Flow Time - 85th percentile, in days
4. Flow Load - average items in process
5. Flow Efficiency - active time / (active + wait)
6. Flow Distribution - items done by work type (team of the selector)
7. Flow Predictability
8. Flow Load - items in process at the end of each day

Under each image, copy the note the workbook shows under the chart ("N of M points are not computable ...").

**Steering note:** the agent's note (job `steering`), after the RTE has reviewed it. Keep its title "Proposal for
the RTE".

**Change log** at the bottom: date, what changed, who.

## 2. How to fill it from the workbook

1. Run **Extract from Jira and compute**. Wait for the status cell to say it is done, and note the as-of date.
2. **Rows:** the sheet `Metrics` has one row per period of the `Periods` sheet and per team, for all PIs. Paste
   only the rows of this PI's periods, plus the last period of the previous PI, so the first iteration has a
   trend. Filter the column Period first; Excel copies only the visible rows of a filtered table.
3. **Table A:** in the sheet `Metrics`, select the columns of table A, holding Ctrl to add each one. Copy them and
   paste into the page. Excel copies several columns when they cover the same rows. The pasted table keeps the
   values as Excel shows them (Verify in UI).
4. **Table B:** do the same with Period, Team and the five `DistDone` columns.
5. **Table C:** in the sheet `DataChecks`, copy the columns Id, Title, Severity, Items and Result.
6. **Charts:** for each chart of the `Dashboard` sheet, use Home > Copy > Copy as Picture, then paste it into the
   page. In a French Excel the menu is Accueil > Copier > Copier comme image.
7. Run test 1 of `rovo_train_steering.md` section 6 (page read) once on the first page. Run it again if the size
   of the tables changes a lot.

## 3. Rules

- **Paste values,** never formulas or links to the workbook.
- **Keep "not computable" as text.** Never replace it with 0 or leave the cell blank. A blank in the workbook (a
  day after the as-of date) stays blank.
- **Keep tables short.** Only this PI's periods, plus one before (step 2). If table A grows past about 60 rows, split it
  into one table per team. Long tables can be cut in Rovo answers (ROVO-523, Community), and knowledge may read only
  parts of them.
- **Never paste these sheets:**
  - `Config`: it holds the `ApiToken`;
  - `Issues`, `ItemFlow`, `StatusIntervals`, `DailyLoad`: one row per work item or per day. They are too big for a
    page, and the charts already show them.
- **Never change a value by hand.** If a figure looks wrong, fix the data or the parameters, then extract again.

The column names come from the workbook: `METRICS_KEYS` and `METRICS_COLUMNS` in
`src/jira2dashboard/flow/formulas.py`, `DATACHECKS_KEYS` and `DATACHECKS_COLUMNS` in
`src/jira2dashboard/flow/check_formulas.py`, and the chart titles in `src/jira2dashboard/flow/dashboard_spec.py`. If
one of them changes, update this template, the explain skill and the agent's test.
