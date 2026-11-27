# Workflows parameters

Rows of the sheet `Workflows`: **exceptions to what Jira says** (AL, 2026-11-18). The workbook reads the status categories and the workflows from
Jira by itself (new -> Backlog, in progress -> Active, done -> Done and the end of the flow, a GLOBAL transition = free; the flow starts at
the first status in progress, or at the statuses of Config FeatureStartStatuses / TeamStartStatuses; `flow/status_rules.py`); a row here wins, cell by cell (an empty cell keeps Jira's value, an empty `Workflow` touches both).
The rows below are the example of the delivered workbook (French statuses): at a client, keep only what Jira cannot decide (Wait, Removed, a Done status
that is no delivery). Change a category here (or in the workbook), never in code.
Meaning of the columns and decisions: `docs/safe-flow-metrics-and-data-checks.md`, section 1.
A test keeps this table consistent with the statuses table of that document.

The ROW order of the table is the normal flow (the chart Cycle time by status stacks the statuses in that order).
`Category`: Active, Wait, Done, Removed, Backlog or Unclassified (or empty). `IsEnd`, `IsFree` are 0, 1 or empty. A status is compared
as Excel compares it (trimmed, case ignored). Flow Efficiency needs at least one status declared Wait: without one it is not computable.

| Workflow | Status | Category | IsEnd | IsFree |
|---|---|---|---|---|
| Team | Nouveau | Backlog | 0 | 0 |
| Team | En analyse | Active | 0 | 0 |
| Team | Prêt | Wait | 0 | 0 |
| Team | Planifié | Wait | 0 | 0 |
| Team | En développement | Active | 0 | 0 |
| Team | En test | Active | 0 | 0 |
| Team | À déployer | Wait | 0 | 0 |
| Team | Terminé | Done | 1 | 0 |
| Team | Annulé | Removed | 0 | 0 |
| Team | Dépriorisé | Backlog | 0 | 1 |
| Feature | Nouveau | Backlog | 0 | 0 |
| Feature | En analyse | Active | 0 | 0 |
| Feature | Prêt | Wait | 0 | 0 |
| Feature | Planifié | Wait | 0 | 0 |
| Feature | En développement | Active | 0 | 0 |
| Feature | En test / recette | Active | 0 | 0 |
| Feature | À déployer | Wait | 0 | 0 |
| Feature | Terminé | Done | 1 | 0 |
| Feature | Annulé | Removed | 0 | 0 |
| Feature | Dépriorisé | Backlog | 0 | 1 |
