# Workflows parameters

Rows of the sheet `Workflows`: one row per status and workflow. The application reads this table and writes it
into the workbook; the VBA reads the sheet. Change a category here (or in the workbook), never in code.
Meaning of the columns and decisions: `docs/safe-flow-metrics-and-data-checks.md`, section 1.
A test keeps this table consistent with the statuses table of that document.

The ROW order of the table is the normal flow (the chart Cycle time by status stacks the statuses in that order).
`Category`: Active, Wait, Done, Removed, Backlog or Unclassified. `IsStart`, `IsEnd`, `IsFree` are 0 or 1. The `Feature` workflow is always needed; the `Team` workflow
only when the sheet Projects has a Team project.

| Workflow | Status | Category | IsStart | IsEnd | IsFree |
|---|---|---|---|---|---|
| Team | Nouveau | Backlog | 0 | 0 | 0 |
| Team | En analyse | Active | 1 | 0 | 0 |
| Team | Prêt | Wait | 0 | 0 | 0 |
| Team | Planifié | Wait | 0 | 0 | 0 |
| Team | En développement | Active | 0 | 0 | 0 |
| Team | En test | Active | 0 | 0 | 0 |
| Team | À déployer | Wait | 0 | 0 | 0 |
| Team | Terminé | Done | 0 | 1 | 0 |
| Team | Annulé | Removed | 0 | 0 | 0 |
| Team | Dépriorisé | Backlog | 0 | 0 | 1 |
| Feature | Nouveau | Backlog | 0 | 0 | 0 |
| Feature | En analyse | Active | 1 | 0 | 0 |
| Feature | Prêt | Wait | 0 | 0 | 0 |
| Feature | Planifié | Wait | 0 | 0 | 0 |
| Feature | En développement | Active | 0 | 0 | 0 |
| Feature | En test / recette | Active | 0 | 0 | 0 |
| Feature | À déployer | Wait | 0 | 0 | 0 |
| Feature | Terminé | Done | 0 | 1 | 0 |
| Feature | Annulé | Removed | 0 | 0 | 0 |
| Feature | Dépriorisé | Backlog | 0 | 0 | 1 |
