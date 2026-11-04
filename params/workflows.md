# Workflows parameters

Rows of the sheet `Workflows`: one row per status and workflow. The application reads this table and writes it
into the workbook; the VBA reads the sheet. Change a category here (or in the workbook), never in code.
Meaning of the columns and decisions: `docs/safe-flow-metrics-and-data-checks.md`, section 1.
A test keeps this table consistent with the statuses table of that document.

`Order` is the position in the normal flow; `-` means none. `Category`: Active, Wait, Done, Removed, Backlog or
Unclassified. `IsStart`, `IsEnd`, `IsFree` are 0 or 1.

| Workflow | Order | Status | Category | IsStart | IsEnd | IsFree |
|---|---|---|---|---|---|---|
| Team | 1 | Nouveau | Backlog | 0 | 0 | 0 |
| Team | 2 | En analyse | Active | 1 | 0 | 0 |
| Team | 3 | Prêt | Wait | 0 | 0 | 0 |
| Team | 4 | Planifié | Wait | 0 | 0 | 0 |
| Team | 5 | En développement | Active | 0 | 0 | 0 |
| Team | 6 | En test | Active | 0 | 0 | 0 |
| Team | 7 | À déployer | Wait | 0 | 0 | 0 |
| Team | 8 | Terminé | Done | 0 | 1 | 0 |
| Team | - | Annulé | Removed | 0 | 0 | 0 |
| Team | - | Dépriorisé | Backlog | 0 | 0 | 1 |
| Feature | 1 | Nouveau | Backlog | 0 | 0 | 0 |
| Feature | 2 | En analyse | Active | 1 | 0 | 0 |
| Feature | 3 | Prêt | Wait | 0 | 0 | 0 |
| Feature | 4 | Planifié | Wait | 0 | 0 | 0 |
| Feature | 5 | En développement | Active | 0 | 0 | 0 |
| Feature | 6 | En test / recette | Active | 0 | 0 | 0 |
| Feature | 7 | À déployer | Wait | 0 | 0 | 0 |
| Feature | 8 | Terminé | Done | 0 | 1 | 0 |
| Feature | - | Annulé | Removed | 0 | 0 | 0 |
| Feature | - | Dépriorisé | Backlog | 0 | 0 | 1 |
