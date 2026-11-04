# Transitions parameters

Rows of the sheet `Transitions`: one row per allowed move, forward and backward, for each workflow. The
application reads this table and writes it into the workbook; the VBA reads the sheet.
A status with `IsFree = 1` in `workflows.md` (Dépriorisé) is reachable from any status: that is carried by the
flag, not by extra rows here. Meaning and decisions: `docs/safe-flow-metrics-and-data-checks.md`, section 1.
A test keeps this table consistent with the transitions table of that document.

| Workflow | From | To | Type |
|---|---|---|---|
| Team | Nouveau | En analyse | forward |
| Team | En analyse | Prêt | forward |
| Team | Prêt | Planifié | forward |
| Team | Prêt | En analyse | back |
| Team | Planifié | En développement | forward |
| Team | Planifié | Prêt | back |
| Team | En développement | En test | forward |
| Team | En test | À déployer | forward |
| Team | En test | En développement | back |
| Team | À déployer | Terminé | forward |
| Team | À déployer | En test | back |
| Team | Dépriorisé | Annulé | exit |
| Team | Dépriorisé | Nouveau | back to backlog |
| Feature | Nouveau | En analyse | forward |
| Feature | En analyse | Prêt | forward |
| Feature | Prêt | Planifié | forward |
| Feature | Prêt | En analyse | back |
| Feature | Planifié | En développement | forward |
| Feature | Planifié | Prêt | back |
| Feature | En développement | En test / recette | forward |
| Feature | En test / recette | À déployer | forward |
| Feature | En test / recette | En développement | back |
| Feature | À déployer | Terminé | forward |
| Feature | À déployer | En test / recette | back |
| Feature | Dépriorisé | Annulé | exit |
| Feature | Dépriorisé | Nouveau | back to backlog |
