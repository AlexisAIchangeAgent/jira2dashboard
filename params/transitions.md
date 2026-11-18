# Transitions parameters

Rows of the sheet `Transitions`: one row per allowed move, forward and backward, for each workflow. The
application reads this table and writes it into the workbook; the VBA reads the sheet.
A status with `IsFree = 1` in `workflows.md` (Dépriorisé) is reachable from any status: that is carried by the
flag, not by extra rows here. Meaning and decisions: `docs/safe-flow-metrics-and-data-checks.md`, section 1.
A test keeps this table consistent with the transitions table of that document.

| Workflow | From | To |
|---|---|---|
| Team | Nouveau | En analyse |
| Team | En analyse | Prêt |
| Team | Prêt | Planifié |
| Team | Prêt | En analyse |
| Team | Planifié | En développement |
| Team | Planifié | Prêt |
| Team | En développement | En test |
| Team | En test | À déployer |
| Team | En test | En développement |
| Team | À déployer | Terminé |
| Team | À déployer | En test |
| Team | Dépriorisé | Annulé |
| Team | Dépriorisé | Nouveau |
| Feature | Nouveau | En analyse |
| Feature | En analyse | Prêt |
| Feature | Prêt | Planifié |
| Feature | Prêt | En analyse |
| Feature | Planifié | En développement |
| Feature | Planifié | Prêt |
| Feature | En développement | En test / recette |
| Feature | En test / recette | À déployer |
| Feature | En test / recette | En développement |
| Feature | À déployer | Terminé |
| Feature | À déployer | En test / recette |
| Feature | Dépriorisé | Annulé |
| Feature | Dépriorisé | Nouveau |
