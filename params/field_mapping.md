# FieldMapping parameters

Rows of the sheet `FieldMapping` (one row per column of the Jira ticket extract). The application reads this
table and writes it into the workbook; the VBA reads the sheet. Edit a row here (or in the workbook), never in code.
Source and meaning of the columns: `docs/safe-flow-metrics-and-data-checks.md`, section "Sheet `FieldMapping`".
A test keeps this table equal to the one in that document.

| Extract | Logical | Type | CanBeEmpty | UsedBy | Role | IfMissing | TypeConfirmed |
|---|---|---|---|---|---|---|---|
| Projet | Project | text | no | team mapping (Config), every metric per team; DC14 (needs `Team`, not in the extract) | Required | no team split; DC14 not computable | Yes |
| Cle | Key | text | no | join key of all tables, traceability | Required | nothing computable | Yes |
| Type | IssueType | text | no | DC01, DC02, Flow Velocity per type | Required | no per-type metric | Yes |
| TypeCategorie | WorkType | text (`Feature`, `Enabler`, `Autre`, `Defect`) | no | Flow Distribution | Required | Flow Distribution not computable | Yes |
| Resume | Summary | text | no | readable lists of keys | Display | none | Yes |
| Status | Status | text | no | `StatusIntervals` (CURRENT_MISMATCH), DC08, DC28 | Required | no intervals check | Yes |
| Categorie | JiraStatusCategory | text (`À faire`, `En cours`, `Terminé`) | no | DC04, DC05; DC15 (needs `Sprint`, not in the extract) | Required | DC04, DC05, DC15 not computable | Yes |
| Livre | Released | True/False (released = available to the users) | no | **usage to decide** (see below) | Optional | none | Yes |
| Resolution | Resolution | text | yes | none today | Optional | none | Yes |
| Assigne | Assignee | text | yes | none today | Display | none | Yes |
| Priorite | Priority | text | no | Plans ranking (display) | Display | none | Yes |
| Cree | CreatedUtc | date-time | no | Lead Time, BEFORE_CREATED | Required | no lead time, no interval start | Yes |
| MisAjour | UpdatedUtc | date-time | no | "as of" control | Optional | none | Yes |
| Resolu | ResolvedUtc | date-time | yes | DC04, DC05 | Required | DC04, DC05 not computable | Yes |
| Echeance | DueDate | date | yes | DC20 if the plan uses due dates | Optional | none | Yes |
| Label | Labels | text | yes | none today | Optional | none | Yes |
| FixVersion | FixVersion | text | yes | DC26 | Optional | DC26 not computable | Yes |
| PIsource | PI | text (`PI#XX-YYYY`, `XX` 2 digits) | yes | DC17, DC18, DC29 | Required for PI views | DC17, DC18 not computable | Yes |
| Parent | ParentKey | text (issue key) | yes | DC01, DC02, DC24, hierarchy | Required | DC01, DC02, DC24 not computable | Yes |
| BusinessValue | BusinessValue | integer | yes | DC21 only (Flow Predictability needs the business value of PI objectives, which are not in the extract) | Optional | DC21 not computable | Yes |
| StoryPoint | StoryPoints | number | yes | DC12, Flow Velocity in points, fill rate; DC11 (needs `Sprint`, not in the extract) | Optional | points view "not reliable"; DC11 not computable | Yes |
| URL | Url | text | no | link from a metric to the issue (traceability rule 3) | Display | none | Yes |
