# Projects parameters

Rows of the sheet `Projects`: the Jira projects to extract and the role of each. The application writes this
table into the workbook and the VBA reads the sheet (`modParams.ProjectKeys`, `ProgramProject`, `TeamOfProject`);
change it in the workbook, never in code.

Exactly one project is the **Program** (the train, it holds the Features). The others are **Team** projects
(Stories, Tasks, Sub-tasks); `Team` is the team name shown in the metrics (used by the team split and by DC14).

The keys below are the anonymised keys of the existing extraction script (the first one is the Program).
Replace them with the real keys in the workbook.

| Project | Role | Team |
|---|---|---|
| TTT | Program |  |
| AAA | Team | AAA |
| EEE | Team | EEE |
| FFFF | Team | FFFF |
| TTTT | Team | TTTT |
| SSSS | Team | SSSS |
| SSS | Team | SSS |
| TETET | Team | TETET |
| CCCC | Team | CCCC |
