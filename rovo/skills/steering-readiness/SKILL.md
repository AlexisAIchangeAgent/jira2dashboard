---
name: steering-readiness
description: Check the ART's Jira data before the workbook extraction with the data checks DC01-DC29. Use when asked whether the data is ready, which work items fail a check, or about a check id DCxx. Gives JQL links and the keys received, never a computed total. Read only.
---

# Data readiness before the extraction

1. Read the table "DC checks" on the Confluence page "ART data rules". For each check it gives the Id, the Title, the
   JQL match (exact, partial, to test, none) and the JQL template. Read three more things from the same page: the
   placeholders (projects, filter, field names), the site address, and the threshold T measured on this site.
2. Take the checks the person asks for. If none is named, take every check whose JQL match is exact.
3. JQL match "none": explain the rule in one sentence and write "workbook only: see the sheet DataChecks". Then go
   to the next check.
4. Otherwise fill the placeholders of the template, prefix it with the scope from the page, and run it once with
   Search with JQL.
5. Count the keys you received. If there are fewer than T, the list is "complete". If there are T or more, it is
   "truncated". Never write a total of your own.
6. Build the link https://<site>/issues/?jql=<the JQL, URL-encoded>.
7. Partial checks: say in one line what the JQL misses, as written on the page.

Output, one row per check:
| Id | Title | JQL match | Items received | Complete | Keys (at most 50) | Link |
Last line: "Exact counts: the saved filters DQ DCxx on the Jira dashboard, or the sheet DataChecks of the workbook."

If a JQL fails, show it with the error, name the placeholder that is probably wrong, and go on to the next check.
