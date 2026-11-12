---
name: add-data-check
description: Add a data-quality check (DCnn) to jira2dashboard end to end - Python reference, Excel formula, fake Jira case, counts, docs - with BDD/TDD. Use when AL asks for a new data check or a change to the rule of an existing one.
---

# Add a data check

A check lives in three places that must agree: the Python reference (`flow/checks.py`), the Excel formula
(`flow/check_formulas.py`, a `DCnn` 0/1 column of `ItemFlow`) and a known case of the fake Jira
(`mockjira/anomalies.py`). The `flow_excel.feature` scenario compares Excel with the reference issue by issue.

## 0. Before coding: settle the rule with AL

Ask only what is missing: the id (next free: one after the last `Check("DC..")` in `CHECKS`), the title, the severity,
the rule in one sentence, the Jira fields it needs (a `...Field` of `Config`?), and whether the interpretation is a
**proposal** to validate with the teams (it usually is: label it so in the docs).

## 1. Red tests first (TDD)

- `tests/unit/test_checks.py`: a case that must fail the new check and a near case that must pass, then run
  `.\scripts\run.ps1 test tests/unit/test_checks.py`: red.
- The catalogue test (`test_the_catalogue_covers_dc01_to_dcNN_...`) pins the ids: update its range and name.

## 2. Code, in this order

1. `src/jira2dashboard/flow/checks.py`: a `Check("DCnn", "title", "severity", (needed Config fields,))` row in
   `CHECKS`, and the rule in `run_checks`: status history in `_status_history`, issue fields in `_field_checks`,
   "Blocks" links in `_link_flags`, fill rates in `fill_rates`.
2. `src/jira2dashboard/flow/check_formulas.py`: the `"DCnn": _flag(rule, *needs)` entry of `FLAGS`. Read lists and field ids
   from `Config` (`cfg`, `in_list`), never as literals. Helper columns: `HELPERS` (or `INTERVAL_HELPERS` in
   `flow/formulas.py`). Mind the Excel traps of CLAUDE.md (no `_` in a structured reference column name, dates with a
   half-second margin, blank is `""`).
3. A new Jira field: a raw column in `ISSUES` (`flow/formulas.py`), `ExtractFields` and the row in `vba/modExtract.bas`,
   the fixtures in `vba/modTests.bas`, `FlowIssue` and `mockjira/flow_adapter.py`, and the mapping of the Python path
   (`standalone/mapping.py`, oracle `test_mapping.py`). A new parameter: a row in `params/config.md` (with Example and
   Mandatory) and the Config row count in `tests/features/params.feature`.
4. `src/jira2dashboard/mockjira/anomalies.py`: a known case that breaks only this rule, with the ids it must trigger.

Run `.\scripts\run.ps1 test` until green. Prove the new test discriminates: break the rule on purpose, see it fail,
restore.

## 3. Counts and docs (the count 29 is written in several places)

`grep -rnE "DC01.DC[0-9]+|[0-9]+ (data )?checks|DC[0-9]{2}" --include=*.md --include=*.py . | grep -v archive` and
update every count and range: `README.md` (section 10, the area table), `src/jira2dashboard/manual.py` (sheet
Manual), `claude-code-handoff.md`, `rovo/` (the check coverage of `rovo/rovo-guide.md` section 3 is the reference:
ask AL how the new check is covered by Rovo: exact, partial, to test or workbook only), `params/config.md` if a
parameter description names the checks. Never edit `docs/`.

If a `.bas` changed: update its MD5 in the table of `claude-code-handoff.md` (the post-edit hook reminds you), and
the marker count test if `Cov` markers were added.

## 4. Done when

The data-check scenario of `flow_excel.feature` passes on the real workbook: run the `full-pass` skill (AL starts it
with `/full-pass`) or at least `coverage`, `vba-tests`, `test-excel -k "contr"`, `build`. Then commit and push.
