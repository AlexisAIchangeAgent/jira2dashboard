"""Reference implementation of the six SAFe flow metrics (pure Python, no Excel, no Jira).

Beside it: formulas.py (the Excel formulas as text), blueprint.py and dashboard_spec.py (the plan written into\nthe workbook, which modFlowBuild.bas turns into tables and charts) and sheets.py (the COM part).\n\nWritten from docs/safe-flow-metrics-and-data-checks.md section 2. It is the oracle against which the Excel
formulas are compared cell by cell, so it must follow the document's text, not the formulas.
"""
