"""The Python path without VBA: read Jira, compute the data and the results, write jira2dashboard.xlsx.

`scripts\\run.ps1 extract` (see __main__.py). It reads the parameter sheets of the workbook as a file, calls Jira with
WinHTTP (the same network stack as the VBA), computes with the pure reference of flow/, and writes values and SVG
charts with the standard library only: no Excel, no VBA. So the data and the results are there even when the
workbook's macros fail.
"""
