"""Compute the flow metrics again from a filled workbook, without Jira (scripts\run.ps1 compute): see __main__.COMPUTE_DOC."""
import sys

from jira2dashboard.standalone.__main__ import main

if __name__ == "__main__":
    sys.exit(main(sys.argv[1:], command="compute"))
