"""Fake Jira Cloud (REST API v3) on localhost, for tests (pytest) and for the VBA client (modHttp).

Built from docs/jira-cloud-api-vba-reference.md. Assumptions that Jira itself must confirm at the client are
listed in that document (section 1). The data comes from issue histories: search and changelog are derived from
the same history, so they never contradict each other.
"""
