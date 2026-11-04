# Config parameters

Rows of the sheet `Config`: the connection to Jira, including the proxy. The application writes this table into
the workbook and the VBA reads the sheet (`modJira.JiraConfigFromSheet`). **Fill the values in the workbook**:
nothing from the client may be kept in this repo, so `BaseUrl`, `Email`, `ProxyServer` and `ProxyBypass` stay
empty here, and **`ApiToken` must stay empty** (a test refuses a token in this file).

| Parameter | Value | Description |
|---|---|---|
| BaseUrl |  | Jira site, for example https://your-site.atlassian.net (no trailing slash needed) |
| Email |  | Account e-mail (Basic authentication with the API token) |
| ApiToken |  | Type the API token in the workbook only. Spaces and line breaks around it are removed |
| Filter |  | Extra JQL restriction added with AND to the project list (sheet Projects), for example updated >= -90d. Empty = none |
| WorkTypeField |  | Id of the custom field that holds the work type (TypeCategorie: Feature, Enabler, Defect, Autre), for example customfield_10200. Find it with GET /rest/api/3/field |
| ProxyServer |  | Proxy as host:port or http://host:port. Empty = no explicit proxy (WinHTTP default) |
| ProxyBypass |  | Hosts that skip the proxy, for example *.client.local;<local> |
| ProxyWindowsAuth | Yes | Yes = answer the proxy's Windows authentication with the Windows session. No = never send it |
| FlowLoadMode | A | A = a ticket counts from its first start until it is done or removed, except while in a Backlog status. B = only from its last start (after the last return to backlog) |
| AsOf |  | UTC date and time of the extraction, for example 2026-11-03 18:00. Open statuses run up to it, so type it for reproducible figures. Empty = the moment the metrics are computed |
| TimeoutSeconds | 60 | Connect, send and receive time limit of one request, in seconds |
| MaxRetries | 4 | Retries after a 429 or a server error (5xx), so 5 attempts at most |
