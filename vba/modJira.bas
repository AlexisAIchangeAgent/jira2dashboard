Option Explicit
' Jira Cloud REST API v3: search all issues with all fields, bulk changelogs, write tables.

Public Const JIRA_PAGE_SIZE As Long = 100
Public Const CHANGELOG_CHUNK As Long = 50             ' issues per bulkfetch call and per step (the API accepts 1000): a step must stay short

Public Const MAX_SEARCH_PAGES As Long = 2000                  ' a search that has not ended after this many pages is given up

Public Function JiraConfig(ByVal baseUrl As String, ByVal email As String, ByVal apiToken As String, _
                           ByVal proxyServer As String, ByVal proxyBypass As String, _
                           ByVal proxyWindowsAuth As Boolean) As Object
    Cov "JIR01"
    Dim c As Object
    Set c = NewHttpConfig()
    baseUrl = TrimWhite(baseUrl)
    Do While Right$(baseUrl, 1) = "/"
        baseUrl = Left$(baseUrl, Len(baseUrl) - 1)
    Loop
    If Len(baseUrl) > 0 And LCase$(Left$(baseUrl, 8)) <> "https://" And Not IsLocalUrl(baseUrl) Then
        Cov "JIR22"
        Err.Raise vbObjectError + 602, "JiraConfig", "BaseUrl must start with https:// (the API token would travel in clear text). Fix it in the Config sheet."
    End If
    c("BaseUrl") = baseUrl
    c("Authorization") = BasicAuthHeader(TrimWhite(email), TrimWhite(apiToken))   ' as the PowerShell script: $Token.Trim()
    c("ProxyServer") = NormalizeProxy(proxyServer)
    c("ProxyBypass") = TrimWhite(proxyBypass)
    c("ProxyWindowsAuth") = proxyWindowsAuth
    Set JiraConfig = c
End Function

Private Function IsLocalUrl(ByVal url As String) As Boolean
    ' The fake Jira of the tests listens on this computer (or behind the fake proxy under a .invalid name, which never
    ' resolves to a real server): http is fine there.
    Dim h As String, p As Long
    h = LCase$(url)
    If Left$(h, 7) <> "http://" Then Exit Function
    h = Mid$(h, 8)
    p = InStr(h, "/"): If p > 0 Then h = Left$(h, p - 1)         ' the host alone: "localhost.evil.example" is not local
    p = InStr(h, "?"): If p > 0 Then h = Left$(h, p - 1)
    p = InStr(h, ":"): If p > 0 Then h = Left$(h, p - 1)
    IsLocalUrl = (h = "127.0.0.1" Or h = "localhost" Or Right$(h, 8) = ".invalid")
End Function

Public Function NormalizeProxy(ByVal proxy As String) As String
    ' host:port whatever was typed: "http://proxy.name.com:8080/" -> "proxy.name.com:8080" (WinHTTP SetProxy wants no scheme).
    Cov "JIR14"
    Dim p As String
    p = TrimWhite(proxy)
    If LCase$(Left$(p, 7)) = "http://" Then p = Mid$(p, 8)
    If LCase$(Left$(p, 8)) = "https://" Then p = Mid$(p, 9)
    Do While Right$(p, 1) = "/"
        p = Left$(p, Len(p) - 1)
    Loop
    NormalizeProxy = p
End Function

Public Function JiraConfigFromSheet(ByVal wsConfig As Object) As Object
    ' The connection and the proxy come from the Config sheet of the workbook (written from params\config.md);
    ' nothing about the client is in the code.
    ' ProxyServer: typed = that proxy; DIRECT = no proxy; empty = the proxy of the Windows settings, found as the browser
    ' finds it (modProxy.DetectProxy: WPAD, PAC script or fixed proxy), once per run. A typed ProxyBypass stays.
    Cov "JIR15"
    Dim c As Object, secs As Long, proxy As String, bypass As String, found As Object
    proxy = ConfigValue(wsConfig, "ProxyServer"): bypass = ConfigValue(wsConfig, "ProxyBypass")
    Set c = JiraConfig(ConfigValue(wsConfig, "BaseUrl"), ConfigValue(wsConfig, "Email"), ConfigValue(wsConfig, "ApiToken"), _
                       IIf(UCase$(proxy) = "DIRECT", "", proxy), bypass, IsYes(ConfigValue(wsConfig, "ProxyWindowsAuth", "Yes")))
    If UCase$(proxy) = "DIRECT" Then
        Cov "JIR26"
        c("ProxyMode") = "direct (ProxyServer = DIRECT in the Config sheet)"
    ElseIf Len(proxy) = 0 Then
        Cov "JIR27"
        LogBegin "proxy detection"
        Set found = DetectProxy(c("BaseUrl"))
        c("ProxyServer") = NormalizeProxy(found("Server"))
        If Len(bypass) = 0 Then c("ProxyBypass") = found("Bypass")
        c("ProxyMode") = found("Mode")
        LogEnd "proxy detection: " & found("Mode")
    Else
        c("ProxyMode") = "the proxy of the Config sheet"
    End If
    secs = CLng(Val(ConfigValue(wsConfig, "TimeoutSeconds", "60")))
    If secs < 1 Then secs = 60
    c("ConnectMs") = secs * 1000: c("SendMs") = secs * 1000: c("ReceiveMs") = secs * 1000
    c("MaxRetries") = CLng(Val(ConfigValue(wsConfig, "MaxRetries", "4")))
    If c("MaxRetries") < 0 Then c("MaxRetries") = 0                 ' a negative count would skip the request altogether
    Set JiraConfigFromSheet = c
End Function

Public Function JiraConfigFromWorkbook() As Object
    Set JiraConfigFromWorkbook = JiraConfigFromSheet(ThisWorkbook.Worksheets(SheetName("Config")))
End Function

Public Function ProjectsJql(ByVal projectKeys As Collection, Optional ByVal filter As String = "") As String
    ' ("ABC","DEF") -> project in ("ABC","DEF") ORDER BY key ASC   (a bounded query, as the API requires)
    ' A filter (Config sheet) is a JQL restriction added with AND: ... AND (updated >= -90d) ORDER BY key ASC.
    Cov "JIR02"
    Dim parts() As String, i As Long, jql As String
    If projectKeys.Count = 0 Then Cov "JIR03": Err.Raise vbObjectError + 600, "ProjectsJql", "No project to extract: fill the Projects sheet (Project, Role, Team)."
    ReDim parts(1 To projectKeys.Count)
    For i = 1 To projectKeys.Count
        parts(i) = """" & Replace(Trim$(projectKeys(i)), """", "") & """"
    Next i
    jql = "project in (" & Join(parts, ",") & ")"
    filter = TrimWhite(filter)
    If Len(filter) > 0 Then
        Cov "JIR19"
        If InStr(1, filter, "order by", vbTextCompare) > 0 Then
            Cov "JIR20": Err.Raise vbObjectError + 603, "ProjectsJql", "The Filter of the Config sheet must not contain ORDER BY"
        End If
        jql = jql & " AND (" & filter & ")"
    End If
    ProjectsJql = jql & " ORDER BY key ASC"
End Function

Public Function ExtractionJql(ByVal wsProjects As Object, ByVal wsConfig As Object) As String
    ' What the extraction searches: the projects of the Projects sheet, restricted by the Filter of the Config sheet.
    Cov "JIR21"
    ExtractionJql = ProjectsJql(ProjectKeys(wsProjects), ConfigValue(wsConfig, "Filter"))
End Function

Public Function JiraSearchUrl(ByVal baseUrl As String, ByVal jql As String, ByVal nextPageToken As String, _
                              Optional ByVal maxResults As Long = JIRA_PAGE_SIZE, _
                              Optional ByVal fields As String = "*all", _
                              Optional ByVal expand As String = "names") As String
    Cov "JIR04"
    Dim u As String
    u = baseUrl & "/rest/api/3/search/jql?jql=" & UrlEncode(jql) & "&maxResults=" & maxResults & _
        "&fields=" & UrlEncode(fields)
    If Len(expand) > 0 Then u = u & "&expand=" & UrlEncode(expand)
    If Len(nextPageToken) > 0 Then Cov "JIR05": u = u & "&nextPageToken=" & UrlEncode(nextPageToken)
    JiraSearchUrl = u
End Function

Public Function JiraSearchPage(ByVal cfg As Object, ByVal jql As String, ByVal token As String, ByVal issues As Collection, _
                               ByVal names As Object, ByVal fields As String, ByVal expand As String, ByRef isLast As Boolean) As String
    ' One page of the search: adds its issues and names, returns the token of the next page. isLast is True on the last page.
    Cov "JIR23"
    Dim resp As Object, data As Variant, issue As Variant
    Set resp = HttpRequest(cfg, "GET", JiraSearchUrl(cfg("BaseUrl"), jql, token, JIRA_PAGE_SIZE, fields, expand))
    CheckStatus resp, "search/jql"
    RequireJson resp, "search/jql"
    Set data = JsonParse(resp("Text"))
    For Each issue In AsList(JGet(data, "issues"))
        issues.Add issue
    Next issue
    MergeNames names, JGet(data, "names")
    JiraSearchPage = CStr(JGet(data, "nextPageToken") & "")
    isLast = IsTrue(JGet(data, "isLast")) Or Len(JiraSearchPage) = 0
    If Not isLast And JiraSearchPage = token Then
        Cov "JIR25"
        Err.Raise vbObjectError + 606, "JiraSearchPage", "Jira keeps sending the same page token: stop here and try again later."
    End If
End Function

Private Sub MergeNames(ByVal names As Object, ByVal pageNames As Variant)
    Dim k As Variant
    If Not IsObject(pageNames) Then Exit Sub
    For Each k In pageNames.Keys
        If Not names.Exists(k) Then names(k) = pageNames(k)
    Next k
End Sub

Public Sub RequireJson(ByVal resp As Object, ByVal what As String)
    ' A proxy or a login page can answer 200 with HTML: say so, with the start of the text, instead of a parser error.
    Dim first As String
    first = Left$(LTrim$(resp("Text")), 1)
    If first = "{" Or first = "[" Then Exit Sub
    Cov "JIR24"
    Err.Raise vbObjectError + 607, "Jira " & what, "Jira answered 200 with something that is not JSON (a proxy or a login page?). " & _
              "Check ProxyServer and BaseUrl in the Config sheet. Start of the answer: " & Left$(resp("Text"), 200)
End Sub

Public Sub CheckStatus(ByVal resp As Object, ByVal what As String)
    If resp("Status") = 200 Then Exit Sub
    Cov "JIR08"
    Err.Raise vbObjectError + 602, "Jira " & what, "HTTP " & resp("Status") & " after " & _
              resp("Attempts") & " attempt(s). " & resp("Error") & " " & Left$(resp("Text"), 300) & StatusHint(resp("Status"))
End Sub

Private Function StatusHint(ByVal status As Long) As String
    ' Points to the Config cells to check, for the three failures that come from the connection itself.
    Select Case status
        Case 407: Cov "JIR16": StatusHint = " Hint: the proxy asks for authentication. In the Config sheet check ProxyServer and ProxyWindowsAuth (Yes = the Windows session is sent to the proxy)."
        Case 401: Cov "JIR17": StatusHint = " Hint: Jira refused the account. In the Config sheet check Email and ApiToken."
        Case 0: Cov "JIR18": StatusHint = " Hint: no answer. In the Config sheet check BaseUrl, ProxyServer and ProxyBypass (ProxyServer empty = the proxy of the Windows settings, DIRECT = no proxy)."
    End Select
End Function

Private Function ChangelogBody(ByVal keys As Collection, ByVal nextPageToken As String) As String
    Cov "JIR09"
    Dim body As Object
    Set body = CreateObject("Scripting.Dictionary")
    Set body("issueIdsOrKeys") = keys
    body("maxResults") = 1000
    If Len(nextPageToken) > 0 Then body("nextPageToken") = nextPageToken
    ChangelogBody = JsonStringify(body)
End Function

Public Function JiraChangelogChunk(ByVal cfg As Object, ByVal issueKeys As Collection, ByVal first As Long, _
                                   ByVal rows As Collection, Optional ByVal chunkSize As Long = CHANGELOG_CHUNK) As Long
    ' The changelogs of `chunkSize` issues from the position `first`, added to rows. Returns the position of the next chunk.
    Dim chunk As New Collection, i As Long
    For i = first To IIf(first + chunkSize - 1 > issueKeys.Count, issueKeys.Count, first + chunkSize - 1)
        chunk.Add issueKeys(i)
    Next i
    FetchChangelogChunk cfg, chunk, rows
    JiraChangelogChunk = i
End Function

Private Sub FetchChangelogChunk(ByVal cfg As Object, ByVal chunk As Collection, ByVal rows As Collection)
    Dim token As String, resp As Object, data As Variant, icl As Variant, h As Variant, it As Variant, pages As Long, previous As String
    Do
        Set resp = HttpRequest(cfg, "POST", cfg("BaseUrl") & "/rest/api/3/changelog/bulkfetch", _
                               ChangelogBody(chunk, token))
        CheckStatus resp, "changelog/bulkfetch"
        RequireJson resp, "changelog/bulkfetch"
        Set data = JsonParse(resp("Text"))
        For Each icl In AsList(JGet(data, "issueChangeLogs"))
            For Each h In AsList(JGet(icl, "changeHistories"))
                For Each it In AsList(JGet(h, "items"))
                    rows.Add Array(JGet(icl, "issueId"), JGet(h, "id"), JGet(h, "created"), _
                                   JPath(h, "author.displayName"), JPath(h, "author.accountId"), _
                                   JGet(it, "fieldId"), JGet(it, "field"), JGet(it, "fieldtype"), _
                                   JGet(it, "from"), JGet(it, "fromString"), JGet(it, "to"), JGet(it, "toString"))
                Next it
            Next h
        Next icl
        previous = token
        token = CStr(JGet(data, "nextPageToken") & "")
        pages = pages + 1
        If Len(token) > 0 And (token = previous Or pages > 5000) Then
            Err.Raise vbObjectError + 505, "FetchChangelogChunk", "Jira keeps sending the same changelog page: stop here and try again later."
        End If
    Loop While Len(token) > 0
End Sub

' ---------- Tables ----------

Public Function SafeCell(ByVal value As Variant) As Variant
    ' Every non-empty text gets a leading apostrophe, so Excel keeps it as text (tested).
    ' Without it Excel converts: "true" -> TRUE, "1.10" -> 1.1, "2026-10-01" and "1/2" -> dates,
    ' "007" -> 7, "50%" -> 0.5, "#N/A" -> an error value, "=1+1" -> a formula.
    ' The apostrophe is not part of the value: .Value reads the text back unchanged.
    If VarType(value) = vbString Then
        If Len(value) > 0 Then Cov "JIR12": SafeCell = "'" & Left$(value, 32766): Exit Function
    End If
    SafeCell = value
End Function

Public Sub WriteTable(ByVal ws As Object, ByVal headers As Variant, ByVal rows As Collection, _
                      Optional ByVal topRow As Long = 1)
    ' Clears the sheet, writes headers then all rows in one block (fast).
    Cov "JIR13"
    Dim data() As Variant, r As Long, c As Long, nCols As Long, row As Variant
    nCols = UBound(headers) - LBound(headers) + 1
    ws.Cells.ClearContents
    ReDim data(1 To rows.Count + 1, 1 To nCols)
    For c = 1 To nCols
        data(1, c) = SafeCell(headers(LBound(headers) + c - 1))
    Next c
    r = 1
    For Each row In rows
        r = r + 1
        For c = 1 To nCols
            If c - 1 + LBound(row) <= UBound(row) Then data(r, c) = SafeCell(FieldToCell(row(LBound(row) + c - 1)))
        Next c
    Next row
    ws.Range(ws.Cells(topRow, 1), ws.Cells(topRow + rows.Count, nCols)).Value = data
End Sub
