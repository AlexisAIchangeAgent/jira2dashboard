Option Explicit
' Jira Cloud REST API v3: search all issues with all fields, bulk changelogs, write tables.

Public Const JIRA_PAGE_SIZE As Long = 100
Public Const CHANGELOG_CHUNK As Long = 1000          ' API limit: 1000 issues per bulkfetch call

Public Function JiraConfig(ByVal baseUrl As String, ByVal email As String, ByVal apiToken As String, _
                           ByVal proxyServer As String, ByVal proxyBypass As String, _
                           ByVal proxyWindowsAuth As Boolean) As Object
    Cov "JIR01"
    Dim c As Object
    Set c = NewHttpConfig()
    Do While Right$(baseUrl, 1) = "/"
        baseUrl = Left$(baseUrl, Len(baseUrl) - 1)
    Loop
    c("BaseUrl") = baseUrl
    c("Authorization") = BasicAuthHeader(email, apiToken)
    c("ProxyServer") = proxyServer
    c("ProxyBypass") = proxyBypass
    c("ProxyWindowsAuth") = proxyWindowsAuth
    Set JiraConfig = c
End Function

Public Function ProjectsJql(ByVal projectKeys As Collection) As String
    ' ("ABC","DEF") -> project in ("ABC","DEF") ORDER BY key ASC   (a bounded query, as the API requires)
    Cov "JIR02"
    Dim parts() As String, i As Long
    If projectKeys.Count = 0 Then Cov "JIR03": Err.Raise vbObjectError + 600, "ProjectsJql", "No project keys"
    ReDim parts(1 To projectKeys.Count)
    For i = 1 To projectKeys.Count
        parts(i) = """" & Replace(Trim$(projectKeys(i)), """", "") & """"
    Next i
    ProjectsJql = "project in (" & Join(parts, ",") & ") ORDER BY key ASC"
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

Public Function JiraSearchAll(ByVal cfg As Object, ByVal jql As String, ByRef names As Object, _
                              Optional ByVal maxPages As Long = 2000) As Collection
    ' Reads every page. names receives fieldId -> display name (from expand=names).
    Cov "JIR06"
    Dim issues As New Collection, token As String, page As Long
    Dim resp As Object, data As Variant, issue As Variant
    If names Is Nothing Then Set names = CreateObject("Scripting.Dictionary")
    Do
        page = page + 1
        Set resp = HttpRequest(cfg, "GET", JiraSearchUrl(cfg("BaseUrl"), jql, token))
        CheckStatus resp, "search/jql"
        Set data = JsonParse(resp("Text"))
        For Each issue In AsList(JGet(data, "issues"))
            issues.Add issue
        Next issue
        MergeNames names, JGet(data, "names")
        token = CStr(JGet(data, "nextPageToken") & "")
        Application.StatusBar = "Jira search: " & issues.Count & " issues read (page " & page & ")"
        If IsTrue(JGet(data, "isLast")) Or Len(token) = 0 Then Exit Do
        If page >= maxPages Then Cov "JIR07": Err.Raise vbObjectError + 601, "JiraSearchAll", "Too many pages"
    Loop
    Application.StatusBar = False
    Set JiraSearchAll = issues
End Function

Private Sub MergeNames(ByVal names As Object, ByVal pageNames As Variant)
    Dim k As Variant
    If Not IsObject(pageNames) Then Exit Sub
    For Each k In pageNames.Keys
        If Not names.Exists(k) Then names(k) = pageNames(k)
    Next k
End Sub

Public Sub CheckStatus(ByVal resp As Object, ByVal what As String)
    If resp("Status") = 200 Then Exit Sub
    Cov "JIR08"
    Err.Raise vbObjectError + 602, "Jira " & what, "HTTP " & resp("Status") & " after " & _
              resp("Attempts") & " attempt(s). " & resp("Error") & " " & Left$(resp("Text"), 300)
End Sub

Public Function ChangelogBody(ByVal keys As Collection, ByVal nextPageToken As String) As String
    Cov "JIR09"
    Dim body As Object
    Set body = CreateObject("Scripting.Dictionary")
    Set body("issueIdsOrKeys") = keys
    body("maxResults") = 1000
    If Len(nextPageToken) > 0 Then body("nextPageToken") = nextPageToken
    ChangelogBody = JsonStringify(body)
End Function

Public Function JiraChangelogsAll(ByVal cfg As Object, ByVal issueKeys As Collection, _
                                  Optional ByVal chunkSize As Long = CHANGELOG_CHUNK) As Collection
    ' One row per changed item: Array(issueId, changeId, created, authorName, authorId,
    '                                 fieldId, field, fieldtype, from, fromString, to, toString)
    Cov "JIR10"
    Dim rows As New Collection, first As Long, i As Long, chunk As Collection
    For first = 1 To issueKeys.Count Step chunkSize
        Set chunk = New Collection
        For i = first To IIf(first + chunkSize - 1 > issueKeys.Count, issueKeys.Count, first + chunkSize - 1)
            chunk.Add issueKeys(i)
        Next i
        FetchChangelogChunk cfg, chunk, rows
        Application.StatusBar = "Jira changelogs: " & (i - 1) & " / " & issueKeys.Count & " issues"
    Next first
    Application.StatusBar = False
    Set JiraChangelogsAll = rows
End Function

Private Sub FetchChangelogChunk(ByVal cfg As Object, ByVal chunk As Collection, ByVal rows As Collection)
    Dim token As String, resp As Object, data As Variant, icl As Variant, h As Variant, it As Variant
    Do
        Set resp = HttpRequest(cfg, "POST", cfg("BaseUrl") & "/rest/api/3/changelog/bulkfetch", _
                               ChangelogBody(chunk, token))
        CheckStatus resp, "changelog/bulkfetch"
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
        token = CStr(JGet(data, "nextPageToken") & "")
    Loop While Len(token) > 0
End Sub

' ---------- Tables ----------

Public Function IssueFieldIds(ByVal issues As Collection) As Collection
    ' Union of field ids seen on the issues, in first-seen order.
    Cov "JIR11"
    Dim seen As Object, ids As New Collection, issue As Variant, k As Variant, flds As Variant
    Set seen = CreateObject("Scripting.Dictionary")
    For Each issue In issues
        Set flds = JGet(issue, "fields")
        If IsObject(flds) Then
            For Each k In flds.Keys
                If Not seen.Exists(k) Then seen.Add k, True: ids.Add k
            Next k
        End If
    Next issue
    Set IssueFieldIds = ids
End Function

Public Function IssueRow(ByVal issue As Variant, ByVal fieldIds As Collection) As Variant
    ' Array: key, id, then one cell value per field id.
    Dim v() As Variant, i As Long
    ReDim v(0 To fieldIds.Count + 1)
    v(0) = JGet(issue, "key")
    v(1) = JGet(issue, "id")
    For i = 1 To fieldIds.Count
        v(i + 1) = FieldToCell(JPath(issue, "fields." & fieldIds(i)))
    Next i
    IssueRow = v
End Function

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
