Option Explicit
' Test runner. Each test is a Public Sub named T_xxx. Results go to sheet "TestResults":
' TestId | Area | Description | Input | Expected | Actual | Result (OK/KO) | RunAt
' Offline tests use fixtures and the HTTP fake: they never call Jira.

Private mRows As Collection
Private mCurrent As String
Private mTotalPlanned As Long      ' tests planned for this run (progress bar total)
Private mDone As Long              ' tests finished in this run
Private mNextRow As Long           ' next free row in TestResults

Private Function J(ByVal s As String) As String
    ' Fixtures are written with ' instead of " to stay readable. \' becomes \" (an escaped quote).
    J = Replace(s, "'", """")
End Function

Private Function Repr(ByVal v As Variant) As String
    If IsObject(v) Then Repr = "<" & TypeName(v) & ">": Exit Function
    If IsNull(v) Then Repr = "Null": Exit Function
    If IsEmpty(v) Then Repr = "Empty": Exit Function
    If VarType(v) = vbDate Then Repr = Format$(v, "yyyy-mm-dd hh:nn:ss"): Exit Function
    If VarType(v) = vbBoolean Then Repr = IIf(v, "True", "False"): Exit Function   ' Str$(True) = "Vrai" in French
    If VarType(v) = vbString Then Repr = Replace(Replace(v, vbCr, "\r"), vbLf, "\n"): Exit Function
    Repr = Trim$(Str$(v))
End Function

Private Function SameValue(ByVal a As Variant, ByVal b As Variant) As Boolean
    If IsObject(a) Or IsObject(b) Then Exit Function
    If IsNull(a) Or IsNull(b) Then SameValue = IsNull(a) And IsNull(b): Exit Function
    If IsEmpty(a) Or IsEmpty(b) Then SameValue = IsEmpty(a) And IsEmpty(b): Exit Function
    If VarType(a) = vbString Or VarType(b) = vbString Then SameValue = (CStr(a) = CStr(b)): Exit Function
    SameValue = Abs(CDbl(a) - CDbl(b)) < 0.000001
End Function

Private Sub Check(ByVal area As String, ByVal description As String, ByVal inputText As String, _
                  ByVal expected As Variant, ByVal actual As Variant)
    Dim n As Long
    n = mRows.Count + 1
    AddRow Array(mCurrent & "." & n, area, description, inputText, Repr(expected), Repr(actual), _
                 IIf(SameValue(expected, actual), "OK", "KO"), Now)
End Sub

Private Function TestNames() As Variant
    TestNames = Array("T_JsonNested", "T_JsonEscapes", "T_JsonEmpty", "T_JsonErrors", "T_JsonRoundTrip", _
        "T_JsonLarge", "T_Base64Auth", "T_UrlEncode", "T_Utf8RoundTrip", "T_JiraDates", "T_FieldToCell", _
        "T_SafeCell", "T_ProjectsJql", "T_SearchUrl", "T_SearchPagination", "T_Retry429", _
        "T_RetryExhausted", "T_ChangelogChunks", "T_IssueRows", "T_WriteTable", "T_StatusIntervals", _
        "T_StatusChanges", "T_RetryDelay", "T_SafeAccess", "T_FilesUtf8", "T_StringifyBadObject", _
        "T_TooManyPages", "T_BackgroundExport", "T_TextNotConverted")
End Function

Public Function RunAllTests() As String
    RunAllTests = RunTests(TestNames())
End Function

Public Function RunNetworkTests() As String
    RunNetworkTests = RunTests(NetworkTestNames())
End Function

Private Function NetworkTestNames() As Variant
    NetworkTestNames = Array("N_HttpGetUtf8", "N_HttpPostJson", "N_Http429", "N_HttpBadHost")
End Function

Private Function CountOfNames(ByVal names As Variant) As Long
    CountOfNames = UBound(names) - LBound(names) + 1
End Function

Private Function RunTests(ByVal names As Variant) As String
    ' Each result row is written as soon as its check runs; the progress bar moves after each test.
    Dim nm As Variant, ok As Long, ko As Long, r As Variant, standalone As Boolean
    Set mRows = New Collection
    Randomize 1
    standalone = (mTotalPlanned = 0)
    If standalone Then mTotalPlanned = CountOfNames(names): mDone = 0
    StartRows
    For Each nm In names
        Application.StatusBar = "Test " & (mDone + 1) & "/" & mTotalPlanned & ": " & nm & "..."
        mCurrent = nm
        On Error Resume Next
        Err.Clear
        Application.Run "modTests." & nm
        If Err.Number <> 0 Then AddRow Array(nm & ".ERR", "runner", "Test raised an error", "", "no error", _
                                             Err.Number & " " & Err.Description, "KO", Now)
        Err.Clear
        On Error GoTo 0
        FakeOff
        mDone = mDone + 1
        ProgressShow mDone, mTotalPlanned
    Next nm
    For Each r In mRows
        If r(6) = "OK" Then ok = ok + 1 Else ko = ko + 1
    Next r
    Application.StatusBar = False
    If standalone Then mTotalPlanned = 0
    RunTests = ok & " OK, " & ko & " KO"
End Function

Private Sub StartRows()
    Dim ws As Object
    Set ws = ThisWorkbook.Worksheets("TestResults")
    If ws.Range("A1").Value = "" Then
        ws.Range("A1:H1").Value = Array("TestId", "Area", "Description", "Input", "Expected", "Actual", "Result", "RunAt")
    End If
    mNextRow = ws.Cells(ws.Rows.Count, 1).End(-4162).Row + 1     ' -4162 = xlUp
End Sub

Private Sub AddRow(ByVal row As Variant)
    ' Writes one result row now, so the user sees each check appear.
    Dim ws As Object, data(1 To 1, 1 To 8) As Variant, c As Long
    mRows.Add row
    For c = 1 To 8
        data(1, c) = SafeCell(row(c - 1))
    Next c
    Set ws = ThisWorkbook.Worksheets("TestResults")
    ws.Range(ws.Cells(mNextRow, 1), ws.Cells(mNextRow, 8)).Value = data
    mNextRow = mNextRow + 1
    DoEvents
End Sub

Public Sub ProgressReset()
    ProgressDraw 0, 0
End Sub

Private Sub ProgressShow(ByVal done As Long, ByVal total As Long)
    ProgressDraw done, total
    Application.StatusBar = "Tests " & done & "/" & total & " (" & Format$(done / total, "0%") & ")"
    DoEvents
End Sub

Private Sub ProgressDraw(ByVal done As Long, ByVal total As Long)
    ' Shapes "pbBack" and "pbFill" and the label cell J10 are created by the generator.
    ' The label starts with an apostrophe: "1 / 2" would otherwise be turned into a date.
    Dim ws As Object, pct As Double
    If total > 0 Then pct = done / total
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("TestResults")
    ws.Shapes("pbFill").Width = ws.Shapes("pbBack").Width * pct + 0.1
    If total > 0 Then
        ws.Range("J10").Value = "'" & done & " / " & total & " tests (" & Format$(pct, "0%") & ")"
    Else
        ws.Range("J10").Value = ""
    End If
    Err.Clear
    On Error GoTo 0
End Sub

' ================= JSON =================

Public Sub T_JsonNested()
    Dim d As Object
    Set d = JsonParse(J("{'a':1,'b':{'c':[1,2,{'d':'x'}]},'e':true,'f':null,'g':-1.5e2}"))
    Check "json", "number", "a", 1, d("a")
    Check "json", "array length", "b.c", 3, JPath(d, "b").Item("c").Count
    Check "json", "nested string", "b.c[3].d", "x", d("b")("c")(3)("d")
    Check "json", "true", "e", True, d("e")
    Check "json", "null", "f", Null, d("f")
    Check "json", "exponent", "g", -150, d("g")
End Sub

Public Sub T_JsonEscapes()
    Dim d As Object
    Set d = JsonParse(J("{'s':'a\'b\\c\n\u00e9\/'}"))
    Check "json", "escapes", "a\""b\\c\n\u00e9\/", "a""b\c" & vbLf & ChrW(233) & "/", d("s")
End Sub

Public Sub T_JsonEmpty()
    Check "json", "empty object", "{}", 0, JsonParse("{}").Count
    Check "json", "empty array", "[]", 0, JsonParse("[]").Count
    Check "json", "whitespace", " { ""a"" : [ ] } ", 0, JsonParse(" { ""a"" : [ ] } ")("a").Count
End Sub

Public Sub T_JsonErrors()
    Dim bad As Variant, raised As Boolean, v As Variant
    For Each bad In Array("", "[1,", "{'a':1 'b':2}", "{", "{'a':1}x", "{'a' 1}", "[1 2]", "{'a':tru}", "'abc", "{1:2}", "{'a':'\q'}", "@")
        raised = False
        On Error Resume Next
        v = Empty
        v = JsonParse(J(bad))
        raised = (Err.Number <> 0)
        Err.Clear
        On Error GoTo 0
        Check "json", "invalid JSON raises", CStr(bad), True, raised
    Next bad
End Sub

Public Sub T_JsonRoundTrip()
    Dim src As String
    src = J("{'a':[1,'x',true,false,null,1.5],'b':{'c':'q\'r\\t'},'e':{},'f':[]}")
    Check "json", "parse then stringify", src, src, JsonStringify(JsonParse(src))
    Check "json", "VB array", "Array(1,""a"")", "[1,""a""]", JsonStringify(Array(1, "a"))
    Check "json", "control char", "Chr(1)", """\u0001""", JsonStringify(Chr$(1))
End Sub

Public Sub T_JsonLarge()
    Dim parts() As String, i As Long, t0 As Double, d As Object, n As Long
    n = 3000
    ReDim parts(1 To n)
    For i = 1 To n
        parts(i) = J("{'id':'" & i & "','key':'K-" & i & "','fields':{'summary':'Story n" & i & _
                     " with some text \u00e9','status':{'name':'En analyse','statusCategory':{'key':'indeterminate'}}," & _
                     "'customfield_10016':5,'labels':['a','b'],'description':null}}")
    Next i
    t0 = Timer
    Set d = JsonParse("{""issues"":[" & Join(parts, ",") & "]}")
    Check "json", "3000 issues parsed", "about " & Format$(Len(Join(parts, ",")) / 1024, "0") & " KB", n, d("issues").Count
    Check "json", "parse time under 20 s", "seconds = " & Format$(Timer - t0, "0.00"), True, (Timer - t0) < 20
End Sub

' ================= Text =================

Public Sub T_Base64Auth()
    Check "text", "basic auth ASCII", "user@example.com:abc123", "Basic dXNlckBleGFtcGxlLmNvbTphYmMxMjM=", _
          BasicAuthHeader("user@example.com", "abc123")
    Check "text", "basic auth accents (UTF-8)", "j?r?me@exemple.fr:t?k", "Basic asOpcsO0bWVAZXhlbXBsZS5mcjp0w7Rr", _
          BasicAuthHeader("j" & ChrW(233) & "r" & ChrW(244) & "me@exemple.fr", "t" & ChrW(244) & "k")
End Sub

Public Sub T_UrlEncode()
    Check "text", "url encode JQL", "project in (ABC, ""DEF"") AND summary ~ ""?""", _
          "project%20in%20%28ABC%2C%20%22DEF%22%29%20AND%20summary%20~%20%22%C3%A9%22", _
          UrlEncode("project in (ABC, ""DEF"") AND summary ~ """ & ChrW(233) & """")
    Check "text", "url encode empty", "", "", UrlEncode("")
End Sub

Public Sub T_Utf8RoundTrip()
    Dim s As String, b() As Byte
    s = "Caf" & ChrW(233) & " " & ChrW(8364) & " ok"
    b = Utf8Encode(s)
    Check "text", "UTF-8 byte count", "Cafe + euro", 12, UBound(b) - LBound(b) + 1
    Check "text", "UTF-8 round trip", "Cafe + euro", s, Utf8Decode(b)
    Check "text", "empty string", "", 0, Len(Utf8Decode(Utf8Encode("")))
End Sub

Public Sub T_JiraDates()
    Dim base As Date
    base = DateSerial(2026, 10, 1)
    Check "text", "offset +0200 to UTC", "2026-10-01T09:12:33.123+0200", base + TimeSerial(7, 12, 33), ParseJiraDate("2026-10-01T09:12:33.123+0200")
    Check "text", "offset +02:00", "2026-10-01T09:12:33+02:00", base + TimeSerial(7, 12, 33), ParseJiraDate("2026-10-01T09:12:33+02:00")
    Check "text", "Z", "2026-10-01T09:12:33.000Z", base + TimeSerial(9, 12, 33), ParseJiraDate("2026-10-01T09:12:33.000Z")
    Check "text", "offset -0500", "2026-10-01T09:12:33.000-0500", base + TimeSerial(14, 12, 33), ParseJiraDate("2026-10-01T09:12:33.000-0500")
    Check "text", "date only", "2026-10-01", base, ParseJiraDate("2026-10-01")
    Check "text", "null", "Null", Empty, ParseJiraDate(Null)
    Check "text", "too short", "2026", Empty, ParseJiraDate("2026")
    Check "text", "not a date", "abcdefghijkl", Empty, ParseJiraDate("abcdefghijkl")
End Sub

Public Sub T_FieldToCell()
    Check "text", "user", "displayName", "Ann", FieldToCell(JsonParse(J("{'displayName':'Ann','accountId':'1'}")))
    Check "text", "option", "value", "High", FieldToCell(JsonParse(J("{'value':'High','id':'3'}")))
    Check "text", "array of objects", "components", "A; B", FieldToCell(JsonParse(J("[{'name':'A'},{'name':'B'}]")))
    Check "text", "empty array", "[]", "", FieldToCell(JsonParse("[]"))
    Check "text", "rich text (ADF)", "description", "Hi there" & vbLf, _
          FieldToCell(JsonParse(J("{'type':'doc','content':[{'type':'paragraph','content':[{'type':'text','text':'Hi '},{'type':'text','text':'there'}]}]}")))
    Check "text", "number stays a number", "5", 5, FieldToCell(5)
    Check "text", "null is empty", "null", Empty, FieldToCell(Null)
    Check "text", "unknown object as JSON", "{'x':1}", "{""x"":1}", FieldToCell(JsonParse(J("{'x':1}")))
    Check "text", "sprint array", "sprint", "PI3-S1; PI3-S2", _
          FieldToCell(JsonParse(J("[{'id':1,'name':'PI3-S1','state':'closed'},{'id':2,'name':'PI3-S2','state':'active'}]")))
End Sub

Public Sub T_SafeCell()
    Check "sheet", "formula-like text is escaped", "=SUM(A1)", "'=SUM(A1)", SafeCell("=SUM(A1)")
    Check "sheet", "minus text is escaped", "-5 days", "'-5 days", SafeCell("-5 days")
    Check "sheet", "any text gets the apostrophe", "abc", "'abc", SafeCell("abc")
    Check "sheet", "empty text unchanged", "(empty)", "", SafeCell("")
    Check "sheet", "long text cut to fit a cell", "40000 x", 32767, Len(SafeCell(String$(40000, "x")))
    Check "sheet", "number unchanged", "-5", -5, SafeCell(-5)
End Sub

' ================= Jira =================

Public Sub T_ProjectsJql()
    Dim c As New Collection
    c.Add "ABC": c.Add " DEF "
    Check "jira", "projects JQL", "ABC, DEF", "project in (""ABC"",""DEF"") ORDER BY key ASC", ProjectsJql(c)
    Dim raised As Boolean
    On Error Resume Next
    ProjectsJql New Collection
    raised = (Err.Number <> 0): Err.Clear
    On Error GoTo 0
    Check "jira", "no project raises", "(empty)", True, raised
End Sub

Public Sub T_SearchUrl()
    Dim u As String
    u = JiraSearchUrl("https://x.atlassian.net", "project = ABC", "")
    Check "jira", "first page URL", "project = ABC", _
          "https://x.atlassian.net/rest/api/3/search/jql?jql=project%20%3D%20ABC&maxResults=100&fields=%2Aall&expand=names", u
    u = JiraSearchUrl("https://x.atlassian.net", "project = ABC", "tok==")
    Check "jira", "next page has token", "tok==", True, (InStr(u, "&nextPageToken=tok%3D%3D") > 0)
End Sub

Private Function TestConfig() As Object
    Dim c As Object
    Set c = JiraConfig("https://x.atlassian.net/", "me@x.com", "t0k", "", "", False)
    c("RetryBaseMs") = 10: c("RetryMaxMs") = 50
    Set TestConfig = c
End Function

Public Sub T_SearchPagination()
    Dim cfg As Object, names As Object, issues As Collection, base As String
    Set cfg = TestConfig()
    Check "jira", "base URL trailing slash removed", "https://x.atlassian.net/", "https://x.atlassian.net", cfg("BaseUrl")
    FakeReset
    FakeAdd "GET", JiraSearchUrl(cfg("BaseUrl"), "project = ABC", ""), 200, _
        J("{'issues':[{'id':'1','key':'A-1','fields':{}},{'id':'2','key':'A-2','fields':{}}],'names':{'summary':'Summary'},'nextPageToken':'tok1','isLast':false}")
    FakeAdd "GET", JiraSearchUrl(cfg("BaseUrl"), "project = ABC", "tok1"), 200, _
        J("{'issues':[{'id':'3','key':'A-3','fields':{}}],'names':{'customfield_10016':'Story point estimate'},'isLast':true}")
    Set issues = JiraSearchAll(cfg, "project = ABC", names)
    Check "jira", "all pages read", "2 pages", 3, issues.Count
    Check "jira", "names merged", "2 pages", 2, names.Count
    Check "jira", "custom field name", "customfield_10016", "Story point estimate", names("customfield_10016")
    Check "jira", "HTTP calls", "2 pages", 2, gHttpLog.Count
    Check "jira", "auth header", "me@x.com:t0k", "Basic " & Base64Encode(Utf8Encode("me@x.com:t0k")), cfg("Authorization")
End Sub

Public Sub T_Retry429()
    Dim cfg As Object, resp As Object
    Set cfg = TestConfig()
    FakeReset
    FakeAdd "GET", "https://x/a", 429, "", "0"
    FakeAdd "GET", "https://x/a", 200, "{}"
    Set resp = HttpRequest(cfg, "GET", "https://x/a")
    Check "http", "retried after 429", "429 then 200", 200, resp("Status")
    Check "http", "attempts", "429 then 200", 2, resp("Attempts")
End Sub

Public Sub T_RetryExhausted()
    Dim cfg As Object, resp As Object, raised As Boolean, msg As String
    Set cfg = TestConfig()
    cfg("MaxRetries") = 2
    FakeReset
    FakeAdd "GET", "https://x/b", 429, "", ""
    FakeAdd "GET", "https://x/b", 429, "", ""
    FakeAdd "GET", "https://x/b", 429, "", ""
    Set resp = HttpRequest(cfg, "GET", "https://x/b")
    Check "http", "gives up after MaxRetries", "3 x 429", 3, resp("Attempts")
    On Error Resume Next
    CheckStatus resp, "test"
    raised = (Err.Number <> 0): msg = Err.Description: Err.Clear
    On Error GoTo 0
    Check "http", "CheckStatus raises", "429", True, raised
    Check "http", "message has status", "429", True, (InStr(msg, "HTTP 429") > 0)
    FakeReset
    FakeAdd "GET", "https://x/c", 503, "", "0"
    FakeAdd "GET", "https://x/c", 200, ""
    Check "http", "503 with Retry-After is retried", "503 then 200", 200, HttpRequest(cfg, "GET", "https://x/c")("Status")
    FakeReset
    FakeAdd "GET", "https://x/d", 503, "", ""
    Check "http", "503 without Retry-After is not retried", "503", 1, HttpRequest(cfg, "GET", "https://x/d")("Attempts")
    Check "http", "no queued response gives 404", "(none)", 404, HttpRequest(cfg, "GET", "https://x/none")("Status")
End Sub

Public Sub T_ChangelogChunks()
    Dim cfg As Object, keys As New Collection, i As Long, rows As Collection, url As String
    Set cfg = TestConfig()
    For i = 1 To 1200
        keys.Add "K-" & i
    Next i
    url = cfg("BaseUrl") & "/rest/api/3/changelog/bulkfetch"
    FakeReset
    FakeAdd "POST", url, 200, J("{'issueChangeLogs':[{'issueId':'10001','changeHistories':[{'id':'1','created':'2026-10-02T10:00:00.000+0200'," & _
        "'author':{'displayName':'Ann','accountId':'a1'},'items':[{'field':'status','fieldId':'status','fieldtype':'jira','from':'1','fromString':'Nouveau','to':'3','toString':'En analyse'}," & _
        "{'field':'Sprint','fieldId':'customfield_10020','fromString':'','toString':'PI3-S1'}]}]}],'nextPageToken':'p2'}")
    FakeAdd "POST", url, 200, J("{'issueChangeLogs':[{'issueId':'10001','changeHistories':[{'id':'2','created':'2026-10-05T10:00:00.000+0200','items':[{'fieldId':'status','fromString':'En analyse','toString':'Pret'}]}]}],'nextPageToken':null}")
    FakeAdd "POST", url, 200, J("{'issueChangeLogs':[{'issueId':'11200','changeHistories':[{'id':'3','created':'2026-10-06T10:00:00.000+0200','items':[{'fieldId':'status','fromString':'Nouveau','toString':'Annule'}]}]}]}")
    Set rows = JiraChangelogsAll(cfg, keys)
    Check "jira", "changelog rows", "3 responses", 4, rows.Count
    Check "jira", "POST calls (2 chunks, 1 extra page)", "1200 keys", 3, gHttpLog.Count
    Check "jira", "first chunk size", "call 1", 1000, CountOf(gHttpLog(1)(2), """K-")
    Check "jira", "first call has no token", "call 1", 0, CountOf(gHttpLog(1)(2), "nextPageToken")
    Check "jira", "second call sends token", "call 2", True, (InStr(gHttpLog(2)(2), """nextPageToken"":""p2""") > 0)
    Check "jira", "second chunk size", "call 3", 200, CountOf(gHttpLog(3)(2), """K-")
    Check "jira", "row fields", "row 1", "10001|1|status|Nouveau|En analyse|Ann", _
          rows(1)(0) & "|" & rows(1)(1) & "|" & rows(1)(5) & "|" & rows(1)(9) & "|" & rows(1)(11) & "|" & rows(1)(3)
    Check "jira", "missing author is empty", "row 3", Empty, rows(3)(3)
End Sub

Private Function YesNo(ByVal b As Boolean) As String
    ' Never write Boolean & "": French VBA gives "Vrai"/"Faux".
    YesNo = IIf(b, "True", "False")
End Function

Private Function CountOf(ByVal haystack As String, ByVal needle As String) As Long
    CountOf = (Len(haystack) - Len(Replace(haystack, needle, ""))) / Len(needle)
End Function

Public Sub T_IssueRows()
    Dim issues As New Collection, ids As Collection, row As Variant
    issues.Add JsonParse(J("{'id':'1','key':'A-1','fields':{'summary':'S1','status':{'name':'Nouveau'}}}"))
    issues.Add JsonParse(J("{'id':'2','key':'A-2','fields':{'summary':'S2','customfield_10016':3}}"))
    Set ids = IssueFieldIds(issues)
    Check "jira", "field union", "2 issues", "summary,status,customfield_10016", ids(1) & "," & ids(2) & "," & ids(3)
    row = IssueRow(issues(2), ids)
    Check "jira", "row key", "A-2", "A-2", row(0)
    Check "jira", "missing field is empty", "status on A-2", Empty, row(3)
    Check "jira", "number field", "customfield_10016", 3, row(4)
End Sub

Public Sub T_WriteTable()
    Dim ws As Object, rows As New Collection
    Set ws = ThisWorkbook.Worksheets("TestScratch")
    rows.Add Array("A-1", "=1+1", 5)
    rows.Add Array("A-2", Null, "-x")
    WriteTable ws, Array("Key", "Text", "Num"), rows
    Check "sheet", "header", "A1", "Key", ws.Range("A1").Value
    Check "sheet", "formula-like stays text", "B2", "=1+1", ws.Range("B2").Value
    Check "sheet", "no formula created", "B2", False, ws.Range("B2").HasFormula
    Check "sheet", "number cell", "C2", 5, ws.Range("C2").Value
    Check "sheet", "null cell empty", "B3", Empty, ws.Range("B3").Value
    Check "sheet", "minus text", "C3", "-x", ws.Range("C3").Value
End Sub

' ================= Flow =================

Public Sub T_StatusIntervals()
    Dim info As Object, changes As New Collection, cats As Object, out As Collection, d0 As Date
    d0 = DateSerial(2026, 10, 1)
    Set info = CreateObject("Scripting.Dictionary")
    info("K-1") = Array(d0, "Termine")
    info("K-2") = Array(d0 + 7, "Nouveau")
    info("K-3") = Array(d0, "En test")
    Set cats = CreateObject("Scripting.Dictionary")
    cats("Nouveau") = "Wait": cats("En analyse") = "Active"
    changes.Add Array("K-1", "c2", d0 + 4, "En analyse", "Termine")    ' out of order on purpose
    changes.Add Array("K-1", "c1", d0 + 1, "Nouveau", "En analyse")
    changes.Add Array("K-3", "c3", d0 + 1, "Nouveau", "En analyse")
    changes.Add Array("K-3", "c4", d0 + 2, "Pret", "Termine")           ' from does not match previous to
    changes.Add Array("K-3", "c0", d0 - 1, "Nouveau", "Nouveau")        ' before creation date
    Set out = BuildStatusIntervals(info, changes, d0 + 9, cats)
    Check "flow", "interval count", "K-1:3, K-2:1, K-3:4", 8, out.Count
    Check "flow", "K-1 first status", "row 1", "Nouveau|Wait|1", out(1)(2) & "|" & out(1)(3) & "|" & out(1)(6)
    Check "flow", "K-1 second status", "row 2", "En analyse|Active|3|c2", out(2)(2) & "|" & out(2)(3) & "|" & out(2)(6) & "|" & out(2)(8)
    Check "flow", "K-1 open last status", "row 3", "Termine|Unclassified|5|True", out(3)(2) & "|" & out(3)(3) & "|" & out(3)(6) & "|" & YesNo(out(3)(7))
    Check "flow", "K-2 no change", "row 4", "Nouveau|2|True", out(4)(2) & "|" & out(4)(6) & "|" & YesNo(out(4)(7))
    Check "flow", "K-3 before created flagged", "row 5", "BEFORE_CREATED", out(5)(9)
    Check "flow", "K-3 from mismatch flagged", "row 7", "FROM_MISMATCH(prev=En analyse)", out(7)(9)
    Check "flow", "K-3 current mismatch flagged", "row 8", "CURRENT_MISMATCH(now=En test)", out(8)(9)
End Sub

Public Sub T_StatusChanges()
    Dim rows As New Collection, map As Object, out As Collection
    rows.Add Array("10001", "1", "2026-10-02T10:00:00.000+0200", "Ann", "a1", "status", "status", "jira", "1", "Nouveau", "3", "En analyse")
    rows.Add Array("10001", "1", "2026-10-02T10:00:00.000+0200", "Ann", "a1", "customfield_10020", "Sprint", "custom", "", "", "", "PI3-S1")
    rows.Add Array("99999", "2", "2026-10-03T10:00:00.000Z", "Bob", "b1", "status", "status", "jira", "3", "En analyse", "4", "Pret")
    Set map = CreateObject("Scripting.Dictionary")
    map("10001") = "A-1"
    Set out = StatusChangesFromChangelog(rows, map)
    Check "flow", "only status rows", "3 rows", 2, out.Count
    Check "flow", "id mapped to key", "10001", "A-1", out(1)(0)
    Check "flow", "unknown id kept", "99999", "99999", out(2)(0)
    Check "flow", "date converted to UTC", "10:00+0200", DateSerial(2026, 10, 2) + TimeSerial(8, 0, 0), out(1)(2)
End Sub

Public Sub T_RetryDelay()
    Dim cfg As Object
    Set cfg = NewHttpConfig()
    Check "http", "Retry-After is a minimum", "Retry-After: 2", True, RetryDelayMs(cfg, 0, "2") >= 2000
    Check "http", "backoff capped", "attempt 10", True, RetryDelayMs(cfg, 10, "") <= 30000
End Sub

Public Sub T_SafeAccess()
    Dim d As Object
    Set d = JsonParse(J("{'a':{'b':null}}"))
    Check "json", "missing path is Empty", "a.x.y", True, IsEmpty(JPath(d, "a.x.y"))
    Check "json", "null value", "a.b", Null, JPath(d, "a.b")
    Check "json", "AsList on non-array", "Empty", 0, AsList(Empty).Count
    Check "json", "IsTrue(Null)", "Null", False, IsTrue(Null)
    Check "json", "JGet on scalar", "5", True, IsEmpty(JGet(5, "a"))
End Sub

Public Sub T_FilesUtf8()
    Dim folder As String, st As Object, files As Collection, s As String
    folder = ThisWorkbook.Path & "\tmp_files"
    If Len(Dir$(folder, vbDirectory)) = 0 Then MkDir folder
    s = "{""name"":""Caf" & ChrW(233) & """}"
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8": st.Open: st.WriteText s
    st.SaveToFile folder & "\page-0002.json", 2: st.SaveToFile folder & "\page-0001.json", 2
    st.Close
    Set files = ListFiles(folder, "page-*.json")
    Check "background", "files sorted", "page-0002, page-0001", "page-0001.json", files(1)
    Check "background", "UTF-8 file read", "Cafe", "Caf" & ChrW(233), JsonParse(ReadUtf8File(folder & "\" & files(1)))("name")
    Check "background", "done flag absent", folder, False, IsExportDone(folder)
End Sub

Public Sub T_TextNotConverted()
    ' Texts that Excel would turn into booleans, numbers, dates, errors or formulas stay text.
    Dim ws As Object, rows As New Collection, texts As Variant, c As Long
    texts = Array("true", "TRUE", "1.10", "2026-10-01", "1/2", "007", "50%", "#N/A", "=1+1", "-x", "PI3-S1")
    Set ws = ThisWorkbook.Worksheets("TestScratch")
    rows.Add texts
    WriteTable ws, texts, rows
    For c = 0 To UBound(texts)
        Check "sheet", "text stays text", CStr(texts(c)), CStr(texts(c)) & "|String", _
              CStr(ws.Cells(2, c + 1).Value) & "|" & TypeName(ws.Cells(2, c + 1).Value)
    Next c
End Sub

Public Sub T_StringifyBadObject()
    Dim raised As Boolean, s As String
    On Error Resume Next
    s = JsonStringify(ThisWorkbook)
    raised = (Err.Number <> 0): Err.Clear
    On Error GoTo 0
    Check "json", "unsupported object raises", "Workbook", True, raised
End Sub

Public Sub T_TooManyPages()
    Dim cfg As Object, names As Object, raised As Boolean, msg As String, c As Collection
    Set cfg = TestConfig()
    FakeReset
    FakeAdd "GET", JiraSearchUrl(cfg("BaseUrl"), "project = ABC", ""), 200, J("{'issues':[],'nextPageToken':'t'}")
    On Error Resume Next
    Set c = JiraSearchAll(cfg, "project = ABC", names, 1)
    raised = (Err.Number <> 0): msg = Err.Description: Err.Clear
    On Error GoTo 0
    Check "jira", "page guard raises", "maxPages=1", True, raised
    Check "jira", "page guard message", "maxPages=1", True, (InStr(msg, "Too many pages") > 0)
End Sub

Public Sub T_BackgroundExport()
    ' Option B without any process: simulate what the external PowerShell script writes.
    Dim folder As String, names As Object, issues As Collection
    folder = ThisWorkbook.Path & "\tmp_export"
    If Len(Dir$(folder, vbDirectory)) = 0 Then MkDir folder
    If Len(Dir$(folder & "\done.flag")) > 0 Then Kill folder & "\done.flag"
    WriteUtf8 folder & "\page-0001.json", J("{'issues':[{'id':'1','key':'A-1','fields':{}}],'names':{'summary':'Summary'}}")
    WriteUtf8 folder & "\page-0002.json", J("{'issues':[{'id':'2','key':'A-2','fields':{}}],'names':{'customfield_10016':'Story point estimate'}}")
    WatchExportFolder folder, 10
    Check "background", "watching", "start", "WATCHING", gExportStatus
    CheckBackgroundExport
    Check "background", "still watching without flag", "no done.flag", "WATCHING", gExportStatus
    WriteUtf8 folder & "\done.flag", "done"
    CheckBackgroundExport
    Check "background", "done flag seen", "done.flag", "DONE", gExportStatus
    Set issues = ImportSearchPages(folder, names)
    Check "background", "pages imported", "2 pages", 2, issues.Count
    Check "background", "names merged", "2 pages", 2, names.Count
    CancelBackgroundCheck
    WatchExportFolder folder & "\empty", 0
    WaitMs 1100
    CheckBackgroundExport
    Check "background", "timeout detected", "timeout 0 min", "TIMEOUT", gExportStatus
    CancelBackgroundCheck
End Sub

Private Sub WriteUtf8(ByVal path As String, ByVal text As String)
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8": st.Open: st.WriteText text
    st.SaveToFile path, 2: st.Close
End Sub

' ================= Network (public test API, not Jira) =================

Private Function NetConfig() As Object
    Dim c As Object
    Set c = NewHttpConfig()
    c("RetryBaseMs") = 10: c("RetryMaxMs") = 50: c("MaxRetries") = 1
    Set NetConfig = c
End Function

Public Sub N_HttpGetUtf8()
    Dim r As Object, v As String
    v = "Caf" & ChrW(233)
    Set r = HttpRequest(NetConfig(), "GET", "https://httpbin.org/get?q=" & UrlEncode(v))
    Check "net", "GET status", "httpbin /get", 200, r("Status")
    Check "net", "UTF-8 query echoed", "q=Cafe", v, JPath(JsonParse(r("Text")), "args.q")
End Sub

Public Sub N_HttpPostJson()
    Dim r As Object, body As Object
    Set body = CreateObject("Scripting.Dictionary")
    body("name") = "Caf" & ChrW(233)
    Set r = HttpRequest(NetConfig(), "POST", "https://httpbin.org/post", JsonStringify(body))
    Check "net", "POST status", "httpbin /post", 200, r("Status")
    Check "net", "UTF-8 JSON body echoed", "name=Cafe", "Caf" & ChrW(233), JPath(JsonParse(r("Text")), "json.name")
End Sub

Public Sub N_Http429()
    Dim r As Object
    Set r = HttpRequest(NetConfig(), "GET", "https://httpbin.org/status/429")
    Check "net", "429 returned", "httpbin /status/429", 429, r("Status")
    Check "net", "429 retried once", "MaxRetries=1", 2, r("Attempts")
End Sub

Public Sub N_HttpBadHost()
    Dim r As Object
    Set r = HttpRequest(NetConfig(), "GET", "https://no-such-host.invalid/")
    Check "net", "network error gives status 0", "bad host", 0, r("Status")
    Check "net", "error text filled", "bad host", True, Len(r("Error")) > 0
End Sub

' ================= Coverage =================

Public Function RunWithCoverage(ByVal includeNetwork As Boolean) As String
    Dim s As String
    mTotalPlanned = CountOfNames(TestNames()) + IIf(includeNetwork, CountOfNames(NetworkTestNames()), 0)
    mDone = 0
    ProgressDraw 0, mTotalPlanned
    CovStart
    s = RunAllTests()
    If includeNetwork Then s = s & " | net: " & RunNetworkTests()
    CovStop
    CovWrite ThisWorkbook.Worksheets("CoverageHits")
    mTotalPlanned = 0
    RunWithCoverage = s
End Function
