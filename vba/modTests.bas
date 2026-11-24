Option Explicit
' Test runner. Each test is a Public Sub named T_xxx. Results go to sheet "TestResults":
' TestId | Area | Description | Input | Expected | Actual | Result (OK/KO) | RunAt
' Offline tests use fixtures and the HTTP fake: they never call Jira.

Private mRows As Collection
Private mCurrent As String
Private mTotalPlanned As Long      ' tests planned for this run (progress bar total)
Private mDone As Long              ' tests finished in this run
Private mNextRow As Long           ' next free row in TestResults
Private mPacUrl As String          ' a proxy script served by run_vba_tests (SetPacUrl, used by T_ProxyScript)
Private mMockUrl As String         ' base URL of the fake Jira (set by MockSetup, used by the M_ tests)
Private mMockEmail As String       ' account the fake Jira accepts (Basic auth: email + API token)
Private mMockToken As String
Private mMockIssues As Long        ' issues the fake Jira holds
Private mMockChanges As Long       ' status changes the fake Jira holds

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
        "T_HttpNetworkAndCancel", "T_RetryExhausted", "T_ChangelogChunks", "T_IssueRows", "T_WriteTable", "T_StatusIntervals", "T_IntervalsUnreadableDate", "T_StatusChangesUnreadableDate", "T_ChangelogRows", "T_UserFieldIds", "T_DateShape", "T_ParseJiraDateEpoch", "T_DateKind", "T_IsPlausibleDate", "T_FindDateKey", "T_ChangelogDateKey", "T_SkippedDatesWarning", "T_SheetName", "T_ProposeFromSheets", "T_ProposeIterations", "T_ReadSprints", "T_TeamNames", "T_ReadStatusCatalog", "T_ReadWorkflowRules", "T_ReadLinkTypes", "T_ProposeTypes", "T_ProposeValuesNotAdmin", "T_HealthFields", "T_StatusRules", "T_RefreshRules", "T_WorkflowExceptions", "T_NotFeatures", "T_ReconciliationFindings", "T_CellKind", "T_FitsType", "T_ShapeOfCell", "T_TypeHints", "T_TypeProblems", "T_TypeEmptyTable", "T_TypeBlocks", "T_TypeFillLines", "T_WireTypes", _
        "T_StatusChanges", "T_RetryDelay", "T_SafeAccess", "T_FilesUtf8", "T_StringifyBadObject", _
        "T_TooManyPages", "T_BackgroundExport", "T_TextNotConverted", "T_ParamsTables", "T_ParamsOnWorkbook", _
        "T_TrimWhite", "T_ConfigSheet", "T_MandatoryConfig", "T_Locked", "T_PiIntervals", "T_ProxyDetect", "T_ProxyScript", "T_ProjectsSheet", "T_StatusNames", "T_ParseMoment", "T_ExtractFlowData", "T_ExtractJob", "T_ExtractCheckFields", "T_FlowBuild", "T_BaseUrlHttps", "T_Log", "T_LogInputsAndTimes", "T_StatusParts", "T_ErrorCodes", "T_LogHintAndOpen", "T_StepsSheet", "T_RovoPrompt", "T_RovoCsv", _
        "T_JiraConfig", "T_AbortExtraction", "T_Utf8Oracle", "T_Utf8Invalid", "T_Base64Oracle")
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
    ' The tests write to logs\vba-tests.log next to the test workbook, never to the real log; each test is framed by a line
    ' "TEST <name>" and "TEST <name> done | <time>": after a crash of Excel, the last TEST line names the test that ran.
    Dim nm As Variant, ok As Long, ko As Long, r As Variant, standalone As Boolean, started As Double, runnerLog As String
    Set mRows = New Collection
    runnerLog = LogFolderFrom(ThisWorkbook.Path, Environ$("LOCALAPPDATA"), Environ$("TEMP")) & "\vba-tests.log"
    If mDone = 0 Then                                               ' a new run of the tests: the log holds the last run only
        On Error Resume Next
        Kill runnerLog
        Err.Clear
        On Error GoTo 0
    End If
    gLogPath = runnerLog
    LogLine "TEST", "run of " & CountOfNames(names) & " tests in " & ThisWorkbook.Name
    Randomize 1
    standalone = (mTotalPlanned = 0)
    If standalone Then mTotalPlanned = CountOfNames(names): mDone = 0
    StartRows
    For Each nm In names
        Application.StatusBar = "Test " & (mDone + 1) & "/" & mTotalPlanned & ": " & nm & "..."
        mCurrent = nm
        Set gProxyFake = FakeProxySettings(False, "", "", "")   ' the proxy of this computer must not change a test
        gLogPath = runnerLog: LogLine "TEST", CStr(nm)
        started = Timer
        On Error Resume Next
        Err.Clear
        Application.Run "modTests." & nm
        If Err.Number <> 0 Then AddRow Array(nm & ".ERR", "runner", "Test raised an error", "", "no error", _
                                             Err.Number & " " & Err.Description, "KO", Now)
        Err.Clear
        On Error GoTo 0
        FakeOff
        gLogPath = runnerLog: LogLine "TEST", nm & " done | " & LogDuration(Timer - started)
        mDone = mDone + 1
        ProgressShow mDone, mTotalPlanned
    Next nm
    For Each r In mRows
        If r(6) = "OK" Then ok = ok + 1 Else ko = ko + 1
    Next r
    Set gProxyFake = Nothing
    Application.StatusBar = False
    If standalone Then mTotalPlanned = 0
    RunTests = ok & " OK, " & ko & " KO"
End Function

Private Sub StartRows()
    Dim ws As Object
    Set ws = ThisWorkbook.Worksheets(SheetName("TestResults"))
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
    Set ws = ThisWorkbook.Worksheets(SheetName("TestResults"))
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
    Set ws = ThisWorkbook.Worksheets(SheetName("TestResults"))
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
    Dim bad As Variant, raised As Boolean
    For Each bad In Array("", "[1,", "{'a':1 'b':2}", "{", "{'a':1}x", "{'a' 1}", "[1 2]", "{'a':tru}", "'abc", "{1:2}", "{'a':'\q'}", "@", "'\u12'", "'\uZZZZ'")
        raised = False
        On Error Resume Next
        Call JsonParse(J(bad))                      ' no assignment: a parsed object would raise 450 and hide an accepted input
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

Public Sub T_TrimWhite()
    Dim c As Object
    Check "text", "spaces", "'  ab  '", "ab", TrimWhite("  ab  ")
    Check "text", "tab and line breaks (a pasted token)", "tok CRLF", "tok", TrimWhite("tok" & vbCrLf)
    Check "text", "leading line break and tab", "LF TAB tok", "tok", TrimWhite(vbLf & vbTab & "tok")
    Check "text", "non-breaking space", "nbsp x nbsp", "x", TrimWhite(ChrW(160) & "x" & ChrW(160))
    Check "text", "inside is kept", "a b", "a b", TrimWhite(" a b ")
    Check "text", "only white gives empty", "'  '", "", TrimWhite("  " & vbCr)
    Check "text", "empty stays empty", "''", "", TrimWhite("")
    Set c = JiraConfig(" https://x.atlassian.net/ " & vbLf, " me@x.com" & vbTab, vbCrLf & "t0k " & vbCr, "", "", False)
    Check "jira", "base URL is trimmed", "spaces and LF around the URL", "https://x.atlassian.net", c("BaseUrl")
    Check "jira", "email and token are trimmed", "like $Token.Trim()", BasicAuthHeader("me@x.com", "t0k"), c("Authorization")
End Sub

Private Function ConfigRows(ByVal ws As Object, ByVal pairs As Variant) As Object
    ' Writes a Config sheet on ws: pairs = Array("Name", "Value", "Name", "Value", ...). Returns the sheet.
    Dim rows As New Collection, i As Long
    For i = LBound(pairs) To UBound(pairs) Step 2
        rows.Add Array(pairs(i), pairs(i + 1), "test")
    Next i
    WriteTable ws, Array("Parameter", "Value", "Description"), rows
    Set ConfigRows = ws
End Function

Private Function MandatoryRows(ByVal ws As Object, ByVal triples As Variant) As Object
    ' Writes a Config sheet with the column Mandatory: triples = Array("Name", "Value", "Mandatory", ...).
    Dim rows As New Collection, i As Long
    For i = LBound(triples) To UBound(triples) Step 3
        rows.Add Array(triples(i), triples(i + 1), triples(i + 2), "test")
    Next i
    WriteTable ws, Array("Parameter", "Value", "Mandatory", "Description"), rows
    Set MandatoryRows = ws
End Function

Private Function RequireError(ByVal ws As Object, ByVal onlyNames As String) As String
    ' "number|description" of RequireMandatory, or "" when it does not stop.
    On Error Resume Next
    RequireMandatory ws, onlyNames
    If Err.Number <> 0 Then RequireError = (Err.Number - vbObjectError) & "|" & Err.Description
    Err.Clear
End Function

Public Sub T_MandatoryConfig()
    ' The column Mandatory of the Config sheet: an empty mandatory value stops the buttons before any call to Jira.
    Dim ws As Object
    Set ws = ThisWorkbook.Worksheets("TestScratch")
    MandatoryRows ws, Array("BaseUrl", "", "True", "Email", "me@x.com", "True", "Filter", "", "False", _
                            "FlowLoadMode", " ", "VRAI", "FeatureTypes", "Feature", True, "AsOf", "", True)
    Check "mandatory", "empty mandatory values, in the order of the sheet", "BaseUrl, blank FlowLoadMode, AsOf", _
          "BaseUrl, FlowLoadMode, AsOf", MissingMandatory(ws)
    Check "mandatory", "an empty value that is not mandatory is fine", "Filter", 0, InStr(MissingMandatory(ws), "Filter")
    Check "mandatory", "limited to the connection", "BaseUrl,Email,ApiToken", "BaseUrl", MissingMandatory(ws, "BaseUrl,Email,ApiToken")
    Check "mandatory", "the stop names them", "error 608", "608|Fill these values of the Config sheet first (Mandatory = True): BaseUrl, FlowLoadMode, AsOf", _
          RequireError(ws, "")
    Check "mandatory", "nothing missing: no stop", "Email, Filter", "", RequireError(ws, "Email, Filter")
    ConfigRows ws, Array("BaseUrl", "", "AsOf", "")
    Check "mandatory", "a sheet without the column has nothing to check", "3 columns", "", MissingMandatory(ws)
    Check "mandatory", "the real Config sheet: the client values are mandatory and empty", "Config", _
          "BaseUrl, Email, ApiToken, AsOf", MissingMandatory(ThisWorkbook.Worksheets(SheetName("Config")))
End Sub

Public Sub T_Locked()
    ' A workbook open in Protected View or read-only stops the buttons that write it (error 612, J2D-SYS-003) with what to do.
    Dim num As Long, desc As String
    Check "locked", "free workbook", "false false", "", LockReason(False, False)
    Check "locked", "Protected View", "false true", "it is open in Protected View", LockReason(False, True)
    Check "locked", "read-only", "true false", "it is open read-only", LockReason(True, False)
    Check "locked", "Protected View comes first", "true true", "it is open in Protected View", LockReason(True, True)
    On Error Resume Next
    RequireEditableIf False, False
    Check "locked", "a free workbook is not stopped", "no error", 0, Err.Number
    RequireEditableIf True, False
    num = Err.Number: desc = Err.Description: Err.Clear
    On Error GoTo 0
    Check "locked", "the stop is error 612", "read-only", vbObjectError + 612, num
    Check "locked", "the stop says what to do", "read-only", True, (InStr(desc, "Enable Editing") > 0 And InStr(desc, "read-only") > 0 And InStr(desc, "Save As") > 0)
    Check "locked", "the support code", "612", "J2D-SYS-003", ErrorCode(num, desc)
    Check "locked", "the same file, whatever the case", "C:\A\b.xlsm", True, SameFile("C:\A\b.xlsm", "c:\a\B.XLSM")
    Check "locked", "another folder is another file", "C:\A\b.xlsm", False, SameFile("C:\A\b.xlsm", "C:\Downloads\b.xlsm")
End Sub

Private Function PiRaw(ByVal key As String, ByVal id As String, ByVal created As String, ByVal fieldId As String, _
                       ByVal fromText As String, ByVal toText As String) As Variant
    ' A row of the changelog as JiraChangelogChunk gives it: key, change id, created, ..., fieldId, ..., fromString, ..., toString.
    PiRaw = Array(key, id, created, "", "", fieldId, "PI", "custom", Empty, fromText, Empty, toText)
End Function

Private Function PiText(ByVal rows As Collection) As String
    ' "key|PI|from|to" per stretch, "open" for no end, joined with ";" (dates as yyyy-mm-dd).
    Dim r As Variant, out As String
    For Each r In rows
        out = out & IIf(Len(out) > 0, ";", "") & r(0) & "|" & r(1) & "|" & Format$(r(2), "yyyy-mm-dd") & "|" & _
              IIf(IsEmpty(r(3)), "open", Format$(r(3), "yyyy-mm-dd"))
    Next r
    PiText = out
End Function

Private Function PiStretches(ByVal changes As Variant, ByVal currentPi As String) As String
    ' The stretches of F-1, created on 2026-01-01, whose current PI is currentPi, after the changelog rows `changes` of the field PIF.
    Dim idToKey As Object, raw As New Collection, info As Object, c As Variant
    Set idToKey = CreateObject("Scripting.Dictionary")
    For Each c In changes
        raw.Add c
    Next c
    Set info = CreateObject("Scripting.Dictionary")
    info("F-1") = Array(DateSerial(2026, 1, 1), currentPi)
    PiStretches = PiText(PiIntervalRows(info, PiChangesFromChangelog(raw, idToKey, "PIF")))
End Function

Public Sub T_PiIntervals()
    ' The PI field of an issue through time (python: flow/pi_history.py, tests/unit/test_pi_history.py): the same cases.
    Dim raw As New Collection, idToKey As Object, kept As Collection
    Check "pi", "no change: the current PI from the creation", "PI-1", "F-1|PI-1|2026-01-01|open", PiStretches(Array(), "PI-1")
    Check "pi", "the value before the first change is its from", "none to PI-2 on 01-05", "F-1|PI-2|2026-01-05|open", _
          PiStretches(Array(PiRaw("F-1", "1", "2026-01-05T00:00:00.000+0000", "PIF", "", "PI-2")), "PI-2")
    Check "pi", "a PI moved ends the stretch, the end is the start of the next", "PI-3 to PI-5 on 01-03", _
          "F-1|PI-3|2026-01-01|2026-01-03;F-1|PI-5|2026-01-03|open", _
          PiStretches(Array(PiRaw("F-1", "1", "2026-01-03T00:00:00.000+0000", "PIF", "PI-3", "PI-5")), "PI-5")
    Check "pi", "a PI emptied ends the stretch and opens none", "PI-1 to empty on 01-03", "F-1|PI-1|2026-01-01|2026-01-03", _
          PiStretches(Array(PiRaw("F-1", "1", "2026-01-03T00:00:00.000+0000", "PIF", "PI-1", "")), "")
    Check "pi", "two changes in one second: by their id as a number (9 before 10)", "10 then 9 given", _
          "F-1|PI-1|2026-01-01|2026-01-03;F-1|PI-3|2026-01-03|open", _
          PiStretches(Array(PiRaw("F-1", "10", "2026-01-03T00:00:00.000+0000", "PIF", "PI-2", "PI-3"), _
                            PiRaw("F-1", "9", "2026-01-03T00:00:00.000+0000", "PIF", "PI-1", "PI-2")), "PI-3")
    Check "pi", "a change that writes the PI it had joins the stretch", "PI-1 to PI-1", "F-1|PI-1|2026-01-01|open", _
          PiStretches(Array(PiRaw("F-1", "1", "2026-01-03T00:00:00.000+0000", "PIF", "PI-1", "PI-1")), "PI-1")
    Check "pi", "blanks around a PI are dropped", " PI-1 ", "F-1|PI-1|2026-01-05|open", _
          PiStretches(Array(PiRaw("F-1", "1", "2026-01-05T00:00:00.000+0000", "PIF", "  ", " PI-1 ")), "PI-1")
    Set idToKey = CreateObject("Scripting.Dictionary")
    raw.Add PiRaw("F-1", "1", "not a date", "PIF", "PI-1", "PI-2")
    raw.Add PiRaw("F-1", "2", "2026-01-03T00:00:00.000+0000", "other", "a", "b")
    raw.Add PiRaw("F-1", "3", "2026-01-04T00:00:00.000+0000", "PIF", "PI-2", "PI-3")
    Set kept = PiChangesFromChangelog(raw, idToKey, "PIF")
    Check "pi", "only the field of PIField with a readable date is kept", "3 rows", "1|3", kept.Count & "|" & kept(1)(1)
    Check "pi", "no PIField: no change", "empty", 0, PiChangesFromChangelog(raw, idToKey, "").Count
End Sub

Public Sub SetPacUrl(ByVal url As String)
    ' run_vba_tests serves a proxy script (PAC) on this computer for T_ProxyScript and gives its address here.
    mPacUrl = url
End Sub

Public Sub T_ProxyDetect()
    ' ProxyServer empty: the proxy of the Windows settings, found as the browser finds it; DIRECT: none; typed: that one.
    Dim d As Object, ws As Object, c As Object, real As Object
    Set gProxyFake = FakeProxySettings(False, "", "fixed.proxy:8080", "<local>")
    Set d = DetectProxy("https://x.atlassian.net")
    Check "proxy", "the fixed proxy of the Windows settings", "Proxy fixed.proxy:8080", "fixed.proxy:8080|<local>", d("Server") & "|" & d("Bypass")
    Check "proxy", "the mode names no address", "fixed proxy", "the fixed proxy of the Windows settings", d("Mode")
    Set gProxyFake = FakeProxySettings(False, "", "", "")
    Set d = DetectProxy("https://x.atlassian.net")
    Check "proxy", "nothing in the Windows settings: direct", "all empty", "|direct (no proxy in the Windows settings)", d("Server") & "|" & d("Mode")
    ' A proxy script that cannot be downloaded (nothing listens on port 9): WinHTTP fails, the fixed proxy is used
    Set gProxyFake = FakeProxySettings(False, "http://127.0.0.1:9/none.pac", "fixed.proxy:8080", "")
    Set d = DetectProxy("https://x.atlassian.net")
    Check "proxy", "a script that cannot run: the fixed proxy", "PAC on port 9", "fixed.proxy:8080", d("Server")
    Check "proxy", "the mode says the script found nothing", "PAC on port 9", True, _
          d("Mode") Like "proxy script (PAC) found nothing (Windows error *), then the fixed proxy of the Windows settings"
    ' The Config sheet
    Set ws = ThisWorkbook.Worksheets("TestScratch")
    Set gProxyFake = FakeProxySettings(False, "", "http://fixed.proxy:8080/", "*.corp")
    ConfigRows ws, Array("BaseUrl", "https://x.atlassian.net", "ProxyServer", "")
    Set c = JiraConfigFromSheet(ws)
    Check "proxy", "empty ProxyServer: the Windows proxy, normalized", "ProxyServer empty", "fixed.proxy:8080|*.corp", c("ProxyServer") & "|" & c("ProxyBypass")
    ConfigRows ws, Array("BaseUrl", "https://x.atlassian.net", "ProxyServer", "", "ProxyBypass", "<local>")
    Set c = JiraConfigFromSheet(ws)
    Check "proxy", "a typed bypass stays", "ProxyBypass <local>", "fixed.proxy:8080|<local>", c("ProxyServer") & "|" & c("ProxyBypass")
    ConfigRows ws, Array("BaseUrl", "https://x.atlassian.net", "ProxyServer", "direct")
    Set c = JiraConfigFromSheet(ws)
    Check "proxy", "DIRECT: no proxy, whatever Windows says", "ProxyServer direct", "|direct (ProxyServer = DIRECT in the Config sheet)", c("ProxyServer") & "|" & c("ProxyMode")
    ConfigRows ws, Array("BaseUrl", "https://x.atlassian.net", "ProxyServer", "typed.proxy:3128")
    Set c = JiraConfigFromSheet(ws)
    Check "proxy", "a typed proxy wins", "ProxyServer typed", "typed.proxy:3128|the proxy of the Config sheet", c("ProxyServer") & "|" & c("ProxyMode")
    ' The real Windows settings of this computer: read without error (their values depend on the computer)
    Set gProxyFake = Nothing
    Set real = WindowsProxySettings()
    Check "proxy", "the real Windows settings are read", "WinHttpGetIEProxyConfigForCurrentUser", "AutoDetect AutoConfigUrl Proxy Bypass", Join(real.Keys, " ")
End Sub

Public Sub T_ProxyScript()
    ' A proxy script (PAC) served on this computer by run_vba_tests: WinHTTP downloads and runs it, and DetectProxy gives
    ' its proxy before the fixed one. Run from the buttons of the workbook there is no such server: nothing to check.
    Dim d As Object
    If Len(mPacUrl) = 0 Then Exit Sub
    Set gProxyFake = FakeProxySettings(False, mPacUrl, "fixed.proxy:8080", "")
    Set d = DetectProxy("https://x.atlassian.net/rest/api/3/myself")
    Check "proxy", "the proxy of the PAC script, before the fixed one", "PROXY 127.0.0.1:9", "127.0.0.1:9", d("Server")
    Check "proxy", "found by the script", "PAC", "proxy script (PAC): a proxy", d("Mode")
End Sub

Private Function HintOf(ByVal status As Long) As String
    Dim resp As Object, msg As String
    Set resp = CreateObject("Scripting.Dictionary")
    resp("Status") = status: resp("Attempts") = 1: resp("Error") = "": resp("Text") = ""
    On Error Resume Next
    CheckStatus resp, "test"
    msg = Err.Description: Err.Clear
    On Error GoTo 0
    HintOf = msg
End Function

Public Sub T_ConfigSheet()
    ' The connection and the proxy come from the Config sheet of the workbook, never from the code.
    Dim ws As Object, c As Object, wsReal As Object
    Set ws = ThisWorkbook.Worksheets("TestScratch")
    ConfigRows ws, Array("BaseUrl", "https://x.atlassian.net/ ", "Email", "me@x.com", "ApiToken", "t0k" & vbLf, _
                         "ProxyServer", "http://proxy.name.com:8080/", "ProxyBypass", "*.client.local;<local>", _
                         "ProxyWindowsAuth", "Oui", "TimeoutSeconds", "30", "MaxRetries", "2")
    Set c = JiraConfigFromSheet(ws)
    Check "config", "base URL", "trailing slash and space", "https://x.atlassian.net", c("BaseUrl")
    Check "config", "authorization from Email and ApiToken", "token + LF", BasicAuthHeader("me@x.com", "t0k"), c("Authorization")
    Check "config", "proxy without scheme", "http://proxy.name.com:8080/", "proxy.name.com:8080", c("ProxyServer")
    Check "config", "proxy bypass", "*.client.local;<local>", "*.client.local;<local>", c("ProxyBypass")
    Check "config", "Windows auth to the proxy: Oui", "ProxyWindowsAuth", True, c("ProxyWindowsAuth")
    Check "config", "timeout in ms", "TimeoutSeconds 30", 30000, c("ReceiveMs")
    Check "config", "connect and send follow", "TimeoutSeconds 30", 30000, c("ConnectMs") + c("SendMs") - 30000
    Check "config", "retries", "MaxRetries 2", 2, c("MaxRetries")
    ConfigRows ws, Array("BaseUrl", "https://y.atlassian.net", "ProxyWindowsAuth", "No", "TimeoutSeconds", "0", "MaxRetries", "0")
    Set c = JiraConfigFromSheet(ws)
    Check "config", "Windows auth off: No", "ProxyWindowsAuth", False, c("ProxyWindowsAuth")
    Check "config", "empty proxy stays empty", "ProxyServer missing", "", c("ProxyServer")
    Check "config", "a timeout below 1 falls back to 60 s", "TimeoutSeconds 0", 60000, c("ReceiveMs")
    Check "config", "0 retries is allowed", "MaxRetries 0", 0, c("MaxRetries")
    ConfigRows ws, Array("BaseUrl", "https://z.atlassian.net", "ProxyWindowsAuth", "", "MaxRetries", "")
    Set c = JiraConfigFromSheet(ws)
    Check "config", "empty cells take the defaults: Windows auth Yes", "ProxyWindowsAuth empty", True, c("ProxyWindowsAuth")
    Check "config", "default retries", "MaxRetries empty", 4, c("MaxRetries")
    Check "config", "default timeout", "TimeoutSeconds missing", 60000, c("SendMs")
    Check "config", "ConfigValue of an unknown parameter", "Nope", "fallback", ConfigValue(ws, "Nope", "fallback")
    Check "config", "yes in three languages", "Yes oui VRAI 1", "TTTT", Left$(IIf(IsYes("Yes"), "T", "F") & IIf(IsYes(" oui "), "T", "F") & IIf(IsYes("VRAI"), "T", "F") & IIf(IsYes("1"), "T", "F"), 4)
    Check "config", "no is not yes", "No non empty 0", "FFFF", IIf(IsYes("No"), "T", "F") & IIf(IsYes("non"), "T", "F") & IIf(IsYes(""), "T", "F") & IIf(IsYes("0"), "T", "F")
    Check "config", "proxy typed with HTTP:// and slashes", "HTTP://P:1//", "P:1", NormalizeProxy("HTTP://P:1//")
    Check "config", "proxy typed with https://", "https://p:2", "p:2", NormalizeProxy("https://p:2")
    Check "config", "proxy typed plain", " p:3 ", "p:3", NormalizeProxy(" p:3 ")
    Check "config", "no proxy", "", "", NormalizeProxy("")
    Check "config", "407 points to the proxy cells", "HTTP 407", True, InStr(HintOf(407), "ProxyWindowsAuth") > 0
    Check "config", "401 points to Email and ApiToken", "HTTP 401", True, InStr(HintOf(401), "ApiToken") > 0
    Check "config", "no answer points to BaseUrl and proxy", "status 0", True, InStr(HintOf(0), "BaseUrl") > 0
    Check "config", "a 500 has no connection hint", "HTTP 500", False, InStr(HintOf(500), "Hint") > 0
    Set wsReal = ThisWorkbook.Worksheets(SheetName("Config"))
    Check "config", "the real Config sheet keeps client values empty", "BaseUrl", "", ConfigValue(wsReal, "BaseUrl")
    Check "config", "the real Config sheet defaults", "ProxyWindowsAuth", "Yes", ConfigValue(wsReal, "ProxyWindowsAuth")
    Check "config", "JiraConfigFromWorkbook reads the sheet", "MaxRetries", 4, JiraConfigFromWorkbook()("MaxRetries")
End Sub

Private Function JoinNames(ByVal names As Collection) As String
    ' "|a|b|c": the names of a collection as text, whatever its size (a check must not raise on a short list).
    Dim i As Long
    For i = 1 To names.Count
        JoinNames = JoinNames & "|" & CStr(names(i))
    Next i
End Function

Private Function StatusBook(ByVal mode As Long) As Object
    ' A scratch workbook as the build sees it: Projects (a Program and a Team), Config, Issues (Key, Type, Project) and StatusIntervals (Key, Seq, Status, Category)
    ' written as raw values. P-1 and P-2 are Features of the Program (P-2 went Review first, a ticket created late), P-3 is a sub-task of the Program, T-1 a
    ' Feature-typed ticket of the Team project: only the first two feed the stack. withJira adds the sheets of the Jira configuration: the Feature type follows
    ' the workflow WF-F (Todo -> Build -> Review -> Done), the sub-task WF-S. mode: 0 no Jira sheet, 1 the sheets above, 2 also a second Feature type (Enabler)
    ' that follows WF-G (Todo -> Review: Review at depth 1 there, 2 in WF-F).
    Dim wb As Object
    Set wb = Workbooks.Add
    PutRows wb, "Projects", Array(Array("Project", "Role", "Team"), Array("P", "Program", ""), Array("T", "Team", "Alpha"))
    PutRows wb, "Config", Array(Array("Parameter", "Value", "Example", "Mandatory", "Description"), Array("FeatureTypes", IIf(mode = 2, "Feature,Enabler", "Feature"), "Feature", "True", ""))
    PutRows wb, "Issues", Array(Array("Key", "Type", "Project"), Array("P-1", "Feature", "P"), Array("P-2", "Feature", "P"), Array("P-3", "Sub-task", "P"), Array("T-1", "Feature", "T"))
    PutRows wb, "StatusIntervals", Array(Array("Key", "Seq", "Status", "Category"), _
        Array("P-1", 1, "Todo", "Backlog"), Array("P-1", 2, "Build", "Active"), Array("P-1", 3, "Review", "Wait"), Array("P-1", 4, "Build", "Active"), Array("P-1", 5, "Done", "Done"), _
        Array("P-2", 1, "Review", "Wait"), Array("P-2", 2, "Build", "Active"), Array("P-2", 3, "Extra", "Active"), _
        Array("P-3", 1, "Sub", "Active"), Array("T-1", 1, "Team", "Active"), Array("P-1", 6, "Odd", "Unclassified"))
    If mode >= 1 Then
        PutRows wb, "JiraIssueTypes", Array(Array("Project", "IssueType", "Workflow"), Array("P", "Feature", "WF-F"), Array("P", "Sub-task", "WF-S"), Array("T", "Feature", "WF-S"), _
            Array("P", "Enabler", IIf(mode = 2, "WF-G", "")))
        PutRows wb, "JiraWorkflows", Array(Array("Workflow", "Status", "Initial"), Array("WF-F", "Todo", 1), Array("WF-F", "Build", ""), Array("WF-F", "Review", ""), _
            Array("WF-F", "Done", ""), Array("WF-S", "Sub", 1), Array("WF-G", "Todo", 1), Array("WF-G", "Review", ""))
        PutRows wb, "JiraTransitions", Array(Array("Workflow", "From", "To"), Array("WF-F", "Todo", "Build"), Array("WF-F", "Build", "Review"), Array("WF-F", "Review", "Done"), _
            Array("WF-S", "Sub", "Todo"), Array("WF-G", "Todo", "Review"))
    End If
    Application.DisplayAlerts = False
    wb.Worksheets(1).Delete
    Application.DisplayAlerts = True
    Set StatusBook = wb
End Function

Private Sub CloseBook(ByVal wb As Object)
    Application.DisplayAlerts = False
    wb.Close False
    Application.DisplayAlerts = True
End Sub

Public Sub T_StatusNames()
    ' The stack of the chart Cycle time by status: the Active and Wait statuses the Features visited, by depth in the Jira workflow, then by the earliest Seq,
    ' then by name (flow/pi_status_time.feature_statuses); and the position of the selector list.
    Dim wb As Object, names As Collection
    Set wb = StatusBook(0)
    Set names = StatusNames(wb)
    Check "statusnames", "no Jira sheet: by the earliest Seq (Review first at 1 for P-2, Build 2, Extra 3), the sub-task, the Team project and the Backlog, Done and Unclassified rows are left out", _
          "Review, Build, Extra", "3|Review|Build|Extra", names.Count & JoinNames(names)
    CloseBook wb
    Set wb = StatusBook(1)
    Set names = StatusNames(wb)
    Check "statusnames", "with the Jira workflow of the Feature type: by depth (Build 1, Review 2), a status that no move reaches last", "Build, Review, Extra", "3|Build|Review|Extra", names.Count & JoinNames(names)
    CloseBook wb
    Set wb = StatusBook(2)
    Set names = StatusNames(wb)
    Check "statusnames", "two Feature types: a status takes its smallest depth (Review 1 in WF-G, tied with Build: the earliest Seq decides)", "Review, Build, Extra", "3|Review|Build|Extra", names.Count & JoinNames(names)
    CloseBook wb
    Set wb = StatusBook(1)
    PutRows wb, "JiraIssueTypes", Array(Array("Project", "IssueType", "Workflow"), Array("P", "Feature", ""), Array("T", "Feature", "WF-F"))
    Set names = StatusNames(wb)
    Check "statusnames", "a Feature type whose workflow is unknown, or that only a Team project follows, gives no depth: the Seq order", "Review, Build, Extra", "3|Review|Build|Extra", names.Count & JoinNames(names)
    PutRows wb, "StatusIntervals", Array(Array("Key", "Seq", "Status", "Category"), Array("P-3", 1, "Sub", "Active"))
    Set names = StatusNames(wb)
    Check "statusnames", "no Feature visited an Active or Wait status: one placeholder", "(no status)", "1|(no status)", names.Count & JoinNames(names)
    CloseBook wb
    Set wb = StatusBook(0)
    PutRows wb, "Issues", Array(Array("Key", "Type", "Project"))
    Set names = StatusNames(wb)
    Check "statusnames", "no issue at all: one placeholder, nothing raised", "(no status)", "1|(no status)", names.Count & JoinNames(names)
    CloseBook wb
    Check "statusnames", "the depth of a status is its number of moves from the creation status, a loop is visited once, an unreachable status has none", "a 0, b 1, c 2", "0|1|2|-", _
          DepthsText(FlowDepths("a", MovesOf("a>b;b>a;b>c;x>a")), "a", "b", "c", "x")
    Check "statusnames", "without a creation status there is no depth", "empty", 0, FlowDepths("", MovesOf("a>b")).Count
    Check "statusnames", "the selector list is right of the widest block: the teams", "9 teams, 4 Feature teams, 6 statuses", 13, SelectorColumn(9, 4, 6)
    Check "statusnames", "the selector list is right of the widest block: the statuses", "3 teams, 2 Feature teams, 12 statuses", 16, SelectorColumn(3, 2, 12)
    Check "statusnames", "the selector list is right of the widest block: the teams of the Features", "3 teams, 12 Feature teams, 4 statuses", 16, SelectorColumn(3, 12, 4)
    Check "statusnames", "the selector list starts at column 10 at least", "2 teams, 1 Feature team, 3 statuses", 10, SelectorColumn(2, 1, 3)
End Sub

Private Function MovesOf(ByVal text As String) As Collection
    ' "a>b;b>c" as moves: Array(from, to) each.
    Dim out As New Collection, part As Variant, p() As String
    For Each part In Split(text, ";")
        p = Split(CStr(part), ">")
        out.Add Array(p(0), p(1))
    Next part
    Set MovesOf = out
End Function

Private Function DepthsText(ByVal depths As Object, ParamArray statuses() As Variant) As String
    ' The depth of each status, "-" for one without depth, joined with |.
    Dim i As Long, parts As New Collection, s As String
    For i = LBound(statuses) To UBound(statuses)
        parts.Add IIf(depths.Exists(CStr(statuses(i))), CStr(depths(CStr(statuses(i)))), "-")
    Next i
    For i = 1 To parts.Count
        s = s & IIf(i > 1, "|", "") & parts(i)
    Next i
    DepthsText = s
End Function

Public Sub T_ProjectsSheet()
    ' Which projects to extract and which one is the Program: a sheet of the workbook, read by the macros.
    Dim ws As Object, rows As New Collection, keys As Collection, wsReal As Object, raised As Boolean, jql As String
    Set ws = ThisWorkbook.Worksheets("TestScratch")
    rows.Add Array("AAA", "Team", "Alpha")
    rows.Add Array("TTT", "program", Null)
    rows.Add Array(" BBB ", "Team", " Beta ")
    WriteTable ws, Array("Project", "Role", "Team"), rows
    Set keys = ProjectKeys(ws)
    Check "projects", "the Program comes first, then the rows in order", "AAA, TTT, BBB", "TTT|AAA|BBB", keys(1) & "|" & keys(2) & "|" & keys(3)
    Check "projects", "Program project (case ignored)", "role 'program'", "TTT", ProgramProject(ws)
    Check "projects", "team of a project is trimmed", "BBB", "Beta", TeamOfProject(ws, "BBB")
    Check "projects", "the Program has no team", "TTT", "", TeamOfProject(ws, "TTT")
    Check "projects", "unknown project has no team", "ZZZ", "", TeamOfProject(ws, "ZZZ")
    rows.Remove 2
    WriteTable ws, Array("Project", "Role", "Team"), rows
    Check "projects", "no Program gives empty text", "AAA, BBB", "", ProgramProject(ws)
    Set wsReal = ThisWorkbook.Worksheets(SheetName("Projects"))
    Check "projects", "real sheet: the Program", "first key of the script", "TTT", ProgramProject(wsReal)
    Check "projects", "real sheet: nine projects", "Projects", 9, ProjectKeys(wsReal).Count
    Check "projects", "real sheet: the Program is first", "Projects", "TTT", ProjectKeys(wsReal)(1)
    Check "projects", "real sheet: team of AAA", "AAA", "AAA", TeamOfProject(wsReal, "AAA")
    ' JQL: projects, then the optional Filter of the Config sheet
    Set keys = New Collection
    keys.Add "TTT": keys.Add "AAA"
    Check "projects", "JQL without filter", "TTT, AAA", "project in (""TTT"",""AAA"") ORDER BY key ASC", ProjectsJql(keys)
    Check "projects", "JQL with a filter", "updated >= -90d", "project in (""TTT"",""AAA"") AND (updated >= -90d) ORDER BY key ASC", _
          ProjectsJql(keys, "  updated >= -90d" & vbLf)
    On Error Resume Next
    jql = ProjectsJql(keys, "updated >= -90d ORDER BY created")
    raised = (Err.Number <> 0): Err.Clear
    On Error GoTo 0
    Check "projects", "a filter with ORDER BY is refused", "ORDER BY created", True, raised
    ConfigRows ws, Array("Filter", "issuetype = Feature")
    Check "projects", "the extraction JQL reads Projects and Config", "Filter issuetype = Feature", _
          "project in (""TTT"",""AAA"",""EEE"",""FFFF"",""TTTT"",""SSSS"",""SSS"",""TETET"",""CCCC"") AND (issuetype = Feature) ORDER BY key ASC", _
          ExtractionJql(wsReal, ws)
    ConfigRows ws, Array("BaseUrl", "x")
    Check "projects", "no filter: projects only", "Filter missing", True, InStr(ExtractionJql(wsReal, ws), " AND ") = 0
    Check "projects", "ExtractionJqlFromWorkbook reads the real sheets", "empty Filter", True, InStr(ExtractionJqlFromWorkbook(), """TTT"",""AAA""") > 0
End Sub

Private Function RaisesOnConfig(ByVal url As String) As Boolean
    On Error Resume Next
    Call JiraConfig(url, "me@x.com", "t", "", "", False)
    RaisesOnConfig = (Err.Number = vbObjectError + 602)
    Err.Clear
End Function

Public Sub T_BaseUrlHttps()
    ' The Basic header holds the API token: it must not travel over http, except to the fake Jira on this computer.
    Dim raised As Boolean
    On Error Resume Next
    Call JiraConfig("http://x.atlassian.net", "me@x.com", "t", "", "", False)
    raised = (Err.Number = vbObjectError + 602): Err.Clear
    On Error GoTo 0
    Check "config", "an http BaseUrl is refused", "http://x.atlassian.net", True, raised
    Check "config", "https is accepted", "https://x.atlassian.net", "https://x.atlassian.net", JiraConfig("https://x.atlassian.net", "me@x.com", "t", "", "", False)("BaseUrl")
    Check "config", "the fake Jira on this computer is accepted", "http://127.0.0.1:8099", "http://127.0.0.1:8099", JiraConfig("http://127.0.0.1:8099/", "me@x.com", "t", "", "", False)("BaseUrl")
    Check "config", "a name that only starts like localhost is not local", "http://localhost.evil.example", True, RaisesOnConfig("http://localhost.evil.example")
    Check "config", "127.0.0.1 inside a longer name is not local", "http://127.0.0.1.evil.example", True, RaisesOnConfig("http://127.0.0.1.evil.example")
    Check "config", ".invalid in the query is not a local host", "http://evil.example/?x=.invalid", True, RaisesOnConfig("http://evil.example/?x=.invalid")
    Check "config", "a .invalid host (the fake proxy tests) is local", "http://jira.invalid:8080/x", False, RaisesOnConfig("http://jira.invalid:8080/x")
    Check "config", "an empty BaseUrl is left to the later check", "", "", JiraConfig("", "me@x.com", "t", "", "", False)("BaseUrl")
    Check "extract", "only the fields the metrics use", "WorkTypeField customfield_10200", "key,project,issuetype,created,status,parent,resolutiondate,fixVersions,issuelinks,customfield_10200", ExtractFields("customfield_10200")
    Check "extract", "no work type field: not asked", "empty", "key,project,issuetype,created,status,parent,resolutiondate,fixVersions,issuelinks", ExtractFields("")
    Check "extract", "the configured fields of the data checks are asked, the empty ones are not", "Team and Sprint", _
          "key,project,issuetype,created,status,parent,resolutiondate,fixVersions,issuelinks,customfield_10400,customfield_10020", _
          ExtractFields("", CheckFields("", "", "", "customfield_10400", "customfield_10020"))
End Sub

Public Sub T_LogHintAndOpen()
    ' After an error: the status names the log file and Notepad is asked to open it (a test only records the command).
    Dim saved As String
    saved = gLogPath
    gLogPath = "C:/x y/run.log"
    gLogOpenFake = True: gLogOpenCommand = ""
    Check "log", "the hint names the log file, as a field of the status", "path", " | Log: C:/x y/run.log", LogHint()
    LogOpen
    Check "log", "the log opens in Notepad, the path quoted", "command", "notepad.exe ""C:/x y/run.log""", gLogOpenCommand
    gLogOpenCommand = "": LogOpen False
    Check "log", "OpenLogOnError = No: Notepad is not asked", "command", "", gLogOpenCommand
    gLogOpenFake = False
    LogOpen                                          ' no person at the keyboard (a test Excel): nothing starts
    Check "log", "an Excel started by a script does not start Notepad", "UserControl", False, Application.UserControl
    gLogPath = saved
End Sub

Public Sub T_RovoCsv()
    ' The CSV answered by Rovo: the reader (quotes, line breaks in quotes, the comma, the semicolon and the tab), the items it
    ' keeps (UNKNOWN and the token are not values), and the load into the sheets Config, Projects and Iterations (saved and
    ' restored: other tests read them).
    Dim recs As Collection, res As Object, text As String, summary As String, saved(1 To 3) As Variant, addr(1 To 3) As String
    Dim names As Variant, i As Long, wsP As Object, wsC As Object, rowProxy As Long, f As Long, file As String, bytes() As Byte
    Set recs = ParseCsv("a,""b,1"",""say """"hi""""""" & vbCrLf & vbCrLf & "x,,y" & vbLf & "last")
    Check "csv", "a quoted comma and a doubled quote", "b,1|say ""hi""", "b,1|say ""hi""", recs(1)(1) & "|" & recs(1)(2)
    Check "csv", "an empty line is no record; CRLF and LF both end one", "3 records", 3, recs.Count
    Check "csv", "an empty value between two commas", "x,,y", "x||y", recs(2)(0) & "|" & recs(2)(1) & "|" & recs(2)(2)
    Check "csv", "the last record has no line break", "last", "last", recs(3)(0)
    Set recs = ParseCsv("a;""b;c""" & vbLf & "1;2")
    Check "csv", "the semicolon when it is the most frequent", "a;b;c", "a|b;c|1|2", recs(1)(0) & "|" & recs(1)(1) & "|" & recs(2)(0) & "|" & recs(2)(1)
    Set recs = ParseCsv("a" & vbTab & "b" & vbLf & "1" & vbTab & "2")
    Check "csv", "the tab when it is the most frequent", "tab", 2, UBound(recs(1)) + 1
    Set recs = ParseCsv("a,""line one" & vbLf & "line two""" & vbLf & "b")
    Check "csv", "a line break inside quotes stays in the value", "2 records", "2|line one" & vbLf & "line two", recs.Count & "|" & recs(1)(1)
    Check "csv", "an empty text has no record", "empty", 0, ParseCsv("").Count
    ' the items
    text = "Sheet,Row,Column,Value,Note" & vbLf & "Config,,FillRateThreshold,85,seen in the page of the ART" & vbLf & _
           "Config,,ApiToken,TO BE TYPED BY THE USER," & vbLf & "Config,,Filter,UNKNOWN,check" & vbLf & "Config,,SprintField,,not found" & vbLf & _
           "Projects,1,Project,ZZA,key" & vbLf & "Projects,1,Role,Program," & vbLf & "Projects,1,Team,Train Z," & vbLf & _
           "Projects,2,Project,ZZB," & vbLf & "Projects,2,Role,Team," & vbLf & "Projects,2,Team,UNKNOWN," & vbLf & _
           "Iterations,1,Iteration-label,PI#01-2026 It1," & vbLf & "Iterations,1,Start,2026-10-05," & vbLf & "Iterations,1,End,2026-10-26," & vbLf & _
           "Elsewhere,1,Name,x," & vbLf & "Projects,two,Project,x," & vbLf & "Config,,BaseUrl,https://evil.example,injected" & vbLf & "Projects,3,Nonsense,x,"
    Set res = RovoCsvRead(ThisWorkbook, text)
    Check "csv", "Config: only a real value is an item (UNKNOWN and the token are not)", "items", 11, res("items").Count
    Check "csv", "UNKNOWN and the token are counted apart from the empty values", "unknown|empty", "2|1", res("unknown") & "|" & res("empty")
    Check "csv", "a sheet that the import does not fill is a problem", "Elsewhere", True, InStr(res("problems")(2), "'Elsewhere' is not one that this import fills") > 0
    Check "csv", "a Row that is not a number is a problem", "two", True, InStr(res("problems")(3), "Row must be a number for the sheet Projects") > 0
    Check "csv", "a connection parameter is a problem, not an item", "BaseUrl", True, InStr(res("problems")(4), "BaseUrl says where the credentials go") > 0
    Check "csv", "the token placeholder of a connection parameter is a problem too", "ApiToken", True, InStr(res("problems")(1), "ApiToken says where the credentials go") > 0
    Set res = RovoCsvRead(ThisWorkbook, "Sheet,Row,Column,Value" & vbLf & "Config,,Base" & ChrW(173) & "Url,https://evil.example" & vbLf & "Config,,BaseUrl" & ChrW(&HFF21) & ",x")
    Check "csv", "a Config name outside ASCII (a soft hyphen, a full-width letter) is refused, never matched loosely", "items", 0, res("items").Count
    Check "csv", "and told as a problem", "ASCII", True, InStr(res("problems")(1), "must be plain ASCII") > 0
    Set res = RovoCsvRead(ThisWorkbook, "Sheet,Value" & vbLf & "Config,x")
    Check "csv", "the columns Sheet, Row, Column and Value are needed", "header", True, InStr(res("problems")(1), "must name the columns Sheet, Row, Column and Value") > 0
    Set res = RovoCsvRead(ThisWorkbook, "")
    Check "csv", "an empty file is a problem", "empty", "The file has no line.", res("problems")(1)
    ' the load: Config, Projects, Iterations are saved, written, read back and restored
    names = Array("Config", "Projects", "Iterations")
    For i = 1 To 3
        saved(i) = ThisWorkbook.Worksheets(SheetName(names(i - 1))).UsedRange.Value
        addr(i) = ThisWorkbook.Worksheets(SheetName(names(i - 1))).UsedRange.Address
    Next i
    Set wsC = ThisWorkbook.Worksheets(SheetName("Config")): Set wsP = ThisWorkbook.Worksheets(SheetName("Projects"))
    For rowProxy = 2 To 60
        If CStr(wsC.Cells(rowProxy, 1).Value & "") = "FillRateThreshold" Then Exit For
    Next rowProxy
    file = Environ$("TEMP") & "\j2d_rovo_csv_" & Format$(Now, "hhnnss") & Int(Rnd() * 1000) & ".csv"     ' not next to the workbook: any folder
    f = FreeFile
    Open file For Binary Access Write As #f
    bytes = Utf8Encode(ChrW(&HFEFF) & text)                         ' a BOM, as Excel writes it
    Put #f, , bytes
    Close #f
    Set res = RovoCsvRead(ThisWorkbook, "Sheet,Row,Column,Value" & vbLf & "Config,,Filter,x")
    res("items").Add Array("Config", 0, "Email", "me@evil.example")           ' an item that got past the reader: the writer refuses it
    RovoCsvApply ThisWorkbook, res
    Check "csv", "the writer refuses a connection parameter by the name of the row it found", "Email", True, Len(CStr(wsC.Cells(FindRowOf(wsC, "Email"), 2).Value & "")) = 0
    Check "csv", "and says so", "problem", True, InStr(res("problems")(1), "Email says where the credentials go") > 0
    summary = ImportRovoCsv(ThisWorkbook, file)
    Check "csv", "Config: the Value of the parameter is written", "FillRateThreshold", "85", CStr(wsC.Cells(rowProxy, 2).Value)
    Check "csv", "a connection parameter is never loaded", "BaseUrl", True, InStr(CStr(wsC.Cells(FindRowOf(wsC, "BaseUrl"), 2).Value & ""), "evil") = 0
    Check "csv", "and is told as a problem", "BaseUrl", True, InStr(summary, " | Problems: ") > 0
    Check "csv", "Config: the other parameters are left as they were", "ApiToken", True, Len(CStr(wsC.Cells(FindRowOf(wsC, "ApiToken"), 2).Value & "")) = 0
    Check "csv", "a table sheet: the rows are the rows of the CSV", "Projects", "ZZA|Program|Train Z|ZZB|Team||", _
          wsP.Range("A2").Value & "|" & wsP.Range("B2").Value & "|" & wsP.Range("C2").Value & "|" & wsP.Range("A3").Value & "|" & wsP.Range("B3").Value & "|" & wsP.Range("C3").Value & "|" & wsP.Range("A4").Value
    Check "csv", "the old rows are gone", "row 4", True, IsEmpty(wsP.Range("A4").Value)
    Check "csv", "an ISO date stays text (the readers of the sheet parse it)", "Start", "2026-10-05", CStr(ThisWorkbook.Worksheets(SheetName("Iterations")).Range("B2").Value)
    Check "csv", "the summary says what was loaded", "summary", True, _
          InStr(summary, "Settings CSV loaded | Config_from_user: 1 value(s): FillRateThreshold | Projects_from_user: 2 row(s) | Iterations_from_user: 1 row(s) | Empty values: 1 (not loaded) | UNKNOWN values: 2 (not loaded)") = 1
    Check "csv", "and the first problem", "problems", True, InStr(summary, " | Problems: ") > 0 And InStr(summary, ", first: Line ") > 0
    For i = 1 To 3
        With ThisWorkbook.Worksheets(SheetName(names(i - 1)))
            .UsedRange.ClearContents
            .Range(addr(i)).Value = saved(i)
        End With
    Next i
    On Error Resume Next
    Kill file
    Err.Clear
    On Error GoTo 0
End Sub

Private Function FindRowOf(ByVal ws As Object, ByVal name As String) As Long
    Dim r As Long
    For r = 2 To 80
        If CStr(ws.Cells(r, 1).Value & "") = name Then FindRowOf = r: Exit Function
    Next r
End Function

Public Sub T_RovoPrompt()
    ' The prompt for Rovo: the lines of RovoPromptSpec in order, the parameters of Config read live (name, mandatory, meaning,
    ' example), the header of each sheet, and never a Value of Config (it may hold the address, the e-mail or the token).
    Dim lines As Collection, text As String, ln As Variant, ws As Object, wsC As Object, rowBaseUrl As Long, r As Long, saved As Variant
    Dim n As Long, firstText As String
    Set wsC = ThisWorkbook.Worksheets(SheetName("Config"))
    For r = 2 To 60
        If CStr(wsC.Cells(r, 1).Value & "") = "BaseUrl" Then rowBaseUrl = r: Exit For
    Next r
    saved = wsC.Cells(rowBaseUrl, 2).Value
    wsC.Cells(rowBaseUrl, 2).Value = "https://SECRET-HOST.example.net"
    Set lines = RovoPromptLines(ThisWorkbook)
    For Each ln In lines
        text = text & ln & vbLf
    Next ln
    wsC.Cells(rowBaseUrl, 2).Value = saved
    Check "rovo", "the first line is the first row of the spec", "intro", CStr(ReadTable(ThisWorkbook.Worksheets(SheetName(SHEET_ROVO_SPEC)))(1)("Text")), CStr(lines(1))
    Check "rovo", "a parameter of Config: name, mandatory, meaning", "FlowLoadMode", True, InStr(text, "- FlowLoadMode (mandatory): ") > 0
    Check "rovo", "an optional parameter says so", "Filter", True, InStr(text, "- Filter (optional): ") > 0
    Check "rovo", "no connection parameter at all, nor what the workbook reads in Jira", "BaseUrl ProxyServer PIField MaxRetries", "0000", _
          InStr(text, "BaseUrl") & InStr(text, "ProxyServer") & InStr(text, "- PIField") & InStr(text, "- MaxRetries")
    Check "rovo", "the example of the form of the value follows", "Example:", True, InStr(text, " Example: ") > 0
    Check "rovo", "no Value of Config is in the prompt", "SECRET-HOST", 0, InStr(text, "SECRET-HOST")
    Check "rovo", "a sheet: its header row, then its rule", "Projects", True, InStr(text, "SHEET Projects_from_user - columns: Project | Role | Team. ") > 0
    Check "rovo", "the column Example of a sheet is not a column to fill", "PIs", True, InStr(text, "SHEET PIs_from_user - columns: PI | Start | End. ") > 0
    Check "rovo", "no sheet that the workbook reads in Jira (sprints, moves)", "Iterations Transitions", "00", _
          InStr(text, "SHEET Iterations_from_user") & InStr(text, "SHEET Transitions_from_user")
    Check "rovo", "the format asks for a CSV with its columns", "Sheet,Row,Column,Value,Note", True, InStr(text, "Sheet,Row,Column,Value,Note") > 0
    n = WriteRovoPrompt(ThisWorkbook)
    Set ws = ThisWorkbook.Worksheets(SheetName(SHEET_ROVO))
    Check "rovo", "the sheet RovoPrompt holds one line per row", "count", lines.Count, n
    Check "rovo", "the last line of the sheet is the last line of the prompt", "last", CStr(lines(lines.Count)), CStr(ws.Cells(n, 1).Value)
    Check "rovo", "nothing under the last line", "next", True, IsEmpty(ws.Cells(n + 1, 1).Value)
    firstText = CStr(ws.Cells(1, 1).Value)
    Check "rovo", "a second run rewrites the same sheet", "first line", firstText, CStr(lines(1))
    Check "rovo", "and writes the same number of lines", "count", n, WriteRovoPrompt(ThisWorkbook)
End Sub

Public Sub T_StepsSheet()
    ' The sheet Steps: a row per step with its status and its time; a stopped run leaves "error", a cancelled one "cancelled".
    Dim saved As String, ws As Object, i As Long
    saved = gLogPath
    gLogPath = Environ$("TEMP") & "\j2d_steps_test_" & Format$(Now, "hhnnss") & Int(Rnd() * 1000) & ".log"
    Application.DisplayAlerts = False
    If SheetExists(ThisWorkbook, "Steps") Then ThisWorkbook.Worksheets(SheetName("Steps")).Delete
    Application.DisplayAlerts = True
    LogRunStart "steps test"
    Set ws = ThisWorkbook.Worksheets(SheetName("Steps"))
    Check "steps", "the sheet is created with its header", "A1:D1", "Step|Status|Time elapsed (s)|Error code", ws.Range("A1").Value & "|" & ws.Range("B1").Value & "|" & ws.Range("C1").Value & "|" & ws.Range("D1").Value
    LogBegin "outer"
    LogBegin "inner"
    Check "steps", "nothing is written while the run goes on (the Dashboard shows its progress)", "A2", True, IsEmpty(ws.Range("A2").Value)
    LogEnd "inner: 5 items"
    LogEnd "outer"
    LogEnd "stray"
    LogBegin "stopped step"
    LogError "Step", vbObjectError + 608, "Fill these values of the Config sheet first"
    LogRunEnd "stopped"
    Check "steps", "at the end, the steps in the order they began; the end text says more", "inner", "outer|done|inner: 5 items|done", _
          ws.Range("A2").Value & "|" & ws.Range("B2").Value & "|" & ws.Range("A3").Value & "|" & ws.Range("B3").Value
    Check "steps", "the time is a number of seconds", "C2", True, IsNumeric(ws.Range("C2").Value) And Not IsEmpty(ws.Range("C2").Value)
    Check "steps", "an end without a begin adds nothing", "A5", True, IsEmpty(ws.Range("A5").Value)
    Check "steps", "a stopped run leaves its running step as error, with the support code of the error", "A4", "stopped step|error|J2D-CFG-002", _
          ws.Range("A4").Value & "|" & ws.Range("B4").Value & "|" & ws.Range("D4").Value
    Check "steps", "a step that ended well has no code", "D2", True, IsEmpty(ws.Range("D2").Value)
    LogRunStart "second run", "Errors"                                ' the level Errors keeps the sheet: it shows every step
    Check "steps", "a new run empties the sheet", "A2", True, IsEmpty(ws.Range("A2").Value)
    LogBegin "waiting"
    LogRunEnd "cancelled during a wait for Jira"
    Check "steps", "a cancelled run says cancelled", "A2", "waiting|cancelled", ws.Range("A2").Value & "|" & ws.Range("B2").Value
    LogRunStart "third run"
    LogBegin "last"
    LogRunEnd "ok"
    Check "steps", "a run that ends ok leaves nothing running", "A2", "last|done", ws.Range("A2").Value & "|" & ws.Range("B2").Value
    On Error Resume Next
    Kill gLogPath
    Err.Clear
    On Error GoTo 0
    gLogPath = saved
End Sub

Public Sub T_StatusParts()
    ' The status of the Dashboard: "Title | Name: value | ..." gives the headline (the cell dashStatus) and one row per field.
    Dim parts As Collection
    Set parts = StatusParts("Stopped | Code: J2D-CFG-002 | Message: Fill these values: FeatureTypes | Log: C:\x\run.log")
    Check "status", "the title, then one row per field", "count", 4, parts.Count
    Check "status", "the title has no name", "title", "|Stopped", parts(1)(0) & "|" & parts(1)(1)
    Check "status", "a field is cut at its first colon", "Message", "Message|Fill these values: FeatureTypes", parts(3)(0) & "|" & parts(3)(1)
    Check "status", "a path keeps its colon", "Log", "Log|C:\x\run.log", parts(4)(0) & "|" & parts(4)(1)
    Set parts = StatusParts("Reading Jira: page 3")
    Check "status", "the progress of a run is the headline alone", "progress", "1|Reading Jira: page 3", parts.Count & "|" & parts(1)(1)
    Set parts = StatusParts("Done | the six flow metrics")
    Check "status", "a part without a name is a row with no name", "no name", "|the six flow metrics", parts(2)(0) & "|" & parts(2)(1)
    Check "status", "an empty status", "empty", "1|", StatusParts("").Count & "|" & StatusParts("")(1)(1)
End Sub

Public Sub T_ErrorCodes()
    ' The support codes (modLog.ErrorCode): the same table as src/jira2dashboard/error_codes.py (a unit test compares the two).
    Check "errcode", "a project error maps to its code", "608", "J2D-CFG-002", ErrorCode(vbObjectError + 608)
    Check "errcode", "the cancel button is a code of its own", "610", "J2D-SYS-001", ErrorCode(vbObjectError + 610)
    Check "errcode", "an answer of Jira carries its HTTP status", "401", "J2D-JIRA-401", ErrorCode(vbObjectError + 602, "HTTP 401 after 1 attempt(s).  {x}")
    Check "errcode", "the BaseUrl check shares 602 and has its own code", "602", "J2D-CFG-005", ErrorCode(vbObjectError + 602, "BaseUrl must start with https://")
    Check "errcode", "no answer at all carries the WinHTTP error", "12007", "J2D-NET-12007", _
          ErrorCode(vbObjectError + 602, "HTTP 0 after 3 attempt(s). -2147012889 The server name or address could not be resolved")
    Check "errcode", "no answer and no known number", "HTTP 0", "J2D-NET-001", ErrorCode(vbObjectError + 602, "HTTP 0 after 3 attempt(s). cancelled")
    Check "errcode", "a WinHTTP HRESULT raised by itself", "&H80072EE7", "J2D-NET-12007", ErrorCode(&H80072EE7)
    Check "errcode", "a WinHTTP number without the HRESULT", "12029", "J2D-NET-12029", ErrorCode(12029)
    Check "errcode", "an error of VBA keeps its number", "13", "J2D-VBA-13", ErrorCode(13)
    Check "errcode", "a negative error of VBA reads in hexadecimal", "&H80004005", "J2D-VBA-80004005", ErrorCode(&H80004005)
    Check "errcode", "a project error that has no code yet", "999", "J2D-GEN-999", ErrorCode(vbObjectError + 999)
End Sub

Public Sub T_Log()
    ' The local log (modLog): one ASCII line per entry, no answer of Jira, a new file above 1 MB, never an error.
    Dim folder As String, file As String, saved As String, text As String, ok As Boolean
    saved = gLogPath
    folder = Environ$("TEMP") & "\j2d_log_test_" & Format$(Now, "hhnnss") & Int(Rnd() * 1000)
    file = folder & "\sub\run.log"
    gLogPath = file
    On Error Resume Next
    DeleteSetting "jira2dashboard", "LogTests"                       ' no flag left by an earlier test run
    On Error GoTo 0
    LogLine "TEST", "first" & vbCrLf & "second" & vbTab & ChrW(233)
    text = FileText(file)
    Check "log", "missing folders are created; one ASCII line per entry", "CR LF TAB e-acute", True, _
          text Like "####-##-## ##:##:##.### TEST first  second ?" & vbCrLf
    LogLine "TEST", String$(400, "x")
    text = FileText(file)
    Check "log", "a long text is cut", "400 x", True, InStr(text, " TEST " & String$(300, "x") & "..." & vbCrLf) > 0
    LogError "Step", vbObjectError + 602, "HTTP 500 after 1 try: {secret answer}"
    LogError "Step", vbObjectError + 605, "Sheet Iterations has no iteration"
    LogError "Step", 13, "Type mismatch"
    text = FileText(file)
    LogError "Step", vbObjectError + 513, "Unexpected character ""x"" at position 3"
    Check "log", "a JSON parse error keeps its number only (its text quotes the answer of Jira)", "513", True, _
          InStr(FileText(file), "ERROR J2D-DATA-001 | Step | vbObjectError+513 | (description not logged") > 0 And InStr(FileText(file), "Unexpected") = 0
    Check "log", "an error that quotes the answer of Jira keeps its number only", "602", True, _
          InStr(text, "ERROR J2D-JIRA-500 | Step | vbObjectError+602 | (description not logged") > 0 And InStr(text, "secret") = 0
    Check "log", "a project error reads vbObjectError+n, with its description", "605", True, _
          InStr(text, "ERROR J2D-GEN-605 | Step | vbObjectError+605 | Sheet Iterations has no iteration") > 0
    Check "log", "a VBA error keeps its number", "13", True, InStr(text, "ERROR J2D-VBA-13 | Step | 13 | Type mismatch") > 0
    WriteFileText file, String$(1000001, "y")                       ' too big: the next run starts a new file
    LogRunStart "test"
    Check "log", "a log above 1 MB becomes .1.log", "1000001 bytes", True, FileLen(folder & "\sub\run.1.log") >= 1000001
    Check "log", "the new file starts with the run line", "RUN test", True, _
          FileText(file) Like "####-##-## ##:##:##.### RUN test | Excel *" & vbCrLf
    LogRunEnd "ok"
    LogRunStart "level", "Errors"
    LogBegin "hidden step"
    LogError "Step", 13, "Type mismatch"
    LogRunEnd "ok"
    LogRunStart "steps again"
    LogBegin "shown step"
    text = FileText(file)
    Check "log", "level Errors: no BEGIN/END, but RUN and ERROR", "Errors", True, _
          InStr(text, "hidden step") = 0 And InStr(text, "RUN level |") > 0 And InStr(text, "RUN end ok") > 0
    Check "log", "a new run is back to the level Steps", "no level", True, InStr(text, "BEGIN shown step") > 0
    Check "log", "a run that ended leaves no warning", "end ok then start", 0, InStr(text, "did not end")
    LogRunStart "after a crash"                                     ' the run "steps again" never ended
    text = FileText(file)
    Check "log", "the next run says the previous one did not end", "no end", True, _
          InStr(text, "RUN after a crash |") > 0 And InStr(text, "RUN previous run did not end (started ") > 0
    LogRunEnd "ok"
    Check "log", "the folder logs next to the workbook", "C:\D", "C:\D\logs", LogFolderFrom("C:\D", "C:\L", "C:\T")
    Check "log", "a workbook with no folder: under LOCALAPPDATA", "empty", "C:\L\jira2dashboard\logs", LogFolderFrom("", "C:\L", "C:\T")
    Check "log", "a workbook opened from a web address: under LOCALAPPDATA", "https", "C:\L\jira2dashboard\logs", _
          LogFolderFrom("https://site/doc", "C:\L", "C:\T")
    Check "log", "TEMP when LOCALAPPDATA is not set", "empty", "C:\T\jira2dashboard\logs", LogFolderFrom("", "", "C:\T")
    gLogFallbackFile = folder & "\sub\fallback.log"                 ' a file that cannot be written: the line goes to the fallback
    gLogPath = folder & "\bad<name>|.log"
    LogLine "TEST", "to the fallback"
    Check "log", "a folder that cannot be written: the line goes to the fallback file", "fallback", True, _
          InStr(FileText(gLogFallbackFile), "TEST to the fallback") > 0
    Check "log", "and the log file is now the fallback", "LogFile", gLogFallbackFile, LogFile()
    gLogFallbackFile = ""
    gLogPath = file
    LogRunStart "fallback over"                                       ' a new run tries the folder again
    Check "log", "a new run goes back to the usual file", "LogFile", file, LogFile()
    LogRunEnd "ok"
    gLogPath = folder & "\bad<name>|.log"                            ' a name Windows refuses
    On Error Resume Next
    LogLine "TEST", "dropped"
    ok = (Err.Number = 0)
    On Error GoTo 0
    Check "log", "a line that cannot be written is dropped, without an error", "bad file name", True, ok
    gLogPath = saved
    On Error Resume Next
    Kill folder & "\sub\*.log"
    RmDir folder & "\sub"
    RmDir folder
    Err.Clear
    On Error GoTo 0
End Sub

Public Sub T_LogInputsAndTimes()
    ' What the log gives to debug a run: the time of each step and of the run, the code of the error that stopped it, and
    ' the inputs (PARAM lines) with no value of the client.
    Dim saved As String, file As String, text As String, other As Object
    saved = gLogPath
    file = Environ$("TEMP") & "\j2d_log_inputs_" & Format$(Now, "hhnnss") & Int(Rnd() * 1000) & ".log"
    gLogPath = file
    Check "logtime", "a time reads in seconds with a point", "1.5", "1.500 s", LogDuration(1.5)
    Check "logtime", "a time is rounded to the millisecond", "61.2346", "61.235 s", LogDuration(61.2346)
    Check "logtime", "no negative time", "-2", "0.000 s", LogDuration(-2)
    Check "logparam", "a setting of the tool is shown", "FlowLoadMode B", "B", ParameterShown("FlowLoadMode", " B ")
    Check "logparam", "a value of the client is not shown", "ApiToken", "(set)", ParameterShown("ApiToken", "secret-token")
    Check "logparam", "nor the address of Jira", "BaseUrl", "(set)", ParameterShown("BaseUrl", "https://x.atlassian.net")
    Check "logparam", "an empty value says so", "Email", "(empty)", ParameterShown("Email", "  ")
    Check "logparam", "a parameter not known reads as a value of the client", "NewOne", "(set)", ParameterShown("NewOne", "v")
    Check "logparam", "AsOf is shown", "AsOf", "2026-11-03 18:00", ParameterShown("AsOf", "2026-11-03 18:00")
    Check "logparam", "a field id is shown", "StoryPointsField", "customfield_10016", ParameterShown("StoryPointsField", "customfield_10016")
    Check "logparam", "nor the filter", "Filter", "(set)", ParameterShown("Filter", "updated >= -90d")
    Check "logparam", "a date typed in a cell reads in ISO", "AsOf date", "2026-11-03 18:00:00", _
          ParameterText(DateSerial(2026, 11, 3) + TimeSerial(18, 0, 0))
    Check "logparam", "a number reads as typed", "60", "60", ParameterText(60)
    LogRunStart "inputs and times"
    LogParameters ThisWorkbook
    LogBegin "outer"
    LogBegin "inner"
    LogEnd "inner: 3 items"
    LogEnd "outer"
    LogEnd "stray"
    Check "logrun", "a run is in progress between its start and its end", "start", True, LogRunActive()
    LogRunEnd "ok"
    Check "logrun", "and no longer after its end", "end", False, LogRunActive()
    text = FileText(file)
    Check "logparam", "a PARAM line per row of Config, the setting in clear", "FlowLoadMode", True, InStr(text, " PARAM FlowLoadMode = A" & vbCrLf) + InStr(text, " PARAM FlowLoadMode = B" & vbCrLf) > 0
    Check "logparam", "and the value of the client as set or empty", "BaseUrl", True, InStr(text, " PARAM BaseUrl = (") > 0 And InStr(text, " PARAM BaseUrl = h") = 0
    Check "logparam", "the size of each parameter sheet", "Projects", True, InStr(text, " PARAM sheet Projects = ") > 0 And InStr(text, " rows" & vbCrLf) > 0
    Check "logtime", "an END line ends with the time of its step", "inner", True, InStr(text, " END inner: 3 items | ") > 0
    Check "logtime", "nested steps each get their own time", "outer", True, InStr(text, " END outer | ") > 0
    Check "logtime", "an END without a BEGIN has no time", "stray", True, InStr(text, " END stray" & vbCrLf) > 0
    Check "logtime", "the line RUN end gives the time of the run", "ok", True, InStr(text, " RUN end ok | ") > 0 And InStr(text, " s" & vbCrLf) > 0
    Check "logversion", "the RUN line names the version of the workbook", "RUN", True, InStr(text, " | jira2dashboard " & WorkbookVersion(ThisWorkbook) & vbCrLf) > 0
    LogRunStart "stopped run"
    LogError "Step", vbObjectError + 608, "Fill these values of the Config sheet first"
    LogRunEnd "stopped"
    LogRunEnd "stopped"                                              ' a second end (no run on): no time
    text = FileText(file)
    Check "logtime", "a stopped run gives the code of its error on its last line", "608", True, InStr(text, " RUN end stopped J2D-CFG-002 | ") > 0
    Check "logtime", "an end without a run start has no time", "twice", True, InStr(text, " RUN end stopped J2D-CFG-002" & vbCrLf) > 0
    Set other = Application.Workbooks.Add                           ' a workbook without any parameter sheet
    LogParameters other
    Check "logversion", "a workbook without the name J2D_Version", "none", "unknown", WorkbookVersion(other)
    other.Names.Add "J2D_Version", "=""v1.0.0 commit 1a2b3c4, built 2026-11-11 19:40""", False
    Check "logversion", "the version written at build", "J2D_Version", "v1.0.0 commit 1a2b3c4, built 2026-11-11 19:40", WorkbookVersion(other)
    other.Worksheets(1).Name = SheetName("PIs")                                 ' a header and one row: "1 row", not "1 rows"
    other.Worksheets(1).Range("A1:A2").Value = Application.Transpose(Array("PI", "PI#01-2026"))
    LogParameters other
    other.Close False
    text = FileText(file)
    Check "logparam", "a workbook without Config says so", "no Config", True, InStr(text, " PARAM no sheet Config" & vbCrLf) > 0
    Check "logparam", "and names the missing sheets", "Workflows", True, InStr(text, " PARAM sheet Workflows = missing" & vbCrLf) > 0
    Check "logparam", "one row reads in the singular", "PIs", True, InStr(text, " PARAM sheet PIs = 1 row" & vbCrLf) > 0
    On Error Resume Next
    Kill file
    Err.Clear
    On Error GoTo 0
    gLogPath = saved
End Sub

Private Function FileText(ByVal file As String) As String
    Dim f As Integer, text As String
    f = FreeFile
    Open file For Binary Access Read As #f
    text = Space$(LOF(f))
    Get #f, , text
    Close #f
    FileText = text
End Function

Private Sub WriteFileText(ByVal file As String, ByVal text As String)
    Dim f As Integer
    f = FreeFile
    Open file For Output As #f
    Print #f, text;
    Close #f
End Sub

Public Sub T_ParseMoment()
    Check "asof", "a date", "2026-11-03", CDbl(DateSerial(2026, 11, 3)), CDbl(ParseMoment("2026-11-03"))
    Check "asof", "date and time", "2026-11-03 18:00", CDbl(DateSerial(2026, 11, 3)) + 0.75, CDbl(ParseMoment("2026-11-03 18:00"))
    Check "asof", "ISO with seconds", "2026-11-03T18:30:15", CDbl(DateSerial(2026, 11, 3) + TimeSerial(18, 30, 15)), CDbl(ParseMoment(" 2026-11-03T18:30:15 "))
    Check "asof", "words are not a date", "tomorrow", True, IsEmpty(ParseMoment("tomorrow"))
    Check "asof", "too short", "2026-11", True, IsEmpty(ParseMoment("2026-11"))
    Check "asof", "a day that does not exist", "2026-13-45", True, IsEmpty(ParseMoment("2026-13-45"))
    Check "asof", "letters in the numbers", "20x6-11-03", True, IsEmpty(ParseMoment("20x6-11-03"))
    Check "asof", "empty", "", True, IsEmpty(ParseMoment(""))
    Check "asof", "an hour on one digit is read, as the sheet reads it (it was dropped)", "2026-11-03 8:00", CDbl(DateSerial(2026, 11, 3)) + 8 / 24, CDbl(ParseMoment("2026-11-03 8:00"))
    Check "asof", "a UTC suffix after the seconds is ignored, as the sheet ignores it", "...:00Z", CDbl(DateSerial(2026, 11, 3) + TimeSerial(18, 0, 0)), CDbl(ParseMoment("2026-11-03T18:00:00Z"))
    Check "asof", "an hour without minutes is no time", "2026-11-03 18", True, IsEmpty(ParseMoment("2026-11-03 18"))
    Check "asof", "a time that does not exist", "25:00", True, IsEmpty(ParseMoment("2026-11-03 25:00")) And IsEmpty(ParseMoment("2026-11-03 18:60"))
    Check "asof", "an offset is refused, not dropped (1 or 2 hours wrong)", "+01:00", True, IsEmpty(ParseMoment("2026-11-03T18:00:00+01:00")) And _
          IsEmpty(ParseMoment("2026-11-03T18:00:00.000Z"))
    Dim ws As Object, raised As Boolean
    Set ws = ThisWorkbook.Worksheets("TestScratch")
    ConfigRows ws, Array("AsOf", "2026-11-03 18:00")
    Check "asof", "AsOf read from the Config sheet", "2026-11-03 18:00", CDbl(DateSerial(2026, 11, 3)) + 0.75, CDbl(AsOfFromConfig(ws))
    ConfigRows ws, Array("AsOf", "")
    On Error Resume Next
    Check "asof", "an empty AsOf raises", "", 0, CDbl(AsOfFromConfig(ws))
    raised = (Err.Number <> 0): Err.Clear
    On Error GoTo 0
    Check "asof", "the error asks to type AsOf", "empty AsOf", True, raised
    ConfigRows ws, Array("AsOf", "x")
    ws.Cells(2, 2).Value = DateSerial(2026, 11, 3) + TimeSerial(18, 0, 0)        ' typed as a date: the cell holds a date, not text
    Check "asof", "AsOf typed as a date in the cell", "2026-11-03 18:00 as a date", CDbl(DateSerial(2026, 11, 3)) + 0.75, CDbl(AsOfFromConfig(ws))
    ws.Cells(2, 2).Value = 5                                                     ' a plain number is not a date
    raised = False
    On Error Resume Next
    Call AsOfFromConfig(ws)
    raised = (Err.Number = vbObjectError + 604): Err.Clear
    On Error GoTo 0
    Check "asof", "a number such as 5 is refused, not read as 1900-01-05", "5", True, raised
    ws.Cells(2, 2).NumberFormat = "General"                                      ' digits typed in a plain cell (the cell above was a date)
    ws.Cells(2, 2).Value = 20261120                                              ' yyyymmdd typed as a number: the year 57 000 (it raised Overflow)
    raised = False
    On Error Resume Next
    Call AsOfFromConfig(ws)
    raised = (Err.Number = vbObjectError + 604): Err.Clear
    On Error GoTo 0
    Check "asof", "20261120 typed as a number is refused with the code of AsOf", "20261120", True, raised
    ConfigRows ws, Array("AsOf", "2126-11-03 18:00")
    raised = False
    On Error Resume Next
    Call AsOfFromConfig(ws)
    raised = (Err.Number = vbObjectError + 604): Err.Clear
    On Error GoTo 0
    Check "asof", "a year from 2100 on is refused", "2126", True, raised
End Sub

Public Sub T_ExtractFlowData()
    ' Raw sheets written from fake Jira answers: one Program feature (Feature workflow), one team story (Team workflow).
    Dim cfg As Object, wsI As Object, wsV As Object, jql As String, pret As String, ts As String, body As String
    Set cfg = TestConfig()
    Set wsI = ThisWorkbook.Worksheets("TestScratch")
    Set wsV = ThisWorkbook.Worksheets("TestScratch2")
    pret = "Pr" & ChrW(234) & "t"
    ts = "T00:00:00.000+0000"
    jql = "project in (TTT,AAA)"
    FakeReset
    FakeAdd "GET", JiraSearchUrl(cfg("BaseUrl"), jql, "", 100, ExtractFields("customfield_10200"), ""), 200, J( _
        "{'issues':[{'id':'10','key':'TTT-1','fields':{'project':{'key':'TTT'},'issuetype':{'name':'Feature'},'created':'2026-10-01" & ts & "'," & _
        "'status':{'name':'" & pret & "','statusCategory':{'key':'indeterminate'}},'parent':null,'resolutiondate':null,'customfield_10200':'Feature'}}," & _
        "{'id':'11','key':'AAA-1','fields':{'project':{'key':'AAA'},'issuetype':{'name':'Story'},'created':'2026-10-01" & ts & "'," & _
        "'status':{'name':'En analyse','statusCategory':{'key':'done'}},'parent':{'key':'TTT-1'},'resolutiondate':'2026-10-04" & ts & "'," & _
        "'customfield_10200':{'value':'Enabler'}}}],'isLast':true}")
    body = "{'issueChangeLogs':[{'issueId':'10','changeHistories':[" & _
        "{'id':'1','created':'2026-10-02" & ts & "','items':[{'fieldId':'status','field':'status','fieldtype':'jira','from':'1','fromString':'Nouveau','to':'2','toString':'En analyse'}]}," & _
        "{'id':'2','created':'2026-10-03" & ts & "','items':[{'fieldId':'status','field':'status','fieldtype':'jira','from':'2','fromString':'En analyse','to':'3','toString':'" & pret & "'}]}]}," & _
        "{'issueId':'11','changeHistories':[" & _
        "{'id':'3','created':'2026-10-02" & ts & "','items':[{'fieldId':'status','field':'status','fieldtype':'jira','from':'1','fromString':'Nouveau','to':'2','toString':'En analyse'}]}]}]}"
    FakeAdd "POST", cfg("BaseUrl") & "/rest/api/3/changelog/bulkfetch", 200, J(body)
    ExtractFlowDataTo cfg, jql, "customfield_10200", DateSerial(2026, 10, 5), ThisWorkbook.Worksheets(SheetName("Projects")), _
                      ThisWorkbook.Worksheets(SheetName("Workflows")), wsI, wsV
    Check "extract", "Issues header", "A1:V1", "Key|Id|Project|Type|WorkType|CreatedUtc|Status|Parent|StatusCategory|ResolvedUtc|" & _
          "FixVersions|StoryPoints|PI|BV|TeamValue|SprintState|SprintStart|SprintEnd|SprintCount|TargetStart|TargetEnd|KeyChanges", _
          RowText(wsI, 1, 22)
    Check "extract", "no parent and no resolution: empty cells", "TTT-1", "Empty|indeterminate|Empty", _
          IIf(IsEmpty(wsI.Range("H2").Value), "Empty", "set") & "|" & wsI.Range("I2").Value & "|" & IIf(IsEmpty(wsI.Range("J2").Value), "Empty", "set")
    Check "extract", "parent, status category key and resolution date", "AAA-1", "TTT-1|done|" & CDbl(DateSerial(2026, 10, 4)), _
          wsI.Range("H3").Value & "|" & wsI.Range("I3").Value & "|" & wsI.Range("J3").Value2
    Check "extract", "one row per issue", "TTT-1, AAA-1", "TTT-1|AAA-1", wsI.Range("A2").Value & "|" & wsI.Range("A3").Value
    Check "extract", "project, type and status", "TTT-1", "TTT|Feature|" & pret, wsI.Range("C2").Value & "|" & wsI.Range("D2").Value & "|" & wsI.Range("G2").Value
    Check "extract", "work type as text", "customfield_10200", "Feature", wsI.Range("E2").Value
    Check "extract", "work type of a select field (object)", "{value: Enabler}", "Enabler", wsI.Range("E3").Value
    Check "extract", "created is a date serial in UTC", "2026-10-01", CDbl(DateSerial(2026, 10, 1)), wsI.Range("F2").Value2
    Check "extract", "5 intervals: 3 for the feature, 2 for the story", "rows", 5, wsV.Cells(wsV.Rows.Count, 1).End(-4162).Row - 1
    Check "extract", "the Program follows the Feature workflow", "TTT-1", "Feature", wsV.Range("B2").Value
    Check "extract", "the others follow the Team workflow", "AAA-1", "Team", wsV.Range("B5").Value
    Check "extract", "first interval: Nouveau is Backlog", "TTT-1 seq 1", "Nouveau|Backlog|1", wsV.Range("D2").Value & "|" & wsV.Range("E2").Value & "|" & wsV.Range("C2").Value
    Check "extract", "category comes from the Workflows sheet", "TTT-1 Pret", pret & "|Wait", wsV.Range("D4").Value & "|" & wsV.Range("E4").Value
    Check "extract", "entered is a UTC serial", "2026-10-02", CDbl(DateSerial(2026, 10, 2)), wsV.Range("F3").Value2
    Check "extract", "the change that ended the interval", "TTT-1 seq 1", "1", CStr(wsV.Range("J2").Value)
    Check "extract", "the last interval is open and has no left date", "AAA-1", "True|Empty", IIf(wsV.Range("I6").Value, "True", "False") & "|" & IIf(IsEmpty(wsV.Range("G6").Value), "Empty", "set")
    Check "extract", "an open interval runs to the as-of date", "10-02 to 10-05", 3, wsV.Range("H6").Value
    Check "extract", "a closed interval has its duration", "TTT-1 seq 1", 1, wsV.Range("H2").Value
End Sub

Private Function RowText(ByVal ws As Object, ByVal row As Long, ByVal cols As Long) As String
    ' The cells of one row joined with |, empty cells as nothing.
    Dim c As Long, s As String
    For c = 1 To cols
        s = s & IIf(c > 1, "|", "") & CStr(ws.Cells(row, c).Value & "")
    Next c
    RowText = s
End Function

Private Function CheckFields(ByVal points As String, ByVal pi As String, ByVal bv As String, ByVal team As String, _
                             ByVal sprint As String, Optional ByVal startDate As String = "", Optional ByVal endDate As String = "", _
                             Optional ByVal linkType As String = "") As Object
    ' The Config settings of the data checks, as FieldSettings gives them.
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("StoryPointsField") = points: d("PIField") = pi: d("BusinessValueField") = bv: d("TeamField") = team
    d("SprintField") = sprint: d("TargetStartField") = startDate: d("TargetEndField") = endDate: d("BlocksLinkType") = linkType
    Set CheckFields = d
End Function

Private Function LastRow(ByVal ws As Object) As Long
    LastRow = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
End Function

Public Sub T_ExtractJob()
    ' The extraction as a job: one step per page of the search, per chunk of changelogs, per batch of issue rows, per workflow
    ' of intervals, and one that writes. The bar of the Dashboard moves with the fraction, and a job that is dropped before
    ' its last step writes nothing.
    Dim cfg As Object, wsI As Object, wsV As Object, jql As String, ts As String, fields As String, wsC As Object
    Dim steps As Long, last As Double, increasing As Boolean, f As Double, first As String, page1 As String, page2 As String
    Set cfg = TestConfig()
    Set wsI = ThisWorkbook.Worksheets("TestScratch")
    Set wsV = ThisWorkbook.Worksheets("TestScratch2")
    ts = "T00:00:00.000+0000"
    jql = "project in (TTT,AAA)"
    fields = ExtractFields("")
    page1 = "{'issues':[{'id':'10','key':'TTT-1','fields':{'project':{'key':'TTT'},'issuetype':{'name':'Feature'},'created':'2026-10-01" & ts & "'," & _
            "'status':{'name':'Nouveau','statusCategory':{'key':'new'}}}}],'isLast':false,'nextPageToken':'p2'}"
    page2 = "{'issues':[{'id':'11','key':'AAA-1','fields':{'project':{'key':'AAA'},'issuetype':{'name':'Story'},'created':'2026-10-01" & ts & "'," & _
            "'status':{'name':'Nouveau','statusCategory':{'key':'new'}}}}],'isLast':true}"
    FakeReset
    FakeAdd "GET", JiraSearchUrl(cfg("BaseUrl"), jql, "", 100, fields, ""), 200, J(page1)
    FakeAdd "GET", JiraSearchUrl(cfg("BaseUrl"), jql, "", 100, fields, ""), 200, J(page1)       ' for the second job
    FakeAdd "GET", JiraSearchUrl(cfg("BaseUrl"), jql, "p2", 100, fields, ""), 200, J(page2)
    FakeAdd "POST", cfg("BaseUrl") & "/rest/api/3/changelog/bulkfetch", 200, J("{'issueChangeLogs':[]}")
    wsI.Cells.ClearContents
    wsI.Range("A2").Value = "KEEP"
    ExtractJobStart cfg, jql, "", DateSerial(2026, 10, 5), ThisWorkbook.Worksheets(SheetName("Projects")), ThisWorkbook.Worksheets(SheetName("Workflows")), wsI, wsV
    ExtractJobStep
    Check "extractjob", "after one step the job is not finished and says what it reads", "page 1", "True|Reading Jira: 1 issues (page 1)", _
          IIf(ExtractJobFraction() > 0 And ExtractJobFraction() < 0.3, "True", "False") & "|" & ExtractJobLabel()
    Check "extractjob", "a job dropped before its last step writes nothing", "A2", "KEEP", wsI.Range("A2").Value
    ExtractJobAbort
    Check "extractjob", "an aborted job is finished", "fraction", 1, ExtractJobFraction()
    ExtractJobStart cfg, jql, "", DateSerial(2026, 10, 5), ThisWorkbook.Worksheets(SheetName("Projects")), ThisWorkbook.Worksheets(SheetName("Workflows")), wsI, wsV
    increasing = True
    Do
        steps = steps + 1
        If ExtractJobStep() Then Exit Do
        f = ExtractJobFraction()
        If f < last Then increasing = False
        last = f
        If steps = 3 Then first = ExtractJobLabel()
        If steps > 20 Then Exit Do
    Loop
    Check "extractjob", "two pages, a chunk, a batch of rows, two workflows and the write make seven steps", "steps", 7, steps
    Check "extractjob", "the fraction never goes back", "fractions", True, increasing
    Check "extractjob", "before the rows the label says what comes", "label", "Preparing the issues: 0 / 2", first
    Check "extractjob", "the raw sheets hold both pages", "TTT-1, AAA-1", "TTT-1|AAA-1", wsI.Range("A2").Value & "|" & wsI.Range("A3").Value
    Check "extractjob", "a finished job reads 100 %", "fraction", 1, ExtractJobFraction()
    ' A batch of one issue: one more step, the same sheets.
    FakeAdd "GET", JiraSearchUrl(cfg("BaseUrl"), jql, "", 100, fields, ""), 200, J(page1)
    FakeAdd "GET", JiraSearchUrl(cfg("BaseUrl"), jql, "p2", 100, fields, ""), 200, J(page2)
    FakeAdd "POST", cfg("BaseUrl") & "/rest/api/3/changelog/bulkfetch", 200, J("{'issueChangeLogs':[]}")
    wsI.Cells.ClearContents
    gRowBatch = 1
    Set wsC = ThisWorkbook.Worksheets("TestScratch3")
    wsC.Cells.Clear
    ExtractJobStart cfg, jql, "", DateSerial(2026, 10, 5), ThisWorkbook.Worksheets(SheetName("Projects")), ThisWorkbook.Worksheets(SheetName("Workflows")), wsI, wsV, wsChangelog:=wsC
    steps = 0
    Do
        steps = steps + 1
        If ExtractJobStep() Then Exit Do
        If steps > 20 Then Exit Do
    Loop
    gRowBatch = 0
    Check "extractjob", "a batch of one issue adds one step", "steps", 8, steps
    Check "extractjob", "the batches give the same rows", "TTT-1, AAA-1", "TTT-1|AAA-1", wsI.Range("A2").Value & "|" & wsI.Range("A3").Value
    Check "extractjob", "the intervals are written once per issue", "2 rows", 3, wsV.Cells(wsV.Rows.Count, 1).End(-4162).Row
    Check "extractjob", "the sheet Changelog gets its header even when Jira sent no item", "Key|ToString", "Key|ToString", wsC.Range("A1").Value & "|" & wsC.Range("L1").Value
End Sub

Public Sub T_ExtractCheckFields()
    ' The fields that the data checks read: team (object or text), sprints, points, PI, dates, fix versions, blocking links
    ' and the moves between projects. One Program feature with almost none of them, one team story with all of them.
    Dim cfg As Object, wsI As Object, wsV As Object, wsL As Object, st As Object, jql As String, ts As String
    Dim wsC As Object, adminRow As Long, oldAdmin As Variant, r As Long, saved As String, file As String, text As String
    Set cfg = TestConfig()
    Set wsI = ThisWorkbook.Worksheets("TestScratch")
    Set wsV = ThisWorkbook.Worksheets("TestScratch2")
    Set wsL = ThisWorkbook.Worksheets("TestScratch3")
    Set st = CheckFields("customfield_1", "customfield_2", "customfield_3", "customfield_4", "customfield_5", "customfield_6", _
                         "customfield_7", "Blocks")
    ts = "T00:00:00.000+0000"
    jql = "project in (TTT,AAA)"
    FakeReset
    FakeAdd "GET", JiraSearchUrl(cfg("BaseUrl"), jql, "", 100, ExtractFields("", st), ""), 200, J( _
        "{'issues':[{'id':'10','key':'TTT-1','fields':{'project':{'key':'TTT'},'issuetype':{'name':'Feature'},'created':'2026-10-01" & ts & "'," & _
        "'status':{'name':'Nouveau','statusCategory':{'key':'new'}},'fixVersions':[],'issuelinks':[],'customfield_1':null," & _
        "'customfield_3':40,'customfield_4':'ART','customfield_5':null}}," & _
        "{'id':'11','key':'AAA-1','fields':{'project':{'key':'AAA'},'issuetype':{'name':'Story'},'created':'2026-10-01" & ts & "'," & _
        "'status':{'name':'En analyse','statusCategory':{'key':'indeterminate'}},'fixVersions':[{'name':'R1'},{'name':'R2'}]," & _
        "'customfield_1':5,'customfield_2':'PI#01-2026','customfield_4':{'id':'t1','name':'Alpha'}," & _
        "'customfield_5':[{'id':1,'state':'closed','startDate':'2026-09-01" & ts & "','endDate':'2026-09-15" & ts & "'}," & _
        "{'id':2,'state':'active','startDate':'2026-10-01" & ts & "','endDate':'2026-10-15" & ts & "'}," & _
        "{'id':3,'state':'future','startDate':'2026-10-15" & ts & "','endDate':'2026-10-29" & ts & "'}]," & _
        "'customfield_6':'2026-10-01','customfield_7':'2026-12-31'," & _
        "'issuelinks':[{'type':{'name':'Blocks'},'outwardIssue':{'key':'AAA-2'}},{'type':{'name':'Blocks'},'inwardIssue':{'key':'AAA-9'}}," & _
        "{'type':{'name':'Relates'},'outwardIssue':{'key':'AAA-3'}}]}}],'isLast':true}")
    FakeAdd "POST", cfg("BaseUrl") & "/rest/api/3/changelog/bulkfetch", 200, J( _
        "{'issueChangeLogs':[{'issueId':'11','changeHistories':[" & _
        "{'id':'1','created':'2026-10-02" & ts & "','items':[{'fieldId':'status','field':'status','fieldtype':'jira','from':'1','fromString':'Nouveau','to':'2','toString':'En analyse'}]}," & _
        "{'id':'2','created':'2026-10-03" & ts & "','items':[{'fieldId':'issuekey','field':'Key','fieldtype':'jira','from':null,'fromString':'OLD-1','to':null,'toString':'AAA-1'}]}]}]}")
    ExtractFlowDataTo cfg, jql, "", DateSerial(2026, 10, 5), ThisWorkbook.Worksheets(SheetName("Projects")), ThisWorkbook.Worksheets(SheetName("Workflows")), _
                      wsI, wsV, False, st, wsL
    Check "extract", "the feature has only what Jira gave: BV 40, team text ART, no sprint, no change of key", "TTT-1", _
          "TTT-1||40|ART||0|0", wsI.Range("A2").Value & "|" & wsI.Range("K2").Value & "|" & wsI.Range("N2").Value & "|" & _
          wsI.Range("O2").Value & "|" & wsI.Range("P2").Value & "|" & wsI.Range("S2").Value & "|" & wsI.Range("V2").Value
    Check "extract", "team as an object, points, PI", "AAA-1", "Alpha|5|PI#01-2026", wsI.Range("O3").Value & "|" & wsI.Range("L3").Value & "|" & wsI.Range("M3").Value
    Check "extract", "fix versions are joined", "R1, R2", "R1, R2", wsI.Range("K3").Value
    Check "extract", "the active sprint counts, over a closed and a future one", "AAA-1", "active|3", wsI.Range("P3").Value & "|" & wsI.Range("S3").Value
    Check "extract", "dates of that sprint", "2026-10-01..10-15", CDbl(DateSerial(2026, 10, 1)) & "|" & CDbl(DateSerial(2026, 10, 15)), _
          wsI.Range("Q3").Value2 & "|" & wsI.Range("R3").Value2
    Check "extract", "target dates (date only)", "2026-10-01 and 12-31", CDbl(DateSerial(2026, 10, 1)) & "|" & CDbl(DateSerial(2026, 12, 31)), _
          wsI.Range("T3").Value2 & "|" & wsI.Range("U3").Value2
    Check "extract", "a change of key is counted", "AAA-1", 1, wsI.Range("V3").Value
    Check "extract", "only the outward blocks link is a row", "AAA-1 blocks AAA-2", "AAA-1|AAA-2|1", _
          wsL.Range("A2").Value & "|" & wsL.Range("B2").Value & "|" & (LastRow(wsL) - 1)
    ' no link type configured: the links are not read; no Sprint field: no sprint
    FakeReset
    Set st = CheckFields("", "", "", "", "")
    FakeAdd "GET", JiraSearchUrl(cfg("BaseUrl"), jql, "", 100, ExtractFields("", st), ""), 200, J( _
        "{'issues':[{'id':'11','key':'AAA-1','fields':{'project':{'key':'AAA'},'issuetype':{'name':'Story'},'created':'2026-10-01" & ts & "'," & _
        "'status':{'name':'En analyse'},'issuelinks':[{'type':{'name':'Blocks'},'outwardIssue':{'key':'AAA-2'}}]}}],'isLast':true}")
    FakeAdd "POST", cfg("BaseUrl") & "/rest/api/3/changelog/bulkfetch", 200, J("{'issueChangeLogs':[]}")
    ' and IsUserProjectAdmin = False: the workflows of Jira are not asked (the log says it), the rules are counted in the log
    Set wsC = ThisWorkbook.Worksheets(SheetName("Config"))
    For r = 2 To LastRow(wsC)
        If CStr(wsC.Cells(r, 1).Value) = "IsUserProjectAdmin" Then adminRow = r
    Next r
    If adminRow > 0 Then oldAdmin = wsC.Cells(adminRow, 2).Value: wsC.Cells(adminRow, 2).Value = "False"
    saved = gLogPath
    file = Environ$("TEMP") & "\j2d_not_admin_" & Format$(Now, "hhnnss") & Int(Rnd() * 1000) & ".log"
    gLogPath = file
    ExtractFlowDataTo cfg, jql, "", DateSerial(2026, 10, 5), ThisWorkbook.Worksheets(SheetName("Projects")), ThisWorkbook.Worksheets(SheetName("Workflows")), _
                      wsI, wsV, False, st, wsL
    gLogPath = saved
    If adminRow > 0 Then wsC.Cells(adminRow, 2).Value = oldAdmin
    text = FileText(file)
    On Error Resume Next
    Kill file
    Err.Clear
    On Error GoTo 0
    Check "extract", "IsUserProjectAdmin False: the workflows of Jira are not asked, the log says why", "False", True, _
          InStr(text, " INFO workflows of Jira not read: IsUserProjectAdmin is False") > 0 And InStr(text, "J2D-JIRA-030") = 0
    Check "extract", "the log counts the rules of each workflow kind", "status rules", True, InStr(text, " INFO status rules | Feature: start ") > 0
    Check "extract", "no link type: no link row", "AAA-1", 0, LastRow(wsL) - 1
    Check "extract", "no sprint field: no sprint", "AAA-1", "|0", wsI.Range("P2").Value & "|" & wsI.Range("S2").Value
End Sub

Private Sub FlowRaw(ByVal withStory As Boolean, Optional ByVal featurePis As Variant, Optional ByVal storyPi As Variant)
    ' Raw sheets as the extraction writes them: a Program feature, and (optionally) a team story with two intervals.
    ' featurePis: one more feature (TTT-2, TTT-3...) per PI field value given; storyPi: the PI field of the story.
    Dim issues As New Collection, ivs As New Collection, t0 As Date, i As Long
    t0 = DateSerial(2026, 10, 1)
    If IsMissing(storyPi) Then storyPi = Empty
    issues.Add Array("TTT-1", "10", "TTT", "Feature", "Feature", t0, "Nouveau", Empty, "new", Empty, Empty, Empty, Empty, Empty, _
                     Empty, Empty, Empty, Empty, 0, Empty, Empty, 0)
    ivs.Add Array("TTT-1", "Feature", 1, "Nouveau", "Backlog", t0, Empty, 4, True, "1", "")
    If Not IsMissing(featurePis) Then
        For i = 0 To UBound(featurePis)
            issues.Add Array("TTT-" & (i + 2), CStr(20 + i), "TTT", "Feature", "Feature", t0, "Nouveau", Empty, "new", Empty, Empty, _
                             Empty, featurePis(i), Empty, IIf(i Mod 2 = 0, "team-b", "team-a"), Empty, Empty, Empty, 0, Empty, Empty, 0)
            ivs.Add Array("TTT-" & (i + 2), "Feature", 1, "Nouveau", "Backlog", t0, Empty, 4, True, CStr(20 + i), "")
        Next i
    End If
    If withStory Then
        issues.Add Array("AAA-1", "11", "AAA", "Story", "Enabler", t0, "En analyse", "TTT-1", "indeterminate", Empty, Empty, Empty, _
                         storyPi, Empty, Empty, Empty, Empty, Empty, 0, Empty, Empty, 0)
        ivs.Add Array("AAA-1", "Team", 1, "Nouveau", "Backlog", t0, t0 + 1, 1, False, "1", "")
        ivs.Add Array("AAA-1", "Team", 2, "En analyse", "Active", t0 + 1, Empty, 3, True, "2", "")
    End If
    WriteTable ThisWorkbook.Worksheets(SheetName("Issues")), Array("Key", "Id", "Project", "Type", "WorkType", "CreatedUtc", "Status", "Parent", _
        "StatusCategory", "ResolvedUtc", "FixVersions", "StoryPoints", "PI", "BV", "TeamValue", "SprintState", "SprintStart", _
        "SprintEnd", "SprintCount", "TargetStart", "TargetEnd", "KeyChanges"), issues
    WriteTable ThisWorkbook.Worksheets(SheetName("IssueLinks")), Array("Blocker", "Blocked"), New Collection
    WriteTable ThisWorkbook.Worksheets(SheetName("StatusIntervals")), Array("Key", "Workflow", "Seq", "Status", "Category", "EnteredUtc", _
        "LeftUtc", "DurationDays", "IsOpen", "ChangeId", "Check"), ivs
    Dim info As Object, issueRow As Variant                              ' the PI field of each issue from its creation (no change)
    Set info = CreateObject("Scripting.Dictionary")
    For Each issueRow In issues
        info(issueRow(0)) = Array(issueRow(5), issueRow(12))
    Next issueRow
    WriteTable ThisWorkbook.Worksheets(SheetName("PIIntervals")), PiIntervalsHeader(), PiIntervalRows(info, New Collection)
    ' the rules the extraction writes: the exceptions of the sheet Workflows (the delivered parameters) and nothing from Jira
    WriteTable ThisWorkbook.Worksheets(SheetName("StatusRules")), Array("Workflow", "Status", "Category", "IsStart", "IsEnd", "IsFree", "Source"), _
        EffectiveRules(Array("Feature", "Team"), Dict(), Dict(), Dict(), Dict(), WorkflowExceptions(ThisWorkbook.Worksheets(SheetName("Workflows"))))
End Sub

Private Sub FlowIterations(ByVal start1 As Variant, ByVal end1 As Variant, Optional ByVal second As Boolean = False)
    Dim ws As Object
    Set ws = ThisWorkbook.Worksheets(SheetName("Iterations"))
    ws.Range("A2:C10").ClearContents
    ws.Range("A2").Value = "P1": ws.Range("B2").Value = start1: ws.Range("C2").Value = end1
    If second Then ws.Range("A3").Value = "P2": ws.Range("B3").Value = "2026-10-15": ws.Range("C3").Value = "2026-10-29"
End Sub

Private Function AxisEnds(ByVal ch As Object, ByVal asNumber As Boolean) As String
    ' "first|last" of the X axis of the first series (numbers: first day | days between first and last); "?" when unreadable.
    Dim xv As Variant
    On Error Resume Next
    xv = ch.SeriesCollection(1).XValues
    If asNumber Then
        AxisEnds = CLng(CDbl(xv(LBound(xv)))) & "|" & (CLng(CDbl(xv(UBound(xv)))) - CLng(CDbl(xv(LBound(xv)))))
    Else
        AxisEnds = CStr(xv(LBound(xv))) & "|" & CStr(xv(UBound(xv)))
    End If
    If Err.Number <> 0 Then AxisEnds = "?"
    Err.Clear
End Function

Private Function ChartsBuildError() As Long
    On Error Resume Next
    BuildChartsIn ThisWorkbook
    ChartsBuildError = Err.Number - vbObjectError
    Err.Clear
End Function

Private Function FlowBuildError() As Long
    On Error Resume Next
    BuildFlowSheetsIn ThisWorkbook
    FlowBuildError = Err.Number - vbObjectError
    Err.Clear
End Function

Private Function IterationCell(ByVal column As String, ByVal row As Long) As String
    ' A cell of the table IterationFeatures as text; "?" when it cannot be read or holds an error (a test must never stop on a dialog).
    Dim v As Variant
    On Error Resume Next
    v = ThisWorkbook.Worksheets(SheetName("IterationFeatures")).ListObjects("IterationFeatures").ListColumns(column).DataBodyRange.Cells(row, 1).Value2
    If Err.Number <> 0 Then IterationCell = "?" Else If IsError(v) Then IterationCell = "?" Else IterationCell = CStr(v)
    Err.Clear
End Function

Private Function ColumnSum(ByVal tableName As String, ByVal column As String) As String
    ' The sum of a column of a table, "?" when it cannot be read or holds an error (a check must not raise).
    Dim v As Variant
    On Error Resume Next
    v = Application.WorksheetFunction.Sum(ThisWorkbook.Worksheets(SheetName(tableName)).ListObjects(tableName).ListColumns(column).DataBodyRange)
    If Err.Number <> 0 Then ColumnSum = "?" Else ColumnSum = CStr(CLng(v))
    On Error GoTo 0
End Function

Private Function TableRows(ByVal sheetKey As String) As Long
    TableRows = ThisWorkbook.Worksheets(SheetName(sheetKey)).ListObjects(sheetKey).ListRows.Count
End Function

Private Function PiRows() As Variant
    Dim rows(1 To 2, 1 To 3) As Variant
    rows(1, 1) = "PI#01-2026": rows(1, 2) = DateSerial(2026, 9, 1): rows(1, 3) = DateSerial(2026, 10, 15)
    rows(2, 1) = "PI#02-2026": rows(2, 2) = "'2026-10-15": rows(2, 3) = "'2026-12-01T00:00:00"   ' text, as Load settings CSV or Rovo writes it
    PiRows = rows
End Function

Public Sub T_FlowBuild()
    ' The flow sheets and the dashboard are built from the plan sheets (FlowFormulas, ChartSpec), and rebuilt to the size
    ' of newer data: nothing of a bigger earlier run stays. Iteration dates come as a date, as text or as a number.
    Dim wb As Object, dash As Object, sheetCount As Long, teams As Long
    Set wb = ThisWorkbook
    Set dash = wb.Worksheets(SheetName("Charts"))                               ' the charts and the team selector (B2)
    ClearFlowSheets wb
    FlowIterations "x", "y"
    FlowRaw True
    Check "flowbuild", "text that is not a date is refused", "Start x", 701, FlowBuildError()
    FlowIterations "1026-10-01", "1026-10-15"
    Check "flowbuild", "a year before 2000 is a typo, not an iteration (it would make 365 000 days of DailyLoad)", "1026", 701, FlowBuildError()
    FlowIterations DateSerial(2026, 10, 15), DateSerial(2026, 10, 1)
    Check "flowbuild", "an iteration that ends before it starts is refused", "15 Oct -> 1 Oct", 704, FlowBuildError()
    FlowIterations "2126-10-01", "2126-10-15"
    Check "flowbuild", "a year from 2100 on is a typo too (20261005 typed as a number is the year 57 000)", "2126", 701, FlowBuildError()
    FlowIterations "2026-10-01 08:00", "2026-10-15 08:00"
    Check "flowbuild", "an iteration is whole days: a time of day is refused, as the Python path refuses it", "08:00", 701, FlowBuildError()
    FlowIterations DateSerial(2026, 10, 1), DateSerial(2026, 10, 15)
    wb.Worksheets(SheetName("PIs")).Range("A2:C2").Value = Array("PI#01-2026", "'05/10/2026", "'2026-10-19")
    Check "flowbuild", "a PI date in French text is refused (the French DATEVALUE read it, the VBA did not)", "05/10/2026", 710, FlowBuildError()
    wb.Worksheets(SheetName("PIs")).Range("A2:C2").Value = Array("PI#01-2026", "'2026-10-19", "'2026-10-05")
    Check "flowbuild", "a PI that ends before it starts is refused", "19 Oct -> 5 Oct", 710, FlowBuildError()
    wb.Worksheets(SheetName("PIs")).Range("A2:C2").ClearContents
    Dim modeCell As Object, oldMode As Variant, cr As Long
    For cr = 2 To LastRow(wb.Worksheets(SheetName("Config")))
        If CStr(wb.Worksheets(SheetName("Config")).Cells(cr, 1).Value) = "FlowLoadMode" Then Set modeCell = wb.Worksheets(SheetName("Config")).Cells(cr, 2)
    Next cr
    If Not modeCell Is Nothing Then
        oldMode = modeCell.Value: modeCell.Value = "Mode A"
        Check "flowbuild", "FlowLoadMode Mode A is refused (every Flow Load was not computable)", "Mode A", 711, FlowBuildError()
        modeCell.Value = oldMode
    End If
    wb.Worksheets(SheetName("Iterations")).Range("A2:C10").ClearContents
    Check "flowbuild", "no iteration is refused", "empty Iterations", 703, FlowBuildError()
    FlowIterations DateSerial(2026, 10, 1), DateSerial(2026, 10, 15), True
    ClearFlowSheets wb
    Check "flowbuild", "no extracted data is refused", "empty Issues", 702, FlowBuildError()
    Check "flowbuild", "the charts need the tables: 3. Draw before 2. Compute is refused", "no Metrics", 712, ChartsBuildError()
    FlowRaw True
    BuildFlowSheetsIn wb
    teams = 9                                                    ' ART and the eight team projects of params\projects.md
    Check "flowbuild", "ItemFlow: one row per issue", "2 issues", 2, TableRows("ItemFlow")
    Check "flowbuild", "Issues and StatusIntervals become tables", "2 + 3 rows", "2|3", TableRows("Issues") & "|" & TableRows("StatusIntervals")
    Check "flowbuild", "DailyLoad: one row per day and team", "28 days x 9", 28 * teams, TableRows("DailyLoad")
    Check "flowbuild", "Metrics: one row per iteration and team", "2 iterations x 9", 2 * teams, TableRows("Metrics")
    Check "flowbuild", "ten charts, on the sheet Charts", "Charts", 10, dash.ChartObjects.Count
    Check "flowbuild", "and none on the Dashboard", "Dashboard", 0, wb.Worksheets(SheetName("Dashboard")).ChartObjects.Count
    Check "flowbuild", "DataChecks: one row per check of ChecksSpec", "29 checks", 29, TableRows("DataChecks")
    Check "flowbuild", "DataChecks: the first check, and the check that needs a field is not computable", "DC01 / DC11", "DC01|OK|not computable", _
          wb.Worksheets(SheetName("DataChecks")).Range("A2").Value & "|" & wb.Worksheets(SheetName("DataChecks")).Range("J2").Value & "|" & wb.Worksheets(SheetName("DataChecks")).Range("J12").Value
    Check "flowbuild", "FillRates: the system fields x teams x issue types (the fields not configured are left out)", "3 x 9 x 2", 54, TableRows("FillRates")
    Check "flowbuild", "IssueLinks is a table even without any link", "rows", 1, TableRows("IssueLinks")
    Check "flowbuild", "one series per team in the Velocity chart", "Velocity", teams, dash.ChartObjects("velocity").Chart.SeriesCollection.Count
    Check "flowbuild", "the Velocity chart writes the value above the bars of its first series (a 0 has no bar), the Flow Time chart does not", "labels", "yes|no", _
          IIf(dash.ChartObjects("velocity").Chart.SeriesCollection(1).HasDataLabels, "yes", "no") & "|" & _
          IIf(dash.ChartObjects("flow_time").Chart.SeriesCollection(1).HasDataLabels, "yes", "no")
    Check "flowbuild", "the Cycle time chart has one series per Active or Wait status of the Feature workflow, stacked at 100 %", "series|type", _
          StatusNames(wb).Count & "|53", _
          dash.ChartObjects("pi_status_time").Chart.SeriesCollection.Count & "|" & dash.ChartObjects("pi_status_time").Chart.ChartType
    Check "flowbuild", "the table PIStatusTime has one row per PI and status", "PIs x statuses", TableRows("PIMetrics") * StatusNames(wb).Count, TableRows("PIStatusTime")
    Check "flowbuild", "the table VelocityCheck has the rows of Metrics, the same Velocity and a control Gap of 0", "rows|Velocity|Gap", _
          TableRows("Metrics") & "|" & ColumnSum("Metrics", "Velocity") & "|0", _
          TableRows("VelocityCheck") & "|" & ColumnSum("VelocityCheck", "Velocity") & "|" & ColumnSum("VelocityCheck", "Gap")
    Check "flowbuild", "the sheet MetricItems holds a formula under the header of each list (the lists themselves are checked by the Excel scenarios)", "E3 | S3", "yes|yes", _
          IIf(ThisWorkbook.Worksheets(SheetName("MetricItems")).Range("E3").HasFormula, "yes", "no") & "|" & _
          IIf(ThisWorkbook.Worksheets(SheetName("MetricItems")).Range("S3").HasFormula, "yes", "no")
    Check "flowbuild", "the table Reconciliation has its 19 controls", "rows", 19, TableRows("Reconciliation")
    Check "flowbuild", "the table FlowDistCheck has the rows of Metrics, the Velocity and a control Gap of 0", "rows|Done|Gap", _
          TableRows("Metrics") & "|" & ColumnSum("Metrics", "Velocity") & "|0", _
          TableRows("FlowDistCheck") & "|" & ColumnSum("FlowDistCheck", "Done") & "|" & ColumnSum("FlowDistCheck", "Gap")
    Check "flowbuild", "the table FlowLoadCheck has the rows of Metrics and a control Gap of 0", "rows|Gap", TableRows("Metrics") & "|0", _
          TableRows("FlowLoadCheck") & "|" & ColumnSum("FlowLoadCheck", "Gap")
    Check "flowbuild", "the table FlowEfficiencyCheck has the rows of Metrics and a control Gap of 0", "rows|Gap", TableRows("Metrics") & "|0", _
          TableRows("FlowEfficiencyCheck") & "|" & ColumnSum("FlowEfficiencyCheck", "Gap")
    Check "flowbuild", "the table FlowTimeCheck has the rows of Metrics, the Flow Time count and a control Gap of 0", "rows|Counted|Gap", _
          TableRows("Metrics") & "|" & ColumnSum("Metrics", "FlowTimeCount") & "|0", _
          TableRows("FlowTimeCheck") & "|" & ColumnSum("FlowTimeCheck", "Counted") & "|" & ColumnSum("FlowTimeCheck", "Gap")
    Check "flowbuild", "one series per work type in the Distribution chart", "flow_distribution", 5, dash.ChartObjects("flow_distribution").Chart.SeriesCollection.Count
    Check "flowbuild", "the Predictability chart has two more series: the 80-100 % band", "Predictability", teams + 2, dash.ChartObjects("flow_predictability").Chart.SeriesCollection.Count
    Check "flowbuild", "the daily Flow Load curve has one series per team and one point per day", "28 days", teams & "|" & 28, _
          dash.ChartObjects("flow_load_daily").Chart.SeriesCollection.Count & "|" & UBound(dash.ChartObjects("flow_load_daily").Chart.SeriesCollection(1).Values)
    Check "flowbuild", "the iteration axis of a chart lists the iterations in order", "P1, P2", "P1|P2", AxisEnds(dash.ChartObjects("velocity").Chart, False)
    Check "flowbuild", "Feature Predictability: one series, ART, and the 80-100 % band", "ART", "3|ART", _
          dash.ChartObjects("feature_predictability").Chart.SeriesCollection.Count & "|" & dash.ChartObjects("feature_predictability").Chart.SeriesCollection(1).Name
    Check "flowbuild", "the band lines carry markers (long bars): one PI shows no line, only its points", "xlLineMarkers, xlMarkerStyleDash", "65|-4115|65|-4115", _
          dash.ChartObjects("feature_predictability").Chart.SeriesCollection(2).ChartType & "|" & dash.ChartObjects("feature_predictability").Chart.SeriesCollection(2).MarkerStyle & "|" & _
          dash.ChartObjects("flow_predictability").Chart.SeriesCollection(teams + 2).ChartType & "|" & dash.ChartObjects("flow_predictability").Chart.SeriesCollection(teams + 2).MarkerStyle
    Check "flowbuild", "without a PI in the sheet PIs the PI axis has one row, a gap", "no PI", "(no PI in the sheet PIs)|(no PI in the sheet PIs)", _
          AxisEnds(dash.ChartObjects("feature_predictability").Chart, False)
    Check "flowbuild", "PIMetrics: one row, not computable, without a PI", "no PI", "1|0|not computable", _
          TableRows("PIMetrics") & "|" & wb.Worksheets(SheetName("PIMetrics")).Range("D2").Value & "|" & wb.Worksheets(SheetName("PIMetrics")).Range("F2").Value
    Check "flowbuild", "the day axis starts on the first day, one row per day", "2026-10-01, 28 days", "46296|27", _
          AxisEnds(dash.ChartObjects("flow_load_daily").Chart, True)
    Check "flowbuild", "the formulas give the issue keys", "ItemFlow.Key", "TTT-1|AAA-1", _
          wb.Worksheets(SheetName("ItemFlow")).ListObjects("ItemFlow").ListColumns("Key").DataBodyRange.Cells(1, 1).Value & "|" & _
          wb.Worksheets(SheetName("ItemFlow")).ListObjects("ItemFlow").ListColumns("Key").DataBodyRange.Cells(2, 1).Value
    Check "flowbuild", "the Program is the team ART, a team project its own team", "ItemFlow.Team", "ART|AAA", _
          wb.Worksheets(SheetName("ItemFlow")).ListObjects("ItemFlow").ListColumns("Team").DataBodyRange.Cells(1, 1).Value & "|" & _
          wb.Worksheets(SheetName("ItemFlow")).ListObjects("ItemFlow").ListColumns("Team").DataBodyRange.Cells(2, 1).Value
    Check "flowbuild", "the mode is read from Config", "Controls!B2", "A", wb.Worksheets(SheetName("Controls")).Range("B2").Value
    sheetCount = wb.Worksheets.Count
    ' a second, smaller run: no row, no sheet and no chart of the first one stays
    FlowRaw False
    FlowIterations 46296, 46310                                      ' 2026-10-01 .. 2026-10-15 as plain numbers
    dash.Range("B2").Value = "Nowhere"
    BuildFlowSheetsIn wb
    Check "flowbuild", "smaller data: ItemFlow shrinks", "1 issue", 1, TableRows("ItemFlow")
    Check "flowbuild", "smaller data: DailyLoad and Metrics shrink", "14 days, 1 iteration", (14 * teams) & "|" & teams, TableRows("DailyLoad") & "|" & TableRows("Metrics")
    Check "flowbuild", "smaller data: StatusIntervals", "1 interval", 1, TableRows("StatusIntervals")
    Check "flowbuild", "smaller data: FillRates follows the issue types", "3 x 9 x 1", 27, TableRows("FillRates")
    Check "flowbuild", "the sheets are the same", "count", sheetCount, wb.Worksheets.Count
    Check "flowbuild", "still ten charts", "Charts", 10, dash.ChartObjects.Count
    BuildTablesIn wb
    Check "flowbuild", "2. Compute: the tables made again, no chart (they read the tables)", "0 charts, Metrics", "0|yes|no", _
          dash.ChartObjects.Count & "|" & IIf(SheetExists(wb, "Metrics"), "yes", "no") & "|" & IIf(SheetExists(wb, "ChartData"), "yes", "no")
    BuildChartsIn wb
    Check "flowbuild", "3. Draw: the ten charts again from the tables", "10 charts", "10|yes", dash.ChartObjects.Count & "|" & IIf(SheetExists(wb, "ChartData"), "yes", "no")
    Check "flowbuild", "an unknown team in the selector goes back to All teams", "Nowhere", "All teams", dash.Range("B2").Value
    dash.Range("B2").Value = "AAA"
    wb.Worksheets(SheetName("PIs")).Range("A2:C3").Value = PiRows()                 ' two PIs: the PI axis lists them in order
    BuildFlowSheetsIn wb
    Check "flowbuild", "a known team stays selected", "AAA", "AAA", dash.Range("B2").Value
    Check "flowbuild", "the PI axis lists the PIs of the sheet PIs in order", "two PIs", "PI#01-2026|PI#02-2026", _
          AxisEnds(dash.ChartObjects("feature_predictability").Chart, False)
    Check "flowbuild", "PIMetrics: one row per PI, its dates looked up in the sheet PIs, a date or an ISO text", "two PIs", _
          "2|" & CLng(DateSerial(2026, 10, 15)) & "|PI#02-2026|" & CLng(DateSerial(2026, 12, 1)), _
          TableRows("PIMetrics") & "|" & wb.Worksheets(SheetName("PIMetrics")).Range("C2").Value2 & "|" & wb.Worksheets(SheetName("PIMetrics")).Range("A3").Value & _
          "|" & wb.Worksheets(SheetName("PIMetrics")).Range("C3").Value2
    ' Features in PIs that the sheet PIs lacks (AL, 2026-11-12): after the sheet's, by year then number, not computable; a PI
    ' of the sheet, a malformed one (DC29) and the PI of a story add no row
    FlowRaw True, Array("PI#01-2027", "PI#04-2026", "PI#04-2026", "PI#02-2026", "PI 4", "pi#05-2026"), "PI#06-2026"
    BuildFlowSheetsIn wb
    Check "flowbuild", "PIMetrics: the PIs of Features missing from the sheet PIs come after it, in order, not computable", _
          "PI#04-2026 (2 features), PI#01-2027", "4|PI#04-2026|2||not computable|PI#01-2027|PI#01-2026|PI#01-2027", _
          TableRows("PIMetrics") & "|" & wb.Worksheets(SheetName("PIMetrics")).Range("A4").Value & "|" & wb.Worksheets(SheetName("PIMetrics")).Range("D4").Value & _
          "|" & wb.Worksheets(SheetName("PIMetrics")).Range("C4").Value & "|" & wb.Worksheets(SheetName("PIMetrics")).Range("F4").Value & "|" & _
          wb.Worksheets(SheetName("PIMetrics")).Range("A5").Value & "|" & AxisEnds(dash.ChartObjects("feature_predictability").Chart, False)
    ' IterationFeatures: one row per iteration, PI that holds its start and team (the ART, then the teams of the Features, alphabetical)
    Check "flowbuild", "IterationFeatures: 1 iteration x 3 teams, in order, in the PI that holds its start", "ART, team-a, team-b", _
          "3|ART|team-a|team-b|PI#01-2026", TableRows("IterationFeatures") & "|" & IterationCell("Team", 1) & "|" & IterationCell("Team", 2) & "|" & _
          IterationCell("Team", 3) & "|" & IterationCell("PI", 3)
    Check "flowbuild", "IterationFeatures: nothing engaged in PI#01-2026 (its only Feature has the PI of 2027), the baseline is the start of P1", "P1", _
          "0|0|0|" & CLng(DateSerial(2026, 10, 1)), IterationCell("Engaged", 1) & "|" & IterationCell("Released", 2) & "|" & IterationCell("CarryOver", 3) & "|" & _
          IterationCell("Baseline", 1)
    Check "flowbuild", "PIIntervals: one row per Feature with a PI, helpers from ItemFlow", "6 features + 1 story", "7", TableRows("PIIntervals")
    FlowRaw False
    wb.Worksheets(SheetName("PIs")).Range("A2:C3").ClearContents
    ' Excel left in manual mode by the user: the build gives the mode back and still ends with every figure calculated
    ' (guard only: removing CalculateFull for this case changed nothing observable, entering a formula calculates it)
    Application.Calculation = -4135                               ' xlCalculationManual
    BuildFlowSheetsIn wb
    Check "flowbuild", "manual calculation: the mode is given back and nothing is left to calculate", "Excel in manual mode", "-4135|0|OK", _
          Application.Calculation & "|" & Application.CalculationState & "|" & wb.Worksheets(SheetName("DataChecks")).Range("J2").Value   ' 0 = xlDone: nothing left to calculate
    Application.Calculation = -4105                               ' xlCalculationAutomatic
    ' leave the workbook as it was
    dash.Range("B2").Value = "All teams"
    wb.Worksheets(SheetName("Iterations")).Range("A2:C10").ClearContents
    ClearFlowSheets wb
    Check "flowbuild", "cleared: the derived sheets are gone", "Metrics", False, SheetExists(wb, "Metrics")
    Check "flowbuild", "cleared: the data checks and fill rates are gone too", "DataChecks", "False|False", _
          IIf(SheetExists(wb, "DataChecks"), "True", "False") & "|" & IIf(SheetExists(wb, "FillRates"), "True", "False")
End Sub

' ---------- Text helpers: plain VBA, proved against the old ADODB / MSXML code kept here only as the oracle ----------
' Every call that can raise goes through a function that reports "error n" as text: a test must never stop on a dialog.

Private Function RefEncode(ByVal s As String) As Byte()
    Dim st As Object, empty_() As Byte
    If Len(s) = 0 Then RefEncode = empty_: Exit Function
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8": st.Open
    st.WriteText s
    st.Position = 0: st.Type = 1: st.Position = 3          ' skip the UTF-8 BOM
    RefEncode = st.Read
    st.Close
End Function

Private Function RefDecode(ByRef bytes() As Byte) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1: st.Open
    st.Write bytes
    st.Position = 0: st.Type = 2: st.Charset = "utf-8"
    RefDecode = st.ReadText
    st.Close
End Function

Private Function RefBase64(ByRef bytes() As Byte) As String
    Dim node As Object
    Set node = CreateObject("MSXML2.DOMDocument.6.0").createElement("b64")
    node.DataType = "bin.base64"
    node.nodeTypedValue = bytes
    RefBase64 = Replace(Replace(node.Text, vbLf, ""), vbCr, "")
End Function

Private Function HexOf(ByRef bytes() As Byte) As String
    Dim i As Long, out As String
    On Error Resume Next
    For i = LBound(bytes) To UBound(bytes)
        If Err.Number <> 0 Then Err.Clear: Exit For
        out = out & Right$("0" & Hex$(bytes(i)), 2) & " "
    Next i
    Err.Clear
    HexOf = Trim$(out)
End Function

Private Function CodesOf(ByVal s As String) As String
    ' "41 FFFD 42": the UTF-16 code units of a text, readable in the results sheet whatever the characters are.
    Dim i As Long, out As String
    For i = 1 To Len(s)
        out = out & Hex$(AscW(Mid$(s, i, 1)) And &HFFFF&) & " "
    Next i
    CodesOf = Trim$(out)
End Function

Private Function BytesOf(ParamArray v() As Variant) As Byte()
    Dim b() As Byte, i As Long
    ReDim b(0 To UBound(v))
    For i = 0 To UBound(v)
        b(i) = CByte(v(i))
    Next i
    BytesOf = b
End Function

Private Function AsciiBytes(ByVal s As String) As Byte()
    AsciiBytes = StrConv(s, vbFromUnicode)
End Function

Private Function RefHex(ByVal s As String) As String
    ' The bytes ADODB gives for a text, as hex.
    On Error Resume Next
    RefHex = HexOf(RefEncode(s))
    If Err.Number <> 0 Then RefHex = "ADODB error " & Err.Number: Err.Clear
End Function

Private Function MineHex(ByVal s As String) As String
    ' The bytes Utf8Encode gives for a text, as hex.
    On Error Resume Next
    MineHex = HexOf(Utf8Encode(s))
    If Err.Number <> 0 Then MineHex = "Utf8Encode error " & Err.Number: Err.Clear
End Function

Private Function DecodedRef(ByVal s As String) As String
    ' Utf8Decode of the bytes ADODB gives for a text.
    On Error Resume Next
    DecodedRef = Utf8Decode(RefEncode(s))
    If Err.Number <> 0 Then DecodedRef = "Utf8Decode error " & Err.Number: Err.Clear
End Function

Private Function RefDecodedMine(ByVal s As String) As String
    ' ADODB reads the bytes Utf8Encode gives for a text.
    On Error Resume Next
    RefDecodedMine = RefDecode(Utf8Encode(s))
    If Err.Number <> 0 Then RefDecodedMine = "ADODB error " & Err.Number: Err.Clear
End Function

Private Function DecodedCodes(ByRef bytes() As Byte) As String
    ' The code units Utf8Decode gives for some bytes.
    On Error Resume Next
    DecodedCodes = CodesOf(Utf8Decode(bytes))
    If Err.Number <> 0 Then DecodedCodes = "Utf8Decode error " & Err.Number: Err.Clear
End Function

Private Function Base64Safe(ByRef bytes() As Byte) As String
    On Error Resume Next
    Base64Safe = Base64Encode(bytes)
    If Err.Number <> 0 Then Base64Safe = "Base64Encode error " & Err.Number: Err.Clear
End Function

Private Function RefBase64Safe(ByRef bytes() As Byte) As String
    On Error Resume Next
    RefBase64Safe = RefBase64(bytes)
    If Err.Number <> 0 Then RefBase64Safe = "MSXML error " & Err.Number: Err.Clear
End Function

Private Function SampleTexts() As Variant
    Dim long_ As String, i As Long
    For i = 1 To 3500
        long_ = long_ & "{""summary"":""R" & ChrW(233) & "union d'" & ChrW(233) & "quipe n" & ChrW(176) & i & " " & ChrW(8364) & " " & _
                ChrW(&HD83D) & ChrW(&HDE00) & """},"
    Next i
    SampleTexts = Array("Hello, world 0123", _
        "Caf" & ChrW(233) & " d" & ChrW(233) & "j" & ChrW(224) & " vu " & ChrW(339), _
        "20 " & ChrW(8364), ChrW(945) & ChrW(946) & ChrW(947), ChrW(&H4E2D) & ChrW(&H6587), _
        ChrW(&HD83D) & ChrW(&HDE00), ChrW(&HD83C) & ChrW(&HDF55) & "x", _
        ChrW(&H7F) & ChrW(&H80) & ChrW(&H7FF) & ChrW(&H800) & ChrW(&HD7FF) & ChrW(&HE000) & ChrW(&HFFFD), _
        "{""summary"":""R" & ChrW(233) & "sum" & ChrW(233) & " " & ChrW(8364) & " " & ChrW(&HD83D) & ChrW(&HDE00) & """}", _
        long_)
End Function

Public Sub T_Utf8Oracle()
    ' Utf8Encode and Utf8Decode give the very bytes and text of ADODB.Stream, on every kind of character.
    Dim samples As Variant, i As Long, s As String
    samples = SampleTexts()
    For i = LBound(samples) To UBound(samples)
        s = CStr(samples(i))
        Check "text", "UTF-8 bytes of text " & i & " are the ones of ADODB", "len " & Len(s), True, (MineHex(s) = RefHex(s))
        Check "text", "decoding the bytes of ADODB gives the text " & i, "len " & Len(s), True, (DecodedRef(s) = s)
        Check "text", "ADODB decodes our bytes the same way " & i, "len " & Len(s), True, (RefDecodedMine(s) = s)
    Next i
    Check "text", "decoding an uninitialised array gives an empty text", "no bytes", "", Utf8Decode(Empty)
End Sub

Public Sub T_Utf8Invalid()
    ' Bad UTF-8 never raises: each bad byte becomes U+FFFD and the text around it is kept. A leading BOM is dropped.
    Dim bad As String
    bad = "FFFD"
    Check "text", "a stray continuation byte", "41 80 42", "41 " & bad & " 42", DecodedCodes(BytesOf(&H41, &H80, &H42))
    Check "text", "a truncated sequence", "41 E2 82", "41 " & bad & " " & bad, DecodedCodes(BytesOf(&H41, &HE2, &H82))
    Check "text", "a truncated sequence before text", "E2 82 42", bad & " " & bad & " 42", DecodedCodes(BytesOf(&HE2, &H82, &H42))
    Check "text", "an overlong form of a character (C0 80)", "41 C0 80 42", "41 " & bad & " " & bad & " 42", _
          DecodedCodes(BytesOf(&H41, &HC0, &H80, &H42))
    Check "text", "an overlong 3-byte form (E0 80 80)", "E0 80 80", bad & " " & bad & " " & bad, DecodedCodes(BytesOf(&HE0, &H80, &H80))
    Check "text", "an encoded surrogate (ED A0 80)", "ED A0 80", bad & " " & bad & " " & bad, DecodedCodes(BytesOf(&HED, &HA0, &H80))
    Check "text", "a byte that starts no sequence (F8)", "F8", bad, DecodedCodes(BytesOf(&HF8))
    Check "text", "a value above U+10FFFF (F4 90 80 80)", "F4 90 80 80", bad & " " & bad & " " & bad & " " & bad, _
          DecodedCodes(BytesOf(&HF4, &H90, &H80, &H80))
    Check "text", "a truncated 4-byte sequence (F0 9F 98)", "F0 9F 98", bad & " " & bad & " " & bad, DecodedCodes(BytesOf(&HF0, &H9F, &H98))
    Check "text", "a leading BOM is dropped", "EF BB BF 41", "41", DecodedCodes(BytesOf(&HEF, &HBB, &HBF, &H41))
    Check "text", "a BOM alone gives an empty text", "EF BB BF", "", DecodedCodes(BytesOf(&HEF, &HBB, &HBF))
    Check "text", "a valid 4-byte sequence (F0 9F 98 80) is a surrogate pair", "F0 9F 98 80", "D83D DE00", _
          DecodedCodes(BytesOf(&HF0, &H9F, &H98, &H80))
    Check "text", "a lone high surrogate is encoded as U+FFFD", "D83D", "EF BF BD", MineHex(ChrW(&HD83D))
    Check "text", "a lone low surrogate is encoded as U+FFFD", "DE00 x", "EF BF BD 78", MineHex(ChrW(&HDE00) & "x")
    Check "text", "a high surrogate not followed by a low one", "D83D x", "EF BF BD 78", MineHex(ChrW(&HD83D) & "x")
End Sub

Public Sub T_Base64Oracle()
    ' The seven vectors of RFC 4648, then 200 byte arrays of 0 to 60 bytes against MSXML.
    Dim words As Variant, expected As Variant, i As Long, n As Long, x As Long, b() As Byte, k As Long, same As Boolean
    words = Array("", "f", "fo", "foo", "foob", "fooba", "foobar")
    expected = Array("", "Zg==", "Zm8=", "Zm9v", "Zm9vYg==", "Zm9vYmE=", "Zm9vYmFy")
    Check "text", "RFC 4648: empty input", "(nothing)", "", Base64Safe(Utf8Encode(""))
    For i = 1 To 6
        Check "text", "RFC 4648: " & words(i), CStr(words(i)), CStr(expected(i)), Base64Safe(AsciiBytes(CStr(words(i))))
    Next i
    x = 7
    same = True
    For i = 1 To 200
        n = i Mod 61
        If n = 0 Then n = 1
        ReDim b(0 To n - 1)
        For k = 0 To n - 1
            x = (x * 75 + 74) Mod 65537
            b(k) = x Mod 256
        Next k
        If Base64Safe(b) <> RefBase64Safe(b) Then same = False: Exit For
    Next i
    Check "text", "200 byte arrays give the Base64 of MSXML", "lengths 1 to 60", True, same
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
    Check "text", "a time without seconds, Z", "2026-10-01T09:12Z", CDbl(base + TimeSerial(9, 12, 0)), CDbl(ParseJiraDate("2026-10-01T09:12Z"))
    Check "text", "a time without seconds, +0200: the offset is not read as seconds", "2026-10-01T09:12+0200", CDbl(base + TimeSerial(7, 12, 0)), _
          CDbl(ParseJiraDate("2026-10-01T09:12+0200"))
    Check "text", "a time without seconds, -05:00", "2026-10-01T09:12-05:00", CDbl(base + TimeSerial(14, 12, 0)), CDbl(ParseJiraDate("2026-10-01T09:12-05:00"))
    Check "text", "an hour alone is no time", "2026-10-01T09+0200", True, IsEmpty(ParseJiraDate("2026-10-01T09+0200"))
    Check "text", "letters in the date do not raise", "20x6-10-01T09:12:33Z", True, IsEmpty(ParseJiraDate("20x6-10-01T09:12:33Z"))
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
    Check "http", "the log says why a call waits", "429", True, InStr(FileText(gLogPath), " RETRY HTTP 429 at attempt 1 of ") > 0
    Check "http", "a RETRY line reads status, attempt and wait", "429", "HTTP 429 at attempt 1 of 5: next attempt in 2.345 s", _
          RetryText(429, "", 1, 5, 2345)
    Check "http", "with no answer, the support code of the WinHTTP error", "12002", _
          "no answer (J2D-NET-12002) at attempt 2 of 5: next attempt in 1.000 s", _
          RetryText(0, "-2147012894 The operation timed out", 2, 5, 1000)
    Check "http", "with no answer and no error number, the generic code", "no number", _
          "no answer (J2D-NET-001) at attempt 1 of 3: next attempt in 0.010 s", RetryText(0, "", 1, 3, 10)
End Sub

Public Sub T_HttpNetworkAndCancel()
    ' A network error (status 0) is retried twice; Cancel ends the request at once; an HTML answer is explained.
    Dim cfg As Object, resp As Object, raised As Boolean, data As Collection, names As Object, isLast As Boolean
    Set cfg = TestConfig()
    cfg("RetryBaseMs") = 1: cfg("RetryMaxMs") = 2
    FakeReset
    FakeAdd "GET", "https://x/net", 0, ""
    FakeAdd "GET", "https://x/net", 200, "{}"
    Set resp = HttpRequest(cfg, "GET", "https://x/net")
    Check "http", "one network error then the answer", "0 then 200", "200|2", resp("Status") & "|" & resp("Attempts")
    FakeReset
    FakeAdd "GET", "https://x/net", 0, ""
    FakeAdd "GET", "https://x/net", 0, ""
    FakeAdd "GET", "https://x/net", 0, ""
    FakeAdd "GET", "https://x/net", 200, "{}"
    cfg("MaxRetries") = 6
    Set resp = HttpRequest(cfg, "GET", "https://x/net")
    Check "http", "a network error is given up after two retries", "0 0 0 200", "0|3", resp("Status") & "|" & resp("Attempts")
    FakeReset
    FakeAdd "GET", "https://x/slow", 503, ""
    FakeAdd "GET", "https://x/slow", 200, "{}"
    gHttpCancel = True
    On Error Resume Next
    Call HttpRequest(cfg, "GET", "https://x/slow")
    raised = (Err.Number = vbObjectError + 610): Err.Clear
    On Error GoTo 0
    gHttpCancel = False
    Check "http", "Cancel stops the request with error 610", "gHttpCancel", True, raised
    Check "http", "the jitter never gives an instant retry", "backoff 2000", True, RetryDelayMs(NewHttpConfig(), 0, "") >= 1000
    cfg("MaxRetries") = 0
    FakeReset
    FakeAdd "GET", JiraSearchUrl(cfg("BaseUrl"), "project = X", "", 100, "key", ""), 200, "<html>Login</html>"
    Set data = New Collection: Set names = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Call JiraSearchPage(cfg, "project = X", "", data, names, "key", "", isLast)
    raised = (Err.Number = vbObjectError + 607)
    Err.Clear
    On Error GoTo 0
    Check "jira", "a 200 answer that is not JSON is explained", "<html>", True, raised
    FakeReset
    FakeAdd "GET", JiraSearchUrl(cfg("BaseUrl"), "project = X", "t1", 100, "key", ""), 200, J("{'issues':[],'isLast':false,'nextPageToken':'t1'}")
    On Error Resume Next
    Call JiraSearchPage(cfg, "project = X", "t1", data, names, "key", "", isLast)
    raised = (Err.Number = vbObjectError + 606)
    Err.Clear
    On Error GoTo 0
    Check "jira", "a page token that comes back is refused", "t1 again", True, raised
    ' Cancel pressed during the wait between two attempts: no request is sent after it
    cfg("MaxRetries") = 2
    FakeReset
    FakeAdd "GET", "https://x/wait", 503, ""
    FakeAdd "GET", "https://x/wait", 200, "{}"
    gHttpCancelInWait = True
    On Error Resume Next
    Call HttpRequest(cfg, "GET", "https://x/wait")
    raised = (Err.Number = vbObjectError + 610): Err.Clear
    On Error GoTo 0
    gHttpCancelInWait = False: gHttpCancel = False
    Check "http", "Cancel during a retry wait stops with error 610", "503 then cancel", True, raised
    Check "http", "Cancel during a retry wait sends no further request", "503 then cancel", 1, gHttpLog.Count
End Sub

Public Sub T_AbortExtraction()
    ' AbortExtraction (workbook closing, end of a test) forgets the run and leaves no text in the status bar of Excel.
    Application.StatusBar = "Reading Jira..."
    AbortExtraction
    Check "actions", "the status bar text of an aborted run is cleared", "text set, then abort", True, (UCase$(CStr(Application.StatusBar)) = "FALSE")     ' Excel gives False or the text FALSE
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
    FakeAdd "GET", "https://x/d", 500, "", ""
    FakeAdd "GET", "https://x/d", 200, ""
    Set resp = HttpRequest(cfg, "GET", "https://x/d")
    Check "http", "a 5xx without Retry-After is retried (like the PowerShell script)", "503, 500, 200", 200, resp("Status")
    Check "http", "three attempts were made", "503, 500, 200", 3, resp("Attempts")
    Check "http", "no queued response gives 404", "(none)", 404, HttpRequest(cfg, "GET", "https://x/none")("Status")
    Check "http", "a 4xx is not retried", "404", 1, HttpRequest(cfg, "GET", "https://x/none")("Attempts")
End Sub

Public Sub T_ChangelogChunks()
    Dim cfg As Object, keys As New Collection, i As Long, rows As Collection, url As String
    Set cfg = TestConfig()
    For i = 1 To CHANGELOG_CHUNK + 20
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
    Check "jira", "POST calls (2 chunks, 1 extra page)", "chunk + 20 keys", 3, gHttpLog.Count
    Check "jira", "first chunk size", "call 1", CHANGELOG_CHUNK, CountOf(gHttpLog(1)(2), """K-")
    Check "jira", "first call has no token", "call 1", 0, CountOf(gHttpLog(1)(2), "nextPageToken")
    Check "jira", "second call sends token", "call 2", True, (InStr(gHttpLog(2)(2), """nextPageToken"":""p2""") > 0)
    Check "jira", "second chunk size", "call 3", 20, CountOf(gHttpLog(3)(2), """K-")
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

Public Sub T_IntervalsUnreadableDate()
    ' A change whose date ParseJiraDate cannot read is Empty, and Empty - Date is a NEGATIVE DATE in VBA: Excel refuses a
    ' block that holds one ("Out of memory", error 7), so no interval may hold a date-typed duration. Found on a real Jira.
    Dim info As Object, changes As New Collection, cats As Object, out As Collection, iv As Variant, d0 As Date, bad As Long
    Dim ws As Object, rows As New Collection, errNumber As Long, i As Long
    d0 = DateSerial(2026, 10, 1)
    Set info = CreateObject("Scripting.Dictionary")
    info("K-4") = Array(d0, "Termine")
    Set cats = CreateObject("Scripting.Dictionary")
    changes.Add Array("K-4", "c1", ParseJiraDate("not a date"), "Nouveau", "En analyse")
    changes.Add Array("K-4", "c2", d0 + 3, "En analyse", "Termine")
    Set out = BuildStatusIntervals(info, changes, d0 + 9, cats)
    Check "flow", "unreadable date reads as Empty", "ParseJiraDate", True, IsEmpty(changes(1)(2))
    For Each iv In out
        If VarType(iv(6)) = vbDate Then bad = bad + 1
    Next iv
    Check "flow", "no duration is a date", "3 intervals", 0, bad
    Check "flow", "unreadable date is flagged", "row 1", "BEFORE_CREATED", out(1)(9)
    For Each iv In out
        rows.Add Array(iv(0), "Team", iv(1), iv(2), iv(3), iv(4), iv(5), iv(6), iv(7), iv(8), iv(9))
    Next iv
    Set ws = ThisWorkbook.Worksheets("TestScratch")
    On Error Resume Next
    WriteTable ws, Array("Key", "Workflow", "Seq", "Status", "Category", "EnteredUtc", "LeftUtc", _
                         "DurationDays", "IsOpen", "ChangeId", "Check"), rows
    errNumber = Err.Number
    On Error GoTo 0
    Check "sheet", "the intervals are written", "no error", 0, errNumber
    ws.Cells.Clear                                                    ' the date formats stay after ClearContents: the next test reads 0/1 there
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

Public Sub T_ChangelogRows()
    ' Every item of the changelog becomes a row: no author, the date as Jira sent it next to the date read, a text cut at 255
    ' characters, the value of a person field (assignee...) not kept.
    Dim raw As New Collection, map As Object, out As Collection, r As Variant, all As String, v As Variant
    Set map = CreateObject("Scripting.Dictionary"): map("10") = "A-1"
    raw.Add Array("10", "h1", "2026-10-02T10:00:00.000+0200", "Ann", "a1", "status", "status", "jira", "1", "Nouveau", "3", "Pret")
    raw.Add Array("10", "h2", "not a date", "Ann", "a1", "description", "description", "jira", "", String$(300, "x"), "", "ok")
    raw.Add Array("11", "h3", "2026-10-03T10:00:00.000Z", "Bob", "b1", "assignee", "assignee", "jira", "acc-1", "Mary Martin", "acc-2", "Paul Durand")
    raw.Add Array("10", "h4", "2026-10-04T10:00:00.000Z", "Ann", "a1", "customfield_10050", "Product owner", "custom", "acc-9", "Mary Martin", "acc-8", "Paul Durand")
    raw.Add Array("10", "h5", "2026-10-05T10:00:00.000Z", "Ann", "a1", "description", "description", "jira", "", "Ask [~accountid:5b10ac8d82e05b22cc7d4ef5] and [~jdoe] now", "", "ok")
    Dim users As Object
    Set users = CreateObject("Scripting.Dictionary"): users("customfield_10050") = True
    Set out = ChangelogRows(raw, map, users)
    Check "changelog", "one row per item", "5", 5, out.Count
    Check "changelog", "the key of the issue, its id, the id of the change", "A-1|10|h1", "A-1|10|h1", out(1)(0) & "|" & out(1)(1) & "|" & out(1)(2)
    Check "changelog", "the date read, in UTC", "10:00+0200", DateSerial(2026, 10, 2) + TimeSerial(8, 0, 0), out(1)(3)
    Check "changelog", "the date as Jira sent it", "text", "2026-10-02T10:00:00.000+0200", out(1)(4)
    Check "changelog", "a date that is not one is not read, the text shows it", "not a date", "empty|not a date", IIf(IsEmpty(out(2)(3)), "empty", "read") & "|" & out(2)(4)      ' not CStr(True): "Vrai" on a French PC
    Check "changelog", "a text is cut at 255 characters", "300 -> 255", 255, Len(out(2)(9))
    Check "changelog", "an id that no issue has is kept as it is", "11", "11", out(3)(0)
    Check "changelog", "a person field keeps no value", "assignee", "||||", out(3)(8) & "|" & out(3)(9) & "|" & out(3)(10) & "|" & out(3)(11) & "|"
    For Each r In out
        For Each v In r
            all = all & "|" & CStr(v)
        Next v
    Next r
    Check "changelog", "a custom field that Jira types as a user keeps no value either", "Product owner", "||||", out(4)(8) & "|" & out(4)(9) & "|" & out(4)(10) & "|" & out(4)(11) & "|"
    Check "changelog", "the mentions of a person in a text become [user]", "5b10 and jdoe", "Ask [user] and [user] now", out(5)(9)
    Check "changelog", "no author and no person name anywhere", "none", True, InStr(all, "Ann") = 0 And InStr(all, "Bob") = 0 And InStr(all, "Martin") = 0 And InStr(all, "Durand") = 0 And InStr(all, "5b10") = 0 And InStr(all, "jdoe") = 0
    Check "changelog", "the columns", "12", 12, UBound(ChangelogHeader()) + 1
End Sub

Public Sub T_StatusChangesUnreadableDate()
    ' A status change whose date cannot be read is left out (as the Python path does), counted, and its shape is kept.
    Dim rows As New Collection, map As Object, out As Collection, skipped As Long, shapes As String
    Set map = CreateObject("Scripting.Dictionary")
    rows.Add Array("1", "h1", "2026-10-02T10:00:00.000+0200", "Ann", "a1", "status", "status", "jira", "1", "A", "2", "B")
    rows.Add Array("1", "h2", "not a date", "Ann", "a1", "status", "status", "jira", "2", "B", "3", "C")
    rows.Add Array("1", "h3", "yesterday", "Ann", "a1", "status", "status", "jira", "2", "B", "3", "C")
    rows.Add Array("1", "h4", Null, "Ann", "a1", "status", "status", "jira", "3", "C", "4", "D")
    rows.Add Array("1", "h5", "not a date", "Ann", "a1", "customfield_10020", "Sprint", "custom", "", "", "", "S1")   ' not a status
    Set out = StatusChangesFromChangelog(rows, map, skipped, shapes)
    Check "flow", "the readable change is kept", "1 of 4", 1, out.Count
    Check "flow", "the unreadable ones are counted", "3", 3, skipped
    Check "flow", "their shapes are counted, with no content", "two texts, Null x1", "String:aaa a aaaa x 1; String:aaaaaaaaa x 1; Null x 1", shapes
End Sub

Public Sub T_SheetName()
    ' The suffix says where the data of a sheet comes from (python: sheet_names.of / logical).
    Check "sheets", "typed by the person", "Config", "Config_from_user", SheetName("Config")
    Check "sheets", "read from Jira", "Issues", "Issues_from_Jira", SheetName("Issues")
    Check "sheets", "a proposal read from Jira", "JiraFields", "JiraFields_from_Jira", SheetName("JiraFields")
    Check "sheets", "calculated by Excel", "Metrics", "Metrics_from_excel", SheetName("Metrics")
    Check "sheets", "the dashboard", "Dashboard", "Dashboard_from_excel", SheetName("Dashboard")
    Check "sheets", "a real name is left alone", "Config_from_user", "Config_from_user", SheetName("Config_from_user")
    Check "sheets", "an unknown sheet is left alone", "TestScratch", "TestScratch", SheetName("TestScratch")
    Check "sheets", "the logical name of a real one", "PIs_from_user", "PIs", LogicalSheet("PIs_from_user")
    Check "sheets", "the logical name of a logical one", "PIs", "PIs", LogicalSheet("PIs")
    Check "sheets", "a suffix on an unknown name is kept", "Foo_from_user", "Foo_from_user", LogicalSheet("Foo_from_user")
    Check "sheets", "the longest name fits Excel", "31", True, Len(SheetName("JiraProjectWorkflows")) <= 31
End Sub

Public Sub T_DateKind()
    ' What a value is, from the value only: the date of a changelog is found by it, never by the label of its key.
    Dim no As Variant
    Check "datekind", "milliseconds as a number", "1761134538011", "epoch ms", DateKind(1761134538011#)
    Check "datekind", "milliseconds as text", "1761134538011", "epoch ms", DateKind("1761134538011")
    Check "datekind", "seconds as a number", "1761134538", "epoch s", DateKind(1761134538#)
    Check "datekind", "seconds as text", "1761134538", "epoch s", DateKind("1761134538")
    Check "datekind", "an ISO date-time", "2026-10-01T09:12:33.123+0200", "ISO", DateKind("2026-10-01T09:12:33.123+0200")
    Check "datekind", "an ISO day", "2026-10-01", "ISO", DateKind("2026-10-01")
    For Each no In Array("843000", 843000, "TA-1", "", Null, True)
        Check "datekind", "not a date", "843000, text, empty, Null, True", "none", DateKind(no)
    Next no
    Check "datekind", "an object is not a date", "{a:1}", "none", DateKind(JsonParse(J("{'a':1}")))
    Check "datekind", "an array is not a date", "[]", "none", DateKind(JsonParse("[]"))
End Sub

Public Sub T_IsPlausibleDate()
    ' Between 2000-01-01 and tomorrow: a change is never in the future.
    Dim now0 As Date
    now0 = DateSerial(2026, 11, 14) + TimeSerial(12, 0, 0)
    Check "plausible", "a real epoch", "1761134538011", "yes", IIf(IsPlausibleDate(1761134538011#, now0), "yes", "no")
    Check "plausible", "now", "2026-11-14T12:00:00Z", "yes", IIf(IsPlausibleDate("2026-11-14T12:00:00Z", now0), "yes", "no")
    Check "plausible", "until tomorrow", "2026-11-15T11:00:00Z", "yes", IIf(IsPlausibleDate("2026-11-15T11:00:00Z", now0), "yes", "no")
    Check "plausible", "not later", "2026-11-16T12:00:00Z", "no", IIf(IsPlausibleDate("2026-11-16T12:00:00Z", now0), "yes", "no")
    Check "plausible", "an epoch in 2100", "4102444800000", "no", IIf(IsPlausibleDate(4102444800000#, now0), "yes", "no")
    Check "plausible", "1999", "1999-12-31", "no", IIf(IsPlausibleDate("1999-12-31", now0), "yes", "no")
    Check "plausible", "2000", "2000-01-01", "yes", IIf(IsPlausibleDate("2000-01-01", now0), "yes", "no")
    Check "plausible", "an id", "843000", "no", IIf(IsPlausibleDate("843000", now0), "yes", "no")
    Check "plausible", "Null", "Null", "no", IIf(IsPlausibleDate(Null, now0), "yes", "no")
    Check "plausible", "an object", "{}", "no", IIf(IsPlausibleDate(JsonParse("{}"), now0), "yes", "no")
End Sub

Private Function Histories(ParamArray jsons() As Variant) As Collection
    Dim c As New Collection, i As Long
    For i = LBound(jsons) To UBound(jsons)
        c.Add JsonParse(J(CStr(jsons(i))))
    Next i
    Set Histories = c
End Function

Private Function DateKeyOf(ByVal hs As Collection) As String
    Dim kind As String, others As String, key As String
    key = FindDateKey(hs, DateSerial(2026, 11, 14) + TimeSerial(12, 0, 0), kind, others)
    DateKeyOf = key & "|" & kind & "|" & others
End Function

Public Sub T_FindDateKey()
    ' The key that holds the date of a changelog history is found by the VALUE of each key (the python twin: mapping.find_date_key).
    Dim hs As Collection, i As Long, kind As String, others As String
    Check "datekey", "created, milliseconds (the real Jira)", "created", "created|epoch ms|", _
          DateKeyOf(Histories("{'id':'843000','created':1761134538011,'author':{'displayName':'x'},'items':[]}"))
    Check "datekey", "created, ISO (what the schema says)", "created", "created|ISO|", _
          DateKeyOf(Histories("{'id':'1','created':'2026-10-01T09:12:33.123+0200','items':[]}"))
    Check "datekey", "a renamed key", "createdAt", "createdAt|epoch ms|", DateKeyOf(Histories("{'id':'1','createdAt':1761134538011,'items':[]}"))
    Check "datekey", "a renamed key, ISO, with another text", "timestamp", "timestamp|ISO|", _
          DateKeyOf(Histories("{'id':'1','timestamp':'2026-10-01T09:12:33Z','note':'x','items':[]}"))
    Check "datekey", "several: created wins, the other is named", "created, other", "created|epoch ms|other", _
          DateKeyOf(Histories("{'id':'1','other':1761134538011,'created':1761134538011}"))
    Check "datekey", "several: the first, the others named", "first, second", "first|epoch ms|second", _
          DateKeyOf(Histories("{'id':'1','first':1761134538011,'second':1761134538011}"))
    Check "datekey", "several: a name that hints at a date", "zzz, createdOn", "createdOn|epoch ms|zzz", _
          DateKeyOf(Histories("{'zzz':1761134538011,'createdOn':1761134538011}"))
    Check "datekey", "several: another hint", "zzz, eventDate", "eventDate|epoch ms|zzz", _
          DateKeyOf(Histories("{'zzz':1761134538011,'eventDate':1761134538011}"))
    Check "datekey", "an id of ten digits in the past reads as a date: created wins", "id 1700000001", "created|ISO|id", _
          DateKeyOf(Histories("{'id':'1700000001','created':'2026-10-01T09:12:33Z'}"))
    Check "datekey", "an id in the future is not a date", "id 2000000001", "created|ISO|", _
          DateKeyOf(Histories("{'id':'2000000001','created':'2026-10-01T09:12:33Z'}"))
    Check "datekey", "one odd history rejects the key", "created, other", "other|epoch ms|", _
          DateKeyOf(Histories("{'created':1761134538011,'other':1761134538011}", "{'created':'not a date','other':1761134538011}"))
    Check "datekey", "a key missing from one history is rejected", "created, other", "other|epoch ms|", _
          DateKeyOf(Histories("{'created':1761134538011,'other':1761134538011}", "{'other':1761134538011}"))
    Check "datekey", "no date: created and none", "created not a date", "created|none|", DateKeyOf(Histories("{'id':'1','created':'not a date'}"))
    Check "datekey", "no history: created and none", "empty", "created|none|", DateKeyOf(New Collection)
    Check "datekey", "a date inside an object is not read", "created {epochMillis}", "created|none|", _
          DateKeyOf(Histories("{'id':'1','created':{'epochMillis':1761134538011}}"))
    Set hs = New Collection
    For i = 1 To DATE_SAMPLE
        hs.Add JsonParse(J("{'created':1761134538011}"))
    Next i
    hs.Add JsonParse(J("{'created':'not a date'}"))
    Check "datekey", "only the first histories are read", "20 good then 1 odd", "created|epoch ms|", DateKeyOf(hs)
End Sub

Public Sub T_ChangelogDateKey()
    ' The changelog of the run is read with the key found by value, and a new run finds it again: a renamed key, then the usual one.
    Dim cfg As Object, keys As New Collection, rows As Collection, url As String, map As Object, out As Collection, skipped As Long, shapes As String
    Set cfg = TestConfig()
    keys.Add "K-1"
    url = cfg("BaseUrl") & "/rest/api/3/changelog/bulkfetch"
    Set map = CreateObject("Scripting.Dictionary")
    FakeReset
    FakeAdd "POST", url, 200, J("{'issueChangeLogs':[{'issueId':'10001','changeHistories':[{'id':'1','createdAt':1761134538011," & _
        "'items':[{'field':'status','fieldId':'status','fieldtype':'jira','from':'1','fromString':'Nouveau','to':'3','toString':'En analyse'}]}]}]}")
    ResetDateKey
    Set rows = JiraChangelogsAll(cfg, keys)
    Check "datekey", "the date of a renamed key is in the row", "createdAt 1761134538011", 1761134538011#, CDbl(rows(1)(2))
    Set out = StatusChangesFromChangelog(rows, map, skipped, shapes)
    Check "datekey", "the status change is not left out", "0 skipped", 0, skipped
    FakeReset
    FakeAdd "POST", url, 200, J("{'issueChangeLogs':[{'issueId':'10001','changeHistories':[{'id':'1','created':'2026-10-02T10:00:00.000+0200'," & _
        "'items':[{'field':'status','fieldId':'status','fieldtype':'jira','from':'1','fromString':'Nouveau','to':'3','toString':'En analyse'}]}]}]}")
    Set rows = JiraChangelogsAll(cfg, keys)
    Check "datekey", "the key is found again when the first history lacks the key kept", "created ISO", "2026-10-02T10:00:00.000+0200", CStr(rows(1)(2))
    FakeReset
    FakeAdd "POST", url, 200, J("{'issueChangeLogs':[{'issueId':'10001','changeHistories':[{'id':'1','created':'not a date'," & _
        "'items':[{'field':'status','fieldId':'status','fieldtype':'jira','from':'1','fromString':'Nouveau','to':'3','toString':'En analyse'}]}]}]}")
    Set rows = JiraChangelogsAll(cfg, keys)
    Check "datekey", "no date anywhere: created, as before", "not a date", "not a date", CStr(rows(1)(2))
    ResetDateKey
End Sub

Public Sub T_ParseJiraDateEpoch()
    ' The changelog of a real Jira dates its histories with a NUMBER of milliseconds since 1970 UTC (1761134538011): it was read as
    ' no date, so every status change was left out. A number of seconds is read too; whole seconds; other numbers are no date.
    Dim want As Double
    want = CDbl(DateSerial(2025, 10, 22) + TimeSerial(12, 2, 18))
    Check "epoch", "milliseconds, as a JSON number (Double)", "1761134538011", want, CDbl(ParseJiraDate(1761134538011#))
    Check "epoch", "the same as text", "1761134538011", want, CDbl(ParseJiraDate("1761134538011"))
    Check "epoch", "seconds", "1761134538", want, CDbl(ParseJiraDate(1761134538#))
    Check "epoch", "the milliseconds are dropped, as in the text form", "...539999", want, CDbl(ParseJiraDate(1761134538999#))
    Check "epoch", "a number too small is no date", "20261001", True, IsEmpty(ParseJiraDate(20261001#)) And IsEmpty(ParseJiraDate("20261001"))
    Check "epoch", "a number too big is no date", "10^15", True, IsEmpty(ParseJiraDate(1E+15))
    Check "epoch", "an ISO date is still read", "2026-10-01", CDbl(DateSerial(2026, 10, 1)), CDbl(ParseJiraDate("2026-10-01"))
End Sub

Public Sub T_DateShape()
    ' No content in the shape: the log must not hold what Jira answered.
    Check "flow", "an ISO date", "2026-10-01T09:12:33.123+0200", "String:9999-99-99A99:99:99.999+9999", DateShape("2026-10-01T09:12:33.123+0200")
    Check "flow", "letters are masked", "Hello World", "String:Aaaaa Aaaaa", DateShape("Hello World")
    Check "flow", "a number", "15", "Integer:99", DateShape(15)               ' not 1.5: its text follows the decimal mark of the machine
    Check "flow", "Empty and Null", "types", "Empty|Null", DateShape(Empty) & "|" & DateShape(Null)
    Check "flow", "an object", "Dictionary", "Object", DateShape(CreateObject("Scripting.Dictionary"))
    Check "flow", "non ASCII", "e-acute", "String:a?", DateShape("e" & ChrW(233))
    Check "flow", "cut at 40 characters", "50 digits", "String:" & String$(40, "9") & "...", DateShape(String$(50, "7"))
End Sub

Public Sub T_SkippedDatesWarning()
    ' J2D-DATA-003 in the log: only when something was left out, with the count and the shapes.
    Dim file As String, saved As String, text As String
    saved = gLogPath
    file = Environ$("TEMP") & "\j2d_skipped_test_" & Format$(Now, "hhnnss") & Int(Rnd() * 1000) & ".log"
    gLogPath = file
    LogSkippedDates 0, ""
    Check "log", "nothing skipped, no line", "0", False, Len(Dir(file)) > 0
    LogSkippedDates 3, "String:9999999999999 x 3"
    text = FileText(file)
    Check "log", "the warning with its code, count and shapes", "3 x epoch", True, _
          text Like "####-##-## ##:##:##.### WARN J2D-DATA-003 | 3 status changes ignored: their date cannot be read | created: String:9999999999999 x 3" & vbCrLf
    gLogPath = saved
    Kill file
End Sub

' ================= Propose values from Jira =================

Private Sub PutRows(ByVal wb As Object, ByVal logicalName As String, ByVal rows As Variant)
    ' A sheet of wb (named like the real one) from rows: Array(Array(header...), Array(row...), ...); "" is an empty cell.
    Dim ws As Object, r As Long, c As Long
    If SheetExists(wb, logicalName) Then
        Set ws = wb.Worksheets(SheetName(logicalName))                       ' written again: the new rows replace the old ones
        ws.Cells.Clear
    Else
        Set ws = wb.Worksheets.Add(, wb.Worksheets(wb.Worksheets.Count))
        ws.Name = SheetName(logicalName)
    End If
    For r = 0 To UBound(rows)
        For c = 0 To UBound(rows(r))
            If Len(CStr(rows(r)(c))) > 0 Then ws.Cells(r + 1, c + 1).Value = rows(r)(c)
        Next c
    Next r
End Sub

Private Function ProposalBook() As Object
    ' A scratch workbook with the parameter sheets (what the person has typed) and the proposal sheets (what Jira gave).
    Dim wb As Object
    Set wb = Workbooks.Add
    PutRows wb, "Projects", Array(Array("Project", "Role", "Team"), Array("PRG", "Program", "ART"), Array("TA", "Team", "Team A"), _
                                  Array("TB", "Team", "Team B"))
    PutRows wb, "Workflows", Array(Array("Workflow", "Status", "Category", "IsStart", "IsEnd", "IsFree"), _
                                   Array("Team", "Open", "Backlog", 1, "", 0), Array("Team", "Doing", "", 0, "", 0))
    PutRows wb, "Transitions", Array(Array("Workflow", "From", "To"), Array("Team", "Open", "Doing"))
    PutRows wb, "Config", Array(Array("Parameter", "Value", "Example", "Mandatory", "Description"), Array("SprintField", "", "x", "False", ""), _
                                Array("TeamField", "customfield_9", "x", "False", ""), Array("StoryPointsField", "", "x", "False", ""), _
                                Array("PIField", "", "x", "False", ""), Array("BusinessValueField", "", "x", "False", ""), _
                                Array("TargetStartField", "", "x", "False", ""))
    PutRows wb, "JiraWorkflows", Array(Array("Workflow", "Status", "Category", "IsStart", "IsEnd", "IsFree", "JiraCategory", "Initial", "Project"), _
                                       Array("WF-A", "Open", "Backlog", "", "", 0, "TODO", 1, "PRG, TA"), _
                                       Array("WF-A", "Doing", "", "", "", 0, "IN_PROGRESS", "", "PRG, TA"), _
                                       Array("WF-A", "Closed", "Done", "", "", 0, "DONE", "", "PRG, TA"), _
                                       Array("WF-F", "Draft", "Backlog", "", "", 1, "TODO", "", "PRG"))
    PutRows wb, "JiraStatuses", Array(Array("Project", "IssueType", "Status", "JiraCategory"), Array("TB", "Story", "Triage", "new"), _
                                      Array("TB", "Story", "Open", "new"))
    PutRows wb, "JiraTransitions", Array(Array("Workflow", "From", "To", "JiraTransition", "TransitionId", "Scope", "Project"), _
                                         Array("WF-A", "Open", "Doing", "Start", "11", "DIRECTED", "PRG, TA"), _
                                         Array("WF-A", "Doing", "Closed", "Close", "12", "DIRECTED", "PRG, TA"), _
                                         Array("WF-F", "Draft", "Ready", "Go", "13", "DIRECTED", "PRG"))
    PutRows wb, "JiraFields", Array(Array("Id", "Name", "Custom", "SchemaType", "SchemaDetail"), _
                                    Array("customfield_1", "Sprint", "Yes", "array", "com.pyxis.greenhopper.jira:gh-sprint"), _
                                    Array("customfield_2", "Team", "Yes", "team", "com.atlassian.teams:rm-teams-custom-field-team"), _
                                    Array("customfield_3", "Story Points", "Yes", "number", ""), _
                                    Array("customfield_4", "PI", "Yes", "option", ""), _
                                    Array("customfield_5", "Business Value", "Yes", "number", ""), _
                                    Array("customfield_6", "business value", "Yes", "number", ""))
    PutRows wb, "Iterations", Array(Array("Iteration-label", "Start", "End"), Array("I0", "2026-01-01", "2026-01-15"))
    PutRows wb, "JiraSprints", Array(Array("Id", "Project", "Board", "Sprint", "State", "Start", "End"), _
        Array("1", "TA", "Board TA", "A Sprint 1", "closed", "2026-01-01T09:00:00.000+0000", "2026-01-14T17:00:00.000+0000"), _
        Array("2", "TA", "Board TA", "A Sprint 2", "closed", "2026-01-15T09:00:00.000+0000", "2026-01-28T17:00:00.000+0000"), _
        Array("3", "TB", "Board TB", "B Sprint 2", "closed", "2026-01-16T09:00:00.000+0000", "2026-01-28T17:00:00.000+0000"), _
        Array("4", "TA", "Board TA", "Sprint 3", "closed", "2026-01-29T09:00:00.000+0000", "2026-02-12T09:00:00.000+0000"), _
        Array("5", "TB", "Board TB", "Sprint 3", "closed", "2026-01-29T09:00:00.000+0000", "2026-02-12T09:00:00.000+0000"), _
        Array("6", "TA", "Board TA", "Sprint 4", "active", "2026-02-12T09:00:00.000+0000", "2026-02-25T17:00:00.000+0000"), _
        Array("7", "TA", "Board TA", "Sprint 5", "future", "", ""), _
        Array("8", "T1", "Board T1", "T1 Sprint 9", "future", "2026-02-26T09:00:00.000+0000", "2026-03-11T17:00:00.000+0000"), _
        Array("9", "T2", "Board T2", "T2 Sprint 9", "future", "2026-02-26T09:00:00.000+0000", "2026-03-18T17:00:00.000+0000"), _
        Array("10", "T3", "Board T3", "T3 Sprint 9", "future", "2026-02-26T09:00:00.000+0000", "2026-03-11T17:00:00.000+0000"), _
        Array("11", "T4", "Board T4", "Broken", "future", "2026-04-10T09:00:00.000+0000", "2026-04-01T17:00:00.000+0000"))
    Application.DisplayAlerts = False
    wb.Worksheets(1).Delete
    Application.DisplayAlerts = True
    Set ProposalBook = wb
End Function

Public Sub T_ProposeFromSheets()
    ' Only EMPTY cells get a proposal, each one yellow, the first of a row with a note; what was typed is never changed.
    Dim wb As Object, wsW As Object, wsT As Object, wsC As Object, summary As String, again As String, yellow As Long
    yellow = RGB(255, 242, 204)
    Set wb = ProposalBook()
    summary = ProposeFromSheets(wb)
    Set wsW = wb.Worksheets(SheetName("Workflows")): Set wsT = wb.Worksheets(SheetName("Transitions")): Set wsC = wb.Worksheets(SheetName("Config"))
    Check "propose", "the summary names Config with its count and no longer the sheets Workflows and Transitions", "3", True, _
          InStr(summary, " | Config_from_user: 3 proposed") > 0 And InStr(summary, "Workflows_from_user") = 0 And InStr(summary, "Transitions_from_user") = 0
    Check "propose", "the sheet Workflows is left as typed: Jira's statuses are read by the metrics, not copied", "rows", "3|Team|Open|Backlog|1||0", _
          LastRow(wsW) & "|" & RowText(wsW, 2, 6)
    Check "propose", "the sheet Transitions is left as typed", "rows", 2, LastRow(wsT)
    Check "propose", "an empty field id gets the one field that fits (by its custom type)", "SprintField", "customfield_1", CStr(wsC.Cells(2, 2).Value)
    Check "propose", "a typed field id is never changed", "TeamField", "customfield_9", CStr(wsC.Cells(3, 2).Value)
    Check "propose", "by the name when the custom type says nothing", "StoryPointsField", "customfield_3", CStr(wsC.Cells(4, 2).Value)
    Check "propose", "the PI field by its name", "PIField", "customfield_4", CStr(wsC.Cells(5, 2).Value)
    Check "propose", "two fields fit: nothing, the person chooses", "BusinessValueField", "", CStr(wsC.Cells(6, 2).Value & "")
    Check "propose", "no field fits: nothing", "TargetStartField", "", CStr(wsC.Cells(7, 2).Value & "")
    Check "propose", "a proposed id is yellow with a note", "SprintField", yellow & "|To confirm: proposed from Jira", wsC.Cells(2, 2).Interior.Color & "|" & wsC.Cells(2, 2).Comment.Text
    again = ProposeFromSheets(wb)
    Check "propose", "a second run adds nothing", "nothing new", True, InStr(again, " | Proposed: nothing new") = 1
    Application.DisplayAlerts = False
    wb.Close False
    Application.DisplayAlerts = True
End Sub

Public Sub T_ProposeIterations()
    ' The sprints of the teams that start and end on the same days (+- 1 day) are one iteration; the sheet gets only the ones it lacks.
    Dim wb As Object, ws As Object, summary As String, yellow As Long
    yellow = RGB(255, 242, 204)
    Set wb = ProposalBook()
    summary = ProposeFromSheets(wb)
    Set ws = wb.Worksheets(SheetName("Iterations"))
    Check "propose", "the iteration the sheet has is left as typed", "I0", "I0", CStr(ws.Cells(2, 1).Value)
    Check "propose", "two teams, two names, the same days: one iteration, named by its start", "Iteration 2026-01-15", _
          "Iteration 2026-01-15|2026-01-15|2026-01-29", RowText(ws, 3, 3)
    Check "propose", "the shared name is the label; the end is cut at the start of the next iteration (no day twice)", "Sprint 3", "Sprint 3|2026-01-29|2026-02-12", RowText(ws, 4, 3)
    Check "propose", "the end is the day after the last day", "Sprint 4", "Sprint 4|2026-02-12|2026-02-26", RowText(ws, 5, 3)
    Check "propose", "three teams that start the same day with three lengths are one iteration, as long as the longest", "T1/T2/T3", _
          "Iteration 2026-02-26|2026-02-26|2026-03-19", RowText(ws, 6, 3)
    Check "propose", "a sprint inside an iteration the sheet has, a sprint without dates, a sprint that ends before it starts: nothing", "rows", 6, LastRow(ws)
    Check "propose", "the summary counts them", "4", True, InStr(summary, " | Iterations_from_user: 4 proposed") > 0
    Check "propose", "a proposed iteration is yellow with a note", "label", yellow & "|To confirm: proposed from Jira", ws.Cells(3, 1).Interior.Color & "|" & ws.Cells(3, 1).Comment.Text
    Check "propose", "its dates are ISO text, as the sheet reads them", "Start", "2026-01-15", CStr(ws.Cells(3, 2).Value)
    Application.DisplayAlerts = False
    wb.Close False
    Application.DisplayAlerts = True
End Sub

Public Sub T_UserFieldIds()
    ' The fields that hold a person come from the schema Jira gives, not from a name: a single user, a list of users, a custom picker.
    Dim cfg As Object, ids As Object, base As String
    Set cfg = TestConfig(): base = cfg("BaseUrl")
    FakeReset
    FakeAdd "GET", base & "/rest/api/3/field", 200, J("[{'id':'assignee','schema':{'type':'user'}},{'id':'customfield_10050','schema':{'type':'user','custom':'x'}}," & _
        "{'id':'customfield_10051','schema':{'type':'array','items':'user'}},{'id':'summary','schema':{'type':'string'}},{'id':'status'}]")
    Set ids = UserFieldIds(cfg)
    Check "userfields", "a user, a custom user and a list of users", "3", "assignee|customfield_10050|customfield_10051", _
          Join(ids.Keys, "|")
    FakeReset
    FakeAdd "GET", base & "/rest/api/3/field", 403, J("{'errorMessages':['no permission']}")
    Check "userfields", "a call that fails gives none, never an error", "403", 0, UserFieldIds(cfg).Count
End Sub

Public Sub T_ReadSprints()
    ' The sprints of the scrum boards of the projects: two boards of a project, paged; a sprint shared by two boards comes once;
    ' a Kanban board (400) and a project whose boards the account cannot read (403) are skipped, never an error.
    Dim cfg As Object, keys As New Collection, base As String, rows As Collection, noBoard As Long
    Set cfg = TestConfig(): base = cfg("BaseUrl")
    keys.Add "TA": keys.Add "ZZ": keys.Add "TK"
    FakeReset
    FakeAdd "GET", BoardsUrl(base, "TA", 0), 200, J("{'isLast':false,'values':[{'id':7,'name':'Board TA'}]}")
    FakeAdd "GET", BoardsUrl(base, "TA", 1), 200, J("{'isLast':true,'values':[{'id':8,'name':'Board TA kanban'}]}")
    FakeAdd "GET", SprintsUrl(base, "7", 0), 200, J("{'isLast':false,'values':[{'id':1,'name':'S1','state':'closed','startDate':'2026-01-01T09:00:00.000+0000'," & _
        "'endDate':'2026-01-14T17:00:00.000+0000'}]}")
    FakeAdd "GET", SprintsUrl(base, "7", 1), 200, J("{'isLast':true,'values':[{'id':2,'name':'S2','state':'future'}]}")
    FakeAdd "GET", SprintsUrl(base, "8", 0), 400, J("{'errorMessages':['The board does not support sprints']}")
    FakeAdd "GET", BoardsUrl(base, "ZZ", 0), 403, J("{'errorMessages':['no permission']}")
    FakeAdd "GET", BoardsUrl(base, "TK", 0), 200, J("{'isLast':true,'values':[{'id':7,'name':'Board TA again'}]}")
    FakeAdd "GET", SprintsUrl(base, "7", 0), 200, J("{'isLast':true,'values':[{'id':1,'name':'S1','state':'closed'}]}")
    Set rows = ReadSprints(cfg, keys, noBoard)
    Check "sprints", "the sprints of a board, page by page", "S1 then S2", "1|TA|Board TA|S1|closed|2026-01-01T09:00:00.000+0000|2026-01-14T17:00:00.000+0000;2|TA|Board TA|S2|future||", _
          Join(Array(Join(rows(1), "|"), Join(rows(2), "|")), ";")
    Check "sprints", "a sprint shared by two boards comes once", "S1", 2, rows.Count
    Check "sprints", "a Kanban board and an unreadable project are skipped", "ZZ", 1, noBoard
End Sub

Public Sub T_TeamNames()
    ' The teams of the charts: ART, then the team of each Team project, once; an empty Team is the project key (AL, 2026-11-18).
    Dim ws As Object, rows As New Collection, names As Collection
    Set ws = ThisWorkbook.Worksheets("TestScratch")
    rows.Add Array("PRG", "Program", "")
    rows.Add Array("TA", "Team", "Team A")
    rows.Add Array("TB", "Team", "")
    rows.Add Array("TC", "Team", "Team A")
    rows.Add Array("TD", "Team", " ")
    WriteTable ws, Array("Project", "Role", "Team"), rows
    Set names = TeamNames(ws, "ART")
    Check "teams", "ART, the team of a project, the key of a project without team (a blank cell too), each once", "PRG TA TB TC TD", "4|ART|Team A|TB|TD", _
          names.Count & JoinNames(names)
    ws.Cells.Clear
End Sub

Public Sub T_ReadStatusCatalog()
    ' Every status of the instance with its Jira category, paged; a refused or unreadable call ends the reading and never stops the run.
    Dim cfg As Object, base As String, rows As Collection, wb As Object, ws As Object
    Set cfg = TestConfig(): base = cfg("BaseUrl")
    FakeReset
    FakeAdd "GET", base & "/rest/api/3/statuses/search?startAt=0&maxResults=200", 200, J("{'isLast':false,'values':[{'id':'1','name':'Open','statusCategory':'TODO'}," & _
        "{'id':'2','name':'Doing','statusCategory':{'key':'indeterminate'}},{'id':'3','name':'','statusCategory':'DONE'}]}")
    FakeAdd "GET", base & "/rest/api/3/statuses/search?startAt=3&maxResults=200", 200, J("{'isLast':true,'values':[{'id':'4','name':'Closed','statusCategory':'DONE'}]}")
    Set rows = ReadStatusCatalog(cfg)
    Check "catalog", "the statuses of the pages, a status without a name left out, the category as an enum or as an object", "3", _
          "1|Open|TODO;2|Doing|indeterminate;4|Closed|DONE", Join(Array(Join(rows(1), "|"), Join(rows(2), "|"), Join(rows(3), "|")), ";") & IIf(rows.Count = 3, "", "?")
    FakeReset
    FakeAdd "GET", base & "/rest/api/3/statuses/search?startAt=0&maxResults=200", 403, J("{'errorMessages':['no permission']}")
    Check "catalog", "a refused call gives none", "403", 0, ReadStatusCatalog(cfg).Count
    FakeReset
    FakeAdd "GET", base & "/rest/api/3/statuses/search?startAt=0&maxResults=200", 200, "not json"
    Check "catalog", "an answer that is not JSON gives none", "text", 0, ReadStatusCatalog(cfg).Count
    Set wb = Workbooks.Add
    Set rows = New Collection
    WriteStatusCatalog wb, rows
    Check "catalog", "nothing is written when the call gave nothing", "no sheet", False, SheetExists(wb, "JiraStatusCatalog")
    rows.Add Array("1", "Open", "TODO")
    WriteStatusCatalog wb, rows
    Set ws = wb.Worksheets(SheetName("JiraStatusCatalog"))
    Check "catalog", "the proposal sheet", "Id|Status|JiraCategory / 1|Open|TODO", "Id|Status|JiraCategory/1|Open|TODO", RowText(ws, 1, 3) & "/" & RowText(ws, 2, 3)
    Application.DisplayAlerts = False
    wb.Close False
    Application.DisplayAlerts = True
End Sub

Public Sub T_ReadWorkflowRules()
    ' The starts and the free statuses of the workflows of the projects, without any sheet: the first status in progress from the creation, the GLOBAL ones.
    Dim cfg As Object, keys As New Collection, base As String, starts As Object, frees As Object, wb As Object
    Set cfg = TestConfig(): base = cfg("BaseUrl")
    keys.Add "PRG"
    Set starts = Dict(): Set frees = Dict()
    FakeReset
    FakeAdd "GET", ProjectSearchUrl(base, keys, 1, 0), 200, J("{'startAt':0,'isLast':true,'values':[{'id':'10100','key':'PRG','style':'classic'," & _
        "'issueTypes':[{'id':'10001','name':'Feature','subtask':false,'hierarchyLevel':1}]}]}")
    FakeAdd "POST", base & "/rest/api/3/workflows/preview", 200, J("{'statuses':[{'statusReference':'r1','name':'Open','statusCategory':'TODO'}," & _
        "{'statusReference':'r2','name':'Doing','statusCategory':'IN_PROGRESS'},{'statusReference':'r3','name':'Closed','statusCategory':'DONE'}," & _
        "{'statusReference':'r4','name':'Parked','statusCategory':'TODO'}]," & _
        "'workflows':[{'name':'WF-A','queryContext':[{'project':'P','issueTypes':['10001']}],'statuses':[{'statusReference':'r1'},{'statusReference':'r2'}," & _
        "{'statusReference':'r3'},{'statusReference':'r4'}],'transitions':[{'type':'INITIAL','name':'Create','toStatusReference':'r1','links':[]}," & _
        "{'type':'DIRECTED','name':'Start','id':'11','toStatusReference':'r2','links':[{'fromStatusReference':'r1'}]}," & _
        "{'type':'DIRECTED','name':'Close','toStatusReference':'r3','links':[{'fromStatusReference':'r2'}]}," & _
        "{'type':'GLOBAL','name':'Park','toStatusReference':'r4','links':[]}]}]}")
    Set wb = Workbooks.Add
    Check "rules", "every project read: nothing refused", "0", 0, ReadWorkflowRules(cfg, keys, starts, frees, wb)
    Check "rules", "the start and the free status of the workflows of the projects", "Doing | Parked", "Doing|Parked", Join(starts.Keys, ",") & "|" & Join(frees.Keys, ",")
    Check "rules", "the workflow of each issue type is kept in a sheet", "PRG|Feature|WF-A", "PRG|Feature|WF-A", RowText(wb.Worksheets(SheetName("JiraIssueTypes")), 2, 3)
    Check "rules", "the creation status of the workflow is marked, the moves are kept", "Open", "Open|Start", _
          IIf(Val(CStr(wb.Worksheets(SheetName("JiraWorkflows")).Cells(2, 8).Value)) = 1, CStr(wb.Worksheets(SheetName("JiraWorkflows")).Cells(2, 2).Value), "?") & "|" & _
          CStr(wb.Worksheets(SheetName("JiraTransitions")).Cells(2, 4).Value)
    Application.DisplayAlerts = False
    wb.Close False
    Application.DisplayAlerts = True
    FakeReset
    Set starts = Dict(): Set frees = Dict()
    Set wb = Workbooks.Add
    Dim failCode As String
    Check "rules", "a reading that fails is told to the caller, with the code of its error", "-1 J2D-JIRA-404", "-1 J2D-JIRA-404", _
          ReadWorkflowRules(cfg, keys, starts, frees, wb, failCode) & " " & failCode
    Check "rules", "a call that fails gives nothing, writes no sheet and never raises", "404", 0, starts.Count + frees.Count + IIf(SheetExists(wb, "JiraWorkflows"), 1, 0)
    FakeAdd "GET", ProjectSearchUrl(base, keys, 1, 0), 200, J("{'startAt':0,'isLast':true,'values':[{'id':'10100','key':'PRG','style':'classic'," & _
        "'issueTypes':[{'id':'10001','name':'Feature','subtask':false,'hierarchyLevel':1}]}]}")
    FakeAdd "POST", base & "/rest/api/3/workflows/preview", 403, J("{'errorMessages':['no permission']}")
    Check "rules", "a project whose workflows Jira refuses (no administrator right) is counted", "1", 1, ReadWorkflowRules(cfg, keys, starts, frees, wb)
    Application.DisplayAlerts = False
    wb.Close False
    Application.DisplayAlerts = True
End Sub

Public Sub T_ReadLinkTypes()
    ' The issue link types of the site: Name | Inward | Outward; an account that may not read them gives none, never an error.
    Dim cfg As Object, base As String, rows As Collection
    Set cfg = TestConfig(): base = cfg("BaseUrl")
    FakeReset
    FakeAdd "GET", base & "/rest/api/3/issueLinkType", 200, J("{'issueLinkTypes':[{'id':'1','name':'Blocks','inward':'is blocked by','outward':'blocks'}," & _
        "{'id':'2','name':'Relates','inward':'relates to','outward':'relates to'}]}")
    Set rows = ReadLinkTypes(cfg)
    Check "linktypes", "the types of the site", "2", "Blocks|is blocked by|blocks;Relates|relates to|relates to", Join(Array(Join(rows(1), "|"), Join(rows(2), "|")), ";")
    FakeReset
    FakeAdd "GET", base & "/rest/api/3/issueLinkType", 403, J("{'errorMessages':['no permission']}")
    Check "linktypes", "a call that is refused gives none, never an error", "403", 0, ReadLinkTypes(cfg).Count
    FakeReset
    FakeAdd "GET", base & "/rest/api/3/issueLinkType", 500, J("{'errorMessages':['server error']}")
    Check "linktypes", "a server error gives none too: it never stops the button", "500", 0, ReadLinkTypes(cfg).Count
    FakeReset
    FakeAdd "GET", base & "/rest/api/3/issueLinkType", 200, "not json"
    Check "linktypes", "an answer that is not JSON gives none", "text", 0, ReadLinkTypes(cfg).Count
End Sub

Private Function LinkTypesBook(ByVal pointsValue As String, ByVal blocksValue As String) As Object
    ' The Config rows (Value = Example: the delivered default) and what Jira says about the issue types and the link types.
    Dim wb As Object
    Set wb = Workbooks.Add
    PutRows wb, "Projects", Array(Array("Project", "Role", "Team"), Array("PRG", "Program", "ART"), Array("TA", "Team", "Team A"))
    PutRows wb, "Config", Array(Array("Parameter", "Value", "Example", "Mandatory", "Description"), _
        Array("SubTaskTypes", "Sub-task", "Sub-task", "False", ""), Array("FeatureTypes", "Feature", "Feature", "True", ""), _
        Array("ParentRequiredTypes", "Story,Task", "Story,Task", "False", ""), Array("PointsRequiredTypes", pointsValue, "Story", "False", ""), _
        Array("BlocksLinkType", blocksValue, "Blocks", "True", ""))
    PutRows wb, "JiraIssueTypes", Array(Array("Project", "IssueType", "Workflow", "Subtask", "HierarchyLevel", "ProjectStyle"), _
        Array("PRG", "Epic", "WF", "No", 1, "classic"), Array("PRG", "Feature", "WF", "No", 2, "classic"), Array("PRG", "Enabler", "WF", "No", 2, "classic"), _
        Array("TA", "Story", "WF", "No", 0, "classic"), Array("TA", "Bug", "WF", "No", 0, "classic"), _
        Array("TA", "Sub-task", "WF", "Yes", -1, "classic"), Array("TA", "Technical task", "WF", "Yes", -1, "classic"), _
        Array("TB", "Task", "WF", "No", 0, "classic"))
    PutRows wb, "JiraLinkTypes", Array(Array("Name", "Inward", "Outward"), Array("Blocks", "is blocked by", "blocks"), Array("Relates", "relates to", "relates to"))
    Application.DisplayAlerts = False
    wb.Worksheets(1).Delete
    Application.DisplayAlerts = True
    Set LinkTypesBook = wb
End Function

Public Sub T_ProposeTypes()
    ' Config: the lists of issue types come from JiraIssueTypes, BlocksLinkType from JiraLinkTypes, where the Value is empty or still the
    ' delivered default (= its Example); a typed value stays; yellow with a note.
    Dim wb As Object, wsC As Object, summary As String, yellow As Long
    yellow = RGB(255, 242, 204)
    Set wb = LinkTypesBook("Story", "Blocks")
    summary = ProposeFromSheets(wb)
    Set wsC = wb.Worksheets(SheetName("Config"))
    Check "propose", "the sub-tasks of Jira replace the default", "SubTaskTypes", "Sub-task,Technical task", CStr(wsC.Cells(2, 2).Value)
    Check "propose", "the types of the Program at its highest level replace the default", "FeatureTypes", "Feature,Enabler", CStr(wsC.Cells(3, 2).Value)
    Check "propose", "a type that no Team project has is taken out of the rule", "ParentRequiredTypes", "Story", CStr(wsC.Cells(4, 2).Value)
    Check "propose", "Jira agrees with the default: nothing written", "PointsRequiredTypes", "Story|" & -4142, CStr(wsC.Cells(5, 2).Value) & "|" & wsC.Cells(5, 2).Interior.ColorIndex
    Check "propose", "the link type that blocks agrees with the default: nothing written", "BlocksLinkType", "Blocks|" & -4142, CStr(wsC.Cells(6, 2).Value) & "|" & wsC.Cells(6, 2).Interior.ColorIndex
    Check "propose", "a replaced default is yellow with the note", "SubTaskTypes", yellow & "|To confirm: proposed from Jira", wsC.Cells(2, 2).Interior.Color & "|" & wsC.Cells(2, 2).Comment.Text
    Check "propose", "the summary counts the three", "3", True, InStr(summary, " | Config_from_user: 3 proposed") > 0
    Check "propose", "a second run adds nothing", "nothing new", True, InStr(ProposeFromSheets(wb), " | Proposed: nothing new") = 1
    Application.DisplayAlerts = False
    wb.Close False
    Application.DisplayAlerts = True
    ' a value the person typed stays, the link type of the site is another one
    Set wb = LinkTypesBook("Story,Task", "Bloque")
    PutRows wb, "JiraLinkTypes", Array(Array("Name", "Inward", "Outward"), Array("Blocks", "is blocked by", "blocks"))
    ProposeFromSheets wb
    Set wsC = wb.Worksheets(SheetName("Config"))
    Check "propose", "a typed link type stays, whatever Jira says", "BlocksLinkType", "Bloque", CStr(wsC.Cells(6, 2).Value)
    Application.DisplayAlerts = False
    wb.Close False
    Application.DisplayAlerts = True
    ' two link types that block: nothing, the person chooses; the default link type is replaced by the one Jira names once
    Set wb = LinkTypesBook("Story", "Blocks")
    PutRows wb, "JiraLinkTypes", Array(Array("Name", "Inward", "Outward"), Array("Blocks", "is blocked by", "blocks"), Array("Stops", "is stopped by", "blocks"))
    ProposeFromSheets wb
    Check "propose", "two link types that block: nothing proposed", "BlocksLinkType", "Blocks", CStr(wb.Worksheets(SheetName("Config")).Cells(6, 2).Value)
    PutRows wb, "JiraLinkTypes", Array(Array("Name", "Inward", "Outward"), Array("Bloquant", "est bloque par", "blocks"))
    ProposeFromSheets wb
    Check "propose", "the one link type that blocks replaces the default name", "BlocksLinkType", "Bloquant", CStr(wb.Worksheets(SheetName("Config")).Cells(6, 2).Value)
    Application.DisplayAlerts = False
    wb.Close False
    Application.DisplayAlerts = True
End Sub

Public Sub T_ProposeValuesNotAdmin()
    ' IsUserProjectAdmin = False: no call to workflows/preview (it needs the project administrator right); the statuses of
    ' each issue type are still read (JiraStatuses), and the status says why the workflows were not read.
    Dim cfg As Object, keys As New Collection, base As String, wb As Object, summary As String, e As Variant, previews As Long
    Dim wsW As Object
    Set cfg = TestConfig(): base = cfg("BaseUrl")
    keys.Add "TB"
    Set wb = ProposalBook()
    FakeReset
    FakeAdd "GET", ProjectSearchUrl(base, keys, 1, 0), 200, J("{'isLast':true,'values':[{'id':'10102','key':'TB','style':'classic'," & _
        "'issueTypes':[{'id':'10002','name':'Story','subtask':false,'hierarchyLevel':0}]}]}")
    FakeAdd "GET", StatusesUrl(base, "TB"), 200, J("[{'id':'1','name':'Story','subtask':false,'statuses':[" & _
        "{'id':'5','name':'Doing','statusCategory':{'key':'indeterminate','name':'In progress'}}," & _
        "{'id':'6','name':'Shipped','statusCategory':{'key':'done','name':'Done'}}]}]")
    FakeAdd "GET", CreateMetaUrl(base, "TB", "10002", 0), 403, J("{'errorMessages':['no permission']}")
    FakeAdd "GET", JiraSearchUrl(base, ProjectsJql(ColOf("TB")), "", JIRA_PAGE_SIZE, "*all", "names"), 200, J("{'isLast':true,'names':{},'issues':[]}")
    FakeAdd "GET", BoardsUrl(base, "TB", 0), 200, J("{'isLast':true,'values':[]}")
    FakeAdd "GET", base & "/rest/api/3/field", 200, J("[{'id':'customfield_7','name':'Sprint','custom':true,'schema':{'type':'array'," & _
        "'custom':'com.pyxis.greenhopper.jira:gh-sprint'}}]")
    summary = ProposeValues(cfg, keys, wb, False)
    For Each e In gHttpLog
        If InStr(e(1), "/workflows/preview") > 0 Then previews = previews + 1
    Next e
    Check "propose", "no call to workflows/preview", "previews", 0, previews
    Check "propose", "no scrum board: the status says how many sprints and which projects", "TB", True, _
          InStr(summary, " | Sprints read: 0 (1 project(s) with no scrum board or sprint that you can read)") > 0
    Check "propose", "the status starts with its title and keeps the counts of the reading", "title", True, _
          InStr(summary, "Values proposed from Jira | Projects: 1 | Workflows: 0 | Fields: 1 | ") = 1
    Check "propose", "the status says why the workflows were not read", "IsUserProjectAdmin", True, _
          InStr(summary, " | Workflows not read: IsUserProjectAdmin is False") > 0
    Set wsW = wb.Worksheets(SheetName("Workflows"))
    Check "propose", "the sheet Workflows is left alone: no status copied from JiraStatuses", "rows", "3|Team|Doing||0||0", LastRow(wsW) & "|" & RowText(wsW, 3, 6)
    Check "propose", "the Sprint field is proposed too", "SprintField", "customfield_7", CStr(wb.Worksheets(SheetName("Config")).Cells(2, 2).Value)
    Application.DisplayAlerts = False
    wb.Close False
    Application.DisplayAlerts = True
End Sub

' ================= What to look at after an extraction =================

Private Function HealthBook(ByVal clean As Boolean) As Object
    ' A scratch workbook with the tables HealthFields reads: ItemFlow (Key, Team, DoneUtc, DC03, Workflow, StartUtc), StatusIntervals, VelocityCheck,
    ' the sheets Iterations, Config (AsOf) and StatusRules.
    Dim wb As Object, ws As Object, done1 As Double, outside As Double, t0 As Double
    Set wb = Workbooks.Add
    done1 = CDbl(DateSerial(2026, 1, 5)): outside = IIf(clean, CDbl(DateSerial(2026, 1, 9)), CDbl(DateSerial(2026, 3, 1)))
    PutRows wb, "Iterations", Array(Array("Iteration-label", "Start", "End"), Array("I1", "2026-01-01", "2026-01-15"))
    t0 = CDbl(DateSerial(2026, 1, 2))
    PutRows wb, "ItemFlow", Array(Array("Key", "Team", "DoneUtc", "DC03", "Workflow", "StartUtc"), _
                                  Array("K1", "Team A", done1, 0, "Team", IIf(clean, t0 + 1, "")), Array("K2", "Team A", outside, 0, "Team", ""), _
                                  Array("K3", IIf(clean, "Team A", ""), "", IIf(clean, 0, "=NA()"), "Team", ""), _
                                  Array("K4", "Team A", done1 + 2, IIf(clean, 0, "=NA()"), "Team", ""))
    ' K1 enters its start status (Weird) in the second of its backlog entry (not clean) or a day later (clean); K3 enters Odd after AsOf (not clean)
    ' EnteredUtc as dates, as the extraction writes them (a cell read back is a Date, not a number)
    PutRows wb, "StatusIntervals", Array(Array("Key", "Status", "Category", "EnteredUtc", "IsStart"), Array("K1", "Open", "Backlog", CDate(t0), 0), _
                                         Array("K1", "Weird", IIf(clean, "Active", "Unclassified"), CDate(IIf(clean, t0 + 1, t0 + 0.000003)), 1), _
                                         Array("K2", "Weird", IIf(clean, "Active", "Unclassified"), CDate(t0 + 1), 0), _
                                         Array("K3", "Odd", IIf(clean, "Done", "Unclassified"), IIf(clean, CDate(t0 + 5), DateSerial(2026, 3, 1)), 0))
    PutRows wb, "Issues", Array(Array("Key"), Array("K1"), Array("K2"), Array("K3"), Array("K4"))
    PutRows wb, "Config", Array(Array("Parameter", "Value", "Example", "Mandatory", "Description"), Array("AsOf", "2026-02-01", "", "True", ""))
    ' no start status and no Wait status in the rules (not clean)
    PutRows wb, "StatusRules", Array(Array("Workflow", "Status", "Category", "IsStart", "IsEnd", "IsFree", "Source"), _
                                     Array("Team", "Open", "Backlog", 0, 0, 0, "Jira"), Array("Team", "Weird", "Active", IIf(clean, 1, 0), 0, 0, "Jira"), _
                                     Array("Team", "Odd", IIf(clean, "Wait", "Done"), 0, IIf(clean, 0, 1), 0, "Jira"))
    PutRows wb, "VelocityCheck", Array(Array("Iteration-label", "Team", "Gap"), Array("I1", "Team A", 0), Array("I1", "ART", IIf(clean, 0, 2)))
    For Each ws In wb.Worksheets
        If ws.Name = SheetName("ItemFlow") Then ws.ListObjects.Add(1, ws.UsedRange, , 1).Name = "ItemFlow"
        If ws.Name = SheetName("StatusIntervals") Then ws.ListObjects.Add(1, ws.UsedRange, , 1).Name = "StatusIntervals"
        If ws.Name = SheetName("VelocityCheck") Then ws.ListObjects.Add(1, ws.UsedRange, , 1).Name = "VelocityCheck"
    Next ws
    Application.DisplayAlerts = False
    wb.Worksheets(1).Delete
    Application.DisplayAlerts = True
    Application.Calculate
    Set HealthBook = wb
End Function

Private Function ProgramBook(ByVal withFeatures As Boolean) As Object
    ' ItemFlow of a Program: Features (team ART) or none, and sub-tasks (team "ART (other types)"), a Config with FeatureTypes, the sheet Iterations.
    Dim wb As Object, ws As Object, done As Double
    Set wb = Workbooks.Add
    done = CDbl(DateSerial(2026, 1, 5))
    PutRows wb, "Iterations", Array(Array("Iteration-label", "Start", "End"), Array("I1", "2026-01-01", "2026-01-15"))
    PutRows wb, "Config", Array(Array("Parameter", "Value", "Example", "Mandatory", "Description"), Array("FeatureTypes", "Feature", "Feature", "True", ""))
    PutRows wb, "ItemFlow", Array(Array("Key", "Type", "Team", "DoneUtc"), _
        Array("P-1", IIf(withFeatures, "Feature", "Epic"), IIf(withFeatures, "ART", "ART (other types)"), done), _
        Array("P-2", "Sous-tache", "ART (other types)", done + 1), Array("P-3", "Sous-tache", "ART (other types)", CDbl(DateSerial(2026, 3, 1))), _
        Array("P-4", "Technique", "ART (other types)", ""))
    For Each ws In wb.Worksheets
        If ws.Name = SheetName("ItemFlow") Then ws.ListObjects.Add(1, ws.UsedRange, , 1).Name = "ItemFlow"
    Next ws
    Application.DisplayAlerts = False
    wb.Worksheets(1).Delete
    Application.DisplayAlerts = True
    Set ProgramBook = wb
End Function

Public Sub T_ReconciliationFindings()
    ' The status says how many controls of the sheet Reconciliation are not OK, and nothing when all are.
    Dim wb As Object, ws As Object
    Set wb = Workbooks.Add
    PutRows wb, "Reconciliation", Array(Array("Control", "Result"), Array("a", "OK"), Array("b", "REVIEW"), Array("c", "REVIEW"), Array("d", "INFO"))
    For Each ws In wb.Worksheets
        If ws.Name = SheetName("Reconciliation") Then ws.ListObjects.Add(1, ws.UsedRange, , 1).Name = "Reconciliation"
    Next ws
    Application.DisplayAlerts = False
    wb.Worksheets(1).Delete
    Application.DisplayAlerts = True
    Check "health", "two controls of four to review, the INFO one is not counted", "2 of 4", True, _
          InStr(HealthFields(wb), " | Reconciliation: 2 of 4 controls to review (sheet Reconciliation_from_excel)") > 0
    wb.Worksheets(SheetName("Reconciliation")).Range("B3:B4").Value = "OK"
    Check "health", "all OK: nothing is said", "(empty)", "", HealthFields(wb)
    Application.DisplayAlerts = False
    wb.Close False
    Application.DisplayAlerts = True
End Sub

Private Function Dict(ParamArray pairs() As Variant) As Object
    ' A Dictionary from key, value, key, value...
    Dim d As Object, i As Long
    Set d = CreateObject("Scripting.Dictionary")
    For i = LBound(pairs) To UBound(pairs) Step 2
        d(CStr(pairs(i))) = pairs(i + 1)
    Next i
    Set Dict = d
End Function

Private Function RulesText(ByVal rules As Collection) As String
    ' Each rule as kind|status|category|start|end|free|source, joined by ;
    Dim r As Variant
    For Each r In rules
        RulesText = RulesText & IIf(Len(RulesText) > 0, ";", "") & Join(Array(CStr(r(0)), CStr(r(1)), CStr(r(2)), CStr(r(3)), CStr(r(4)), CStr(r(5)), CStr(r(6))), "|")
    Next r
End Function

Private Function Moves(ParamArray pairs() As Variant) As Collection
    Dim out As New Collection, i As Long
    For i = LBound(pairs) To UBound(pairs) Step 2
        out.Add Array(CStr(pairs(i)), CStr(pairs(i + 1)))
    Next i
    Set Moves = out
End Function

Public Sub T_StatusRules()
    ' The same cases as tests/unit/test_status_rules.py: the two sides must agree.
    Dim cats As Object, none As Object, rules As Collection, ex As Collection, r As Variant, warns As Collection, w As Variant, s As String
    Set none = Dict()
    Check "rules", "the Jira categories in all their spellings", "new TODO indeterminate IN_PROGRESS done ?", "Backlog|Backlog|Active|Active|Done|", _
          JiraCategoryName("new") & "|" & JiraCategoryName("TODO") & "|" & JiraCategoryName("indeterminate") & "|" & JiraCategoryName("In Progress") & "|" & _
          JiraCategoryName("DONE") & "|" & JiraCategoryName("whatever")
    Set cats = Dict("Nouveau", "new", "Analyse", "indeterminate", "Pret", "indeterminate", "Dev", "indeterminate", "Fini", "done", "Idee", "new")
    Check "rules", "the start is the first status in progress from the creation status", "Analyse", "Analyse", _
          FirstInProgress("Nouveau", Moves("Nouveau", "Analyse", "Analyse", "Pret", "Pret", "Dev", "Dev", "Fini"), cats)
    Check "rules", "the nearest status in progress wins, the first listed breaks a tie", "Pret", "Pret", _
          FirstInProgress("Nouveau", Moves("Nouveau", "Idee", "Idee", "Dev", "Nouveau", "Pret", "Nouveau", "Analyse"), cats)
    Check "rules", "a loop and a status nobody leads to: none", "(empty)", "", FirstInProgress("Nouveau", Moves("Nouveau", "Idee", "Idee", "Nouveau", "Fini", "Dev"), cats)
    Check "rules", "no creation status: none", "(empty)", "", FirstInProgress("", Moves("Nouveau", "Analyse"), cats)
    Check "rules", "a creation status already in progress is the start", "Dev", "Dev", FirstInProgress("Dev", Moves("Dev", "Fini"), cats)
    Set rules = EffectiveRules(Array("Feature"), Dict("Nouveau", "new", "Analyse", "indeterminate", "Fini", "done", "Parque", "new"), none, Dict("Analyse", True), _
                               Dict("Parque", True), New Collection)
    Check "rules", "without exception Jira gives category, end, start and free", "4 rules", _
          "Feature|Analyse|Active|1|0|0|Jira;Feature|Fini|Done|0|1|0|Jira;Feature|Nouveau|Backlog|0|0|0|Jira;Feature|Parque|Backlog|0|0|1|Jira", RulesText(rules)
    Set ex = New Collection
    ex.Add Array("", "Attente", "Wait", Empty, Empty, Empty): ex.Add Array("", "Annule", "Removed", Empty, Empty, Empty)
    ex.Add Array("", "Deploy", "Wait", "", "", ""): ex.Add Array("", "Analyse", "", "", "", 1)
    Set rules = EffectiveRules(Array("Feature"), Dict("Analyse", "indeterminate", "Attente", "indeterminate", "Fini", "done", "Annule", "done", "Deploy", "done"), none, _
                               Dict("Analyse", True), Dict(), ex)
    Check "rules", "an exception wins field by field, an empty cell takes Jira's value, a Done turned Removed or Wait is no end", "5 rules", _
          "Feature|Analyse|Active|1|0|1|user;Feature|Annule|Removed|0|0|0|user;Feature|Attente|Wait|0|0|0|user;Feature|Deploy|Wait|0|0|0|user;Feature|Fini|Done|0|1|0|Jira", RulesText(rules)
    Set ex = New Collection
    ex.Add Array("", "Fini", "", "", 0, ""): ex.Add Array("", "Suite", "", "", 1, "")
    Set rules = EffectiveRules(Array("Feature"), Dict("Fini", "done", "Suite", "indeterminate"), none, Dict(), Dict(), ex)
    Check "rules", "an explicit 0 or 1 beats the default, whatever the category", "Fini 0, Suite 1", "Feature|Fini|Done|0|0|0|user;Feature|Suite|Active|0|1|0|user", RulesText(rules)
    Set ex = New Collection
    ex.Add Array("Team", "Test", "Active", "", "", ""): ex.Add Array("", "Test", "Wait", "", "", "")
    Set rules = EffectiveRules(Array("Feature", "Team"), Dict("Test", "indeterminate"), none, Dict(), Dict(), ex)
    Check "rules", "a row of one workflow touches it only and beats the row of both, whatever their order", "Feature Wait, Team Active", _
          "Feature|Test|Wait|0|0|0|user;Team|Test|Active|0|0|0|user", RulesText(rules)
    Set ex = New Collection
    ex.Add Array("", "Autre", "Wait", "", "", "")
    Set rules = EffectiveRules(Array("Feature"), Dict("Fini", "done"), Dict("Vieux", "indeterminate", "Fini", "indeterminate"), Dict(), Dict(), ex)
    Check "rules", "the current status category of a ticket is the second source, the catalog beats it, an exception alone is ruled", "3 rules", _
          "Feature|Autre|Wait|0|0|0|user;Feature|Fini|Done|0|1|0|Jira;Feature|Vieux|Active|0|0|0|Jira", RulesText(rules)
    Set rules = EffectiveRules(Array("Feature"), none, Dict("Vieux", "whatever"), Dict(), Dict(), New Collection)
    Check "rules", "a status that no source knows stays Unclassified", "Vieux", "Feature|Vieux|Unclassified|0|0|0|Jira", RulesText(rules)
    Set rules = EffectiveRules(Array("Feature", "Team"), Dict("B", "new", "A", "new"), none, Dict(), Dict(), New Collection)
    Check "rules", "the kinds in the order given, then the names", "Feature A B, Team A B", "Feature|A|Backlog|0|0|0|Jira;Feature|B|Backlog|0|0|0|Jira;Team|A|Backlog|0|0|0|Jira;Team|B|Backlog|0|0|0|Jira", RulesText(rules)
    Check "rules", "no source at all: no rule", "0", 0, EffectiveRules(Array("Feature"), none, none, none, none, New Collection).Count
    Check "rules", "the categories of one kind", "Feature A", "Backlog", RuleCategories(rules, "Feature").Item("A")
    Set ex = New Collection: ex.Add Array("", "A", "Wait", "", "", "")
    Set rules = EffectiveRules(Array("Feature"), Dict("A", "indeterminate"), none, Dict(), Dict(), ex)
    Check "rules", "a declared Wait status is told", "Wait", "Waitx", RuleCategories(rules, "Feature").Item("A") & IIf(HasWait(rules), "x", "")
    Check "rules", "no Wait status declared: told too", "no", "no", IIf(HasWait(EffectiveRules(Array("Feature"), Dict("A", "indeterminate"), none, Dict(), Dict(), New Collection)), "yes", "no")
    Set ex = New Collection: ex.Add Array("Team", "Dev", "Wait", "", "", "")
    Set rules = EffectiveRules(Array("Feature", "Team"), Dict("Funnel", "new", "Analyse", "indeterminate", "Dev", "indeterminate", "Done", "done", "Cancelled", "done"), _
                               none, Dict("Analyse", True), Dict(), ex)
    Check "rules", "the counts of each kind: starts, actives, waits, ends", "Feature 1 2 0 2, Team 1 1 1 2", _
          "status rules | Feature: start 1, active 2, wait 0, end 2 | Team: start 1, active 1, wait 1, end 2", RuleCountsLine(rules, Array("Feature", "Team"))
    Check "rules", "a start in each kind and a Wait somewhere: no warning", "0", 0, RuleWarnings(rules, Array("Feature", "Team"), Dict("Feature", True, "Team", True)).Count
    Set rules = EffectiveRules(Array("Feature", "Team"), Dict("Dev", "indeterminate", "Done", "done"), none, Dict(), Dict(), New Collection)
    Set warns = RuleWarnings(rules, Array("Feature", "Team"), Dict("Team", True))
    For Each w In warns
        s = s & IIf(Len(s) > 0, ";", "") & Left$(CStr(w), 48)
    Next w
    Check "rules", "the kind with tickets and no start, then no Wait: a warning each with its code", "Team, Wait", _
          "J2D-FLOW-010 | no start status in the Team workf;J2D-FLOW-011 | no status is declared Wait: Flow ", s
    s = ""
    For Each w In warns                                             ' never warns(1): an empty collection would raise and hang the run
        If InStr(CStr(w), "declare IsStart = 1") > 0 Or InStr(CStr(w), "declare Category = Wait") > 0 Then s = s & "x"
    Next w
    Check "rules", "the warnings name what to do", "IsStart, Wait", "xx", s
    Check "rules", "no ticket at all: no warning", "0", 0, RuleWarnings(rules, Array("Feature", "Team"), Dict()).Count
End Sub

Public Sub T_RefreshRules()
    ' 2. Compute the metrics makes the rules again from what the extraction kept of Jira under the sheet Workflows, and the columns of StatusIntervals that
    ' follow: an exception, a role of Projects or AsOf changed after the extraction takes effect without reading Jira (AL, 2026-11-22).
    Dim wb As Object, ws As Object, d0 As Date, d1 As Date, row As Variant, found As String, n As Long
    d0 = DateSerial(2026, 10, 1): d1 = DateSerial(2026, 10, 3)
    Set wb = Workbooks.Add
    PutRows wb, "Config", Array(Array("Parameter", "Value", "Example", "Mandatory", "Description"), Array("AsOf", "'2026-10-11", "", "True", ""))
    PutRows wb, "Projects", Array(Array("Project", "Role", "Team"), Array("TTT", "Program", ""), Array("AAA", "Team", "Alpha"))
    PutRows wb, "Workflows", Array(Array("Workflow", "Status", "Category", "IsStart", "IsEnd", "IsFree"), Array("", "Review", "Wait", "", "", ""))
    PutRows wb, "JiraStatusCatalog", Array(Array("Id", "Status", "JiraCategory"), Array("1", "Todo", "TODO"), Array("2", "Doing", "IN_PROGRESS"), _
                                           Array("3", "Review", "IN_PROGRESS"), Array("4", "Done", "DONE"))
    PutRows wb, "JiraWorkflows", Array(Array("Workflow", "Status", "IsStart", "IsFree"), Array("WF", "Doing", 1, 0))
    PutRows wb, "Issues", Array(Array("Key", "Project", "Status", "StatusCategory"), Array("TTT-1", "TTT", "Doing", "indeterminate"), _
                                Array("AAA-1", "AAA", "Review", "indeterminate"))
    PutRows wb, "StatusIntervals", Array(Array("Key", "Workflow", "Seq", "Status", "Category", "EnteredUtc", "LeftUtc", "DurationDays", "IsOpen", "ChangeId", "Check"), _
                                         Array("TTT-1", "Feature", 1, "Todo", "Backlog", d0, d1, 2, False, "c1", ""), _
                                         Array("TTT-1", "Feature", 2, "Doing", "Unclassified", d1, "", 99, True, "", ""), _
                                         Array("AAA-1", "Team", 1, "Review", "Active", d0, "", 99, True, "", ""))
    PutRows wb, "StatusRules", Array(Array("Workflow", "Status", "Category", "IsStart", "IsEnd", "IsFree", "Source"))
    Set ws = wb.Worksheets(SheetName("StatusIntervals"))
    RefreshRules wb
    For Each row In ReadTable(wb.Worksheets(SheetName("StatusRules")))
        n = n + 1
        If CStr(row("Status")) = "Review" Or CStr(row("Status")) = "Doing" Then
            found = found & CStr(row("Workflow")) & "|" & CStr(row("Status")) & "|" & CStr(row("Category")) & "|" & CStr(row("IsStart")) & "|" & CStr(row("Source")) & ";"
        End If
    Next row
    Check "rules", "the rules made again from the catalog, the starts of Jira and the exceptions (2 kinds x 4 statuses)", "8 rules", _
          "8|Feature|Doing|Active|1|Jira;Feature|Review|Wait|0|user;Team|Doing|Active|1|Jira;Team|Review|Wait|0|user;", n & "|" & found
    Check "rules", "the open intervals get the categories of the rules and their duration up to AsOf; a closed one keeps its own", "Active 8, Wait 10", _
          "Backlog|2|Active|8|Team|Wait|10", ws.Cells(2, 5).Value & "|" & ws.Cells(2, 8).Value & "|" & ws.Cells(3, 5).Value & "|" & ws.Cells(3, 8).Value & "|" & _
          ws.Cells(4, 2).Value & "|" & ws.Cells(4, 5).Value & "|" & ws.Cells(4, 8).Value
    PutRows wb, "Projects", Array(Array("Project", "Role", "Team"), Array("TTT", "Team", "Beta"), Array("AAA", "Program", ""))
    wb.Worksheets(SheetName("Config")).Range("B2").Value = "'2026-10-13"
    RefreshRules wb
    Check "rules", "a role of Projects and AsOf changed after the extraction", "Team | Feature, 10 | 12", "Team|Feature|10|12", _
          ws.Cells(3, 2).Value & "|" & ws.Cells(4, 2).Value & "|" & ws.Cells(3, 8).Value & "|" & ws.Cells(4, 8).Value
    PutRows wb, "Workflows", Array(Array("Workflow", "Status", "Category", "IsStart", "IsEnd", "IsFree"))
    RefreshRules wb
    Check "rules", "the Wait exception removed: Review is what Jira says (in progress)", "Active", "Active", ws.Cells(4, 5).Value
    ws.Cells(5, 1).Value = "ZZZ-9": ws.Cells(5, 2).Value = "Team": ws.Cells(5, 3).Value = 1: ws.Cells(5, 4).Value = "Ghost": ws.Cells(5, 5).Value = "Active"
    ws.Cells(5, 6).Value = d0: ws.Cells(5, 8).Value = 3: ws.Cells(5, 9).Value = False
    WriteStatusCatalog wb, New Collection, True                         ' the extraction writes it even empty (Jira refused statuses/search)
    Application.DisplayAlerts = False
    wb.Worksheets(SheetName("JiraWorkflows")).Delete                     ' no workflow of Jira read
    Application.DisplayAlerts = True
    RefreshRules wb
    Check "rules", "an empty catalog is written with its header only", "evenEmpty", "Id|", _
          CStr(wb.Worksheets(SheetName("JiraStatusCatalog")).Cells(1, 1).Value) & "|" & CStr(wb.Worksheets(SheetName("JiraStatusCatalog")).Cells(2, 1).Value)
    Check "rules", "no catalog: a status no ticket is in now is Unclassified; a ticket that Issues lacks keeps its workflow; no start without the workflows", _
          "Todo, ZZZ-9", "Unclassified|Team|Unclassified|0", ws.Cells(2, 5).Value & "|" & ws.Cells(5, 2).Value & "|" & ws.Cells(5, 5).Value & "|" & _
          CStr(Application.WorksheetFunction.CountIf(wb.Worksheets(SheetName("StatusRules")).Range("D:D"), 1))
    Application.DisplayAlerts = False
    wb.Close False
    Application.DisplayAlerts = True
End Sub

Public Sub T_WorkflowExceptions()
    ' The sheet Workflows as exceptions: a row without Workflow is read, an empty cell stays empty text, a row without a status is left out; the flagged
    ' statuses of a Jira sheet.
    Dim ws As Object, ex As Collection, flags As Object, rows As New Collection
    Set ws = ThisWorkbook.Worksheets("TestScratch")
    rows.Add Array("Team", "Attente", "Wait", "", "", "")
    rows.Add Array("", "Annule", "Removed", "", 0, "")
    rows.Add Array("Feature", " Analyse ", "", 1, "", 1)
    rows.Add Array("Team", "", "Done", "", "", "")
    WriteTable ws, Array("Workflow", "Status", "Category", "IsStart", "IsEnd", "IsFree"), rows
    Set ex = WorkflowExceptions(ws)
    Check "rules", "the exceptions of the sheet Workflows, also the row of both workflows, none without a status", "3", _
          "Team|Attente|Wait|||;|Annule|Removed||0|;Feature|Analyse||1||1", _
          Join(ex(1), "|") & ";" & Join(ex(2), "|") & ";" & Join(ex(3), "|") & IIf(ex.Count = 3, "", "?")
    ws.Cells.Clear
    Set rows = New Collection
    rows.Add Array("WF", "Open", "", 1, 0)
    rows.Add Array("WF", "Doing", "", 0, 1)
    rows.Add Array("WF", "Parked", "", "", 1)
    WriteTable ws, Array("Workflow", "Status", "Category", "IsStart", "IsFree"), rows
    Set flags = FlaggedStatuses(ws, "IsStart")
    Check "rules", "the statuses whose flag is 1", "IsStart: Open", "Open", Join(flags.Keys, ",")
    Check "rules", "another flag", "IsFree: Doing, Parked", "Doing,Parked", Join(FlaggedStatuses(ws, "IsFree").Keys, ",")
    Check "rules", "a column the sheet lacks flags nothing", "IsEnd", 0, FlaggedStatuses(ws, "IsEnd").Count
    ThisWorkbook.Worksheets("TestScratch").Cells.Clear
    Check "rules", "an empty sheet gives no exception", "0", 0, WorkflowExceptions(ThisWorkbook.Worksheets("TestScratch")).Count
End Sub

Public Sub T_NotFeatures()
    ' The Program tickets that are not Features are told (count, team, where to find them), and a FeatureTypes that matches nothing is named.
    Dim wb As Object, text As String
    Set wb = ProgramBook(True)
    text = HealthFields(wb)
    Check "health", "the tickets of the Program that are not Features, among the Program tickets", "3 of 4", True, _
          InStr(text, " | Not Features: 3 of 4 tickets of the Program (team ART (other types) in ItemFlow_from_excel, left out of the ART metrics)") > 0
    Check "health", "with a Feature in the ART there is no word about FeatureTypes", "no", False, InStr(text, "FeatureTypes") > 0
    Check "health", "the done tickets of that team are not counted as outside the iterations", "none", False, InStr(text, "Done outside the iterations") > 0
    Application.DisplayAlerts = False
    wb.Close False
    Set wb = ProgramBook(False)
    text = HealthFields(wb)
    Check "health", "no Feature at all: FeatureTypes matches no ticket and the types read are named", "Epic, Sous-tache, Technique", True, _
          InStr(text, " | Not Features: 4 of 4 tickets of the Program (team ART (other types) in ItemFlow_from_excel, left out of the ART metrics)" & _
                      " | FeatureTypes ""Feature"" matches no ticket of the Program (types read: Epic, Sous-tache, Technique): type the name of the Feature type of Jira in Config_from_user") > 0
    wb.Close False
    Application.DisplayAlerts = True
End Sub

Public Sub T_HealthFields()
    ' The fields that say why a figure is empty: only when there is something to see, with numbers and status names.
    Dim wb As Object, text As String
    Set wb = HealthBook(False)
    text = HealthFields(wb)
    Check "health", "done tickets that no iteration holds", "1 of 3", True, _
          InStr(text, " | Done outside the iterations: 1 of 3 done tickets (no metric counts them): add iterations to Iterations_from_user") > 0
    Check "health", "a ticket without a team", "1", True, InStr(text, " | Without a team: 1 tickets (a ticket takes its team from Projects_from_user)") > 0
    Check "health", "the statuses the sheet Workflows lacks, each once", "Weird, Odd", True, _
          InStr(text, " | Statuses not in Workflows: 2 (Weird, Odd): add them to Workflows_from_user with a Category") > 0
    Check "health", "the column of a formula that fails on the data", "DC03 x 2", True, _
          InStr(text, " | Cells in error: 1 column(s), first ItemFlow.DC03 x 2") > 0
    Check "health", "a kind with tickets and no start status", "Team", True, InStr(text, " | No start status: Team (no Flow Time, Flow Efficiency or Flow Load): " & _
          "set IsStart = 1 on the first in-progress status in Workflows_from_user") > 0
    Check "health", "no Wait status", "no Wait", True, InStr(text, " | No Wait status (no Flow Efficiency): set Category = Wait on the statuses where the work waits " & _
          "in Workflows_from_user") > 0
    Check "health", "a start in the second of the backlog entry", "K1", True, _
          InStr(text, " | Started in the second of the backlog entry: 1 tickets (no StartUtc, so left out of Flow Time and Flow Load)") > 0
    Check "health", "a status change after AsOf", "K3", True, InStr(text, " | AsOf before the data: 1 status changes after AsOf (Flow Load stops at AsOf): " & _
          "set AsOf in Config_from_user to the time of the extraction") > 0
    Check "health", "a row of VelocityCheck whose Gap is not 0", "1 of 2", True, InStr(text, " | VelocityCheck gap: 1 of 2 rows (sheet VelocityCheck_from_excel)") > 0
    Check "health", "the rules first, then the tickets, then the checks", "order", True, _
          InStr(text, " | No start status") < InStr(text, " | No Wait status") And InStr(text, " | No Wait status") < InStr(text, " | Done outside") And _
          InStr(text, " | Statuses not in") < InStr(text, " | Started in the second") And InStr(text, " | AsOf before") < InStr(text, " | VelocityCheck gap")
    wb.Worksheets(SheetName("Config")).Range("B2").Value = "'x"           ' an AsOf that cannot be read: that field says nothing, the others stay
    text = HealthFields(wb)
    Check "health", "an AsOf that cannot be read leaves out its field only", "x", "no|yes", _
          IIf(InStr(text, " | AsOf before") > 0, "yes", "no") & "|" & IIf(InStr(text, " | VelocityCheck gap: 1 of 2") > 0, "yes", "no")
    wb.Worksheets(SheetName("Issues")).Range("A2:A5").ClearContents       ' no ticket read (a wrong Filter): nothing about the rules
    text = HealthFields(wb)
    Check "health", "no ticket read: nothing about the start or Wait statuses", "no rows", "no|no", _
          IIf(InStr(text, " | No start status") > 0, "yes", "no") & "|" & IIf(InStr(text, " | No Wait status") > 0, "yes", "no")
    Application.DisplayAlerts = False
    wb.Close False
    Set wb = HealthBook(True)
    Check "health", "a clean run adds nothing", "(empty)", "", HealthFields(wb)
    wb.Close False
    Application.DisplayAlerts = True
    Check "health", "it never raises: no workbook, no fields", "Nothing", "", HealthFields(Nothing)
End Sub

' ================= The type of the columns written from Jira =================

Public Sub T_CellKind()
    Check "types", "Empty", "Empty", "empty", CellKind(Empty)
    Check "types", "Null", "Null", "empty", CellKind(Null)
    Check "types", "an empty text", "", "empty", CellKind("")
    Check "types", "a text of digits stays a text", "1761134538011", "text", CellKind("1761134538011")
    Check "types", "a boolean", "True", "boolean", CellKind(True)
    Check "types", "a date", "DateSerial", "date", CellKind(DateSerial(2026, 11, 14))
    Check "types", "a number: epoch milliseconds", "1761134538011", "number", CellKind(1761134538011#)
    Check "types", "a whole number", "3", "number", CellKind(3)
    Check "types", "a currency", "CCur", "number", CellKind(CCur(3))
    Check "types", "an error", "#N/A", "error", CellKind(CVErr(2042))
    Check "types", "a list", "Array", "other", CellKind(Array(1, 2))
End Sub

Private Sub FitCase(ByVal typ As String, ByVal label As String, ByVal v As Variant, ByVal expected As Boolean, Optional ByVal allowed As String = "")
    Check "types", typ & ": " & label, typ, expected, FitsType(typ, v, allowed)
End Sub

Public Sub T_FitsType()
    ' The same table as tests/unit/test_raw_types.py::test_fits: the two sides must agree.
    Dim d As Date, lists As String
    d = DateSerial(2026, 11, 14) + TimeSerial(9, 30, 0)
    FitCase "datetime", "a date", d, True
    FitCase "datetime", "2000-01-01", DateSerial(2000, 1, 1), True
    FitCase "datetime", "1999-12-31 23:59", DateSerial(1999, 12, 31) + TimeSerial(23, 59, 0), False
    FitCase "datetime", "2100-01-01", DateSerial(2100, 1, 1), False
    FitCase "datetime", "epoch milliseconds", 1761134538011#, False
    FitCase "datetime", "an ISO text", "2026-11-14T09:30:00", False
    FitCase "datetime", "a boolean", True, False
    FitCase "date", "midnight", DateSerial(2026, 6, 1), True
    FitCase "date", "a time of day", d, False
    FitCase "date", "a text", "2026-06-01", False
    FitCase "id", "digits as text", "10042", True
    FitCase "id", "a number: Excel converted it", 10042, False
    FitCase "id", "a key", "ABC-1", False
    FitCase "id", "a space", "10 042", False
    FitCase "key", "ABC-12", "ABC-12", True
    FitCase "key", "digit and underscore in the project", "A1_B-7", True
    FitCase "key", "lower case", "abc-12", False
    FitCase "key", "no dash", "ABC12", False
    FitCase "key", "no number", "ABC-", False
    FitCase "key", "starts with a digit", "1AB-1", False
    FitCase "key", "a number", 12, False
    FitCase "integer", "3", 3, True
    FitCase "integer", "3.0", 3#, True
    FitCase "integer", "3.5", 3.5, False
    FitCase "integer", "a text", "3", False
    FitCase "integer", "a boolean", True, False
    FitCase "integer", "a date", d, False
    FitCase "number", "3.5", 3.5, True
    FitCase "number", "0", 0, True
    FitCase "number", "a text", "3.5", False
    FitCase "number", "a boolean", True, False
    FitCase "number", "a date", d, False
    FitCase "boolean", "True", True, True
    FitCase "boolean", "False", False, True
    FitCase "boolean", "1", 1, False
    FitCase "boolean", "a text", "TRUE", False
    lists = "future, active, closed"
    FitCase "enum", "in the list", "active", True, lists
    FitCase "enum", "another case", "Active", False, lists
    FitCase "enum", "outside the list", "done", False, lists
    FitCase "enum", "two values of the list in one text", "future, active", False, lists
    FitCase "enum", "a number", 1, False, lists
    FitCase "text", "a name", "To Do", True
    FitCase "text", "a text of digits", "123", True
    FitCase "text", "a number", 123, False
    FitCase "text", "a date", d, False
    FitCase "text", "a boolean", True, False
    FitCase "list", "one item", "1.0", True
    FitCase "list", "two items", "1.0, 2.0", True
    FitCase "list", "an empty item", "1.0, , 2.0", False
    FitCase "list", "a trailing separator", "1.0,", False
    FitCase "list", "a number", 1.5, False
    Check "types", "the date limits are those of the python side", "serials", "36526|73051", CStr(CLng(DateSerial(2000, 1, 1))) & "|" & CStr(CLng(DateSerial(2100, 1, 1)))
End Sub

Public Sub T_ShapeOfCell()
    ' The type and the shape of a cell, none of its content (raw_types.shape_of).
    Dim d As Date
    d = DateSerial(2026, 11, 14) + TimeSerial(9, 30, 0)
    Check "types", "epoch milliseconds", "1761134538011", "Double:9999999999999", ShapeOfCell(1761134538011#)
    Check "types", "a decimal number", "3.5", "Double:9.9", ShapeOfCell(3.5)
    Check "types", "a key", "ABC-12", "String:AAA-99", ShapeOfCell("ABC-12")
    Check "types", "a name", "To Do", "String:Aa Aa", ShapeOfCell("To Do")
    Check "types", "a date", "DateSerial", "Date:9999-99-99A99:99:99", ShapeOfCell(d)
    Check "types", "a boolean", "True", "Boolean", ShapeOfCell(True)
    Check "types", "an error", "#N/A", "Error", ShapeOfCell(CVErr(2042))
    Check "types", "Empty and an empty text", "", "Empty|Empty", ShapeOfCell(Empty) & "|" & ShapeOfCell("")
    Check "types", "a list", "Array", "Object", ShapeOfCell(Array(1))
    Check "types", "a decimal below 1 has its zero, as python writes it", "0.5", "Double:9.9", ShapeOfCell(0.5)
    Check "types", "a whole number of 16 digits is written in full", "1E15", "Double:" & String$(16, "9"), ShapeOfCell(1E+15)
    Check "types", "cut at 40 characters", "50 letters", "String:" & String$(40, "a") & "...", ShapeOfCell(String$(50, "x"))
    Check "types", "no letter and no digit of the value", "PI#26-2026 secret", "String:AA#99-9999 aaaaaa", ShapeOfCell("PI#26-2026 secret")
End Sub

Public Sub T_TypeHints()
    ' One sentence per (type, kind found): the texts of raw_types.hint (a unit test compares every literal of HintFor with them).
    Check "types", "a column that is not in the sheet", "missing", True, InStr(HintFor("datetime", "missing", True, "", 0), "the column is not in the sheet") = 1
    Check "types", "a required date column that is empty points at J2D-DATA-003", "empty", True, _
          InStr(HintFor("datetime", "empty", True, "", 0), "no cell holds a date: Jira's date was not read (see WARN J2D-DATA-003") = 1
    Check "types", "an optional date column that is empty is fine", "empty", "no ticket has this date: fine when the field is not used", HintFor("date", "empty", False, "", 0)
    Check "types", "a required column that is empty", "empty", True, InStr(HintFor("text", "empty", True, "", 0), "a required column is empty") = 1
    Check "types", "an optional column that is empty is fine", "empty", True, InStr(HintFor("number", "empty", False, "", 0), "no ticket has a value") = 1
    Check "types", "13 digits are epoch milliseconds", "13", "13 digits: epoch milliseconds, Jira sent a number and the date was not converted (modText.ParseJiraDate)", HintFor("datetime", "number", True, "", 13)
    Check "types", "10 digits are epoch seconds", "10", "10 digits: epoch seconds, Jira sent a number and the date was not converted (modText.ParseJiraDate)", HintFor("datetime", "text", True, "", 10)
    Check "types", "a date out of range", "date", True, InStr(HintFor("datetime", "date", True, "", 0), "a date outside 2000-2099") = 1
    Check "types", "a text that is no date", "text", True, InStr(HintFor("datetime", "text", True, "", 0), "a text that is not read as a date") = 1
    Check "types", "another kind than a date", "boolean", True, InStr(HintFor("datetime", "boolean", True, "", 0), "not a date") = 1
    Check "types", "a number in an id column", "number", True, InStr(HintFor("id", "number", True, "", 0), "Excel turned the text into a number") = 1
    Check "types", "another text in a key column", "text", True, InStr(HintFor("key", "text", True, "", 0), "not a Jira id or key (ABC-12 or digits)") = 1
    Check "types", "a text in a number column", "text", True, InStr(HintFor("integer", "text", True, "", 0), "a text where a number is expected") = 1
    Check "types", "decimals in a whole-number column", "number", True, InStr(HintFor("integer", "number", True, "", 0), "a number with decimals") = 1
    Check "types", "another kind in a number column", "boolean", True, InStr(HintFor("number", "boolean", True, "", 0), "not a number") = 1
    Check "types", "a boolean column", "text", True, InStr(HintFor("boolean", "text", True, "", 0), "not TRUE/FALSE") = 1
    Check "types", "an enum names its list", "enum", "not one of future, active, closed: Jira added a value or the field is another one", HintFor("enum", "text", False, "future, active, closed", 0)
    Check "types", "a list with an empty item", "list", True, InStr(HintFor("list", "text", False, "", 0), "an item of the list is empty") = 1
    Check "types", "a text column", "text", True, InStr(HintFor("text", "number", True, "", 0), "a text expected") = 1
End Sub

Public Sub T_TypeEmptyTable()
    ' A table with no row is not read: no alarm for its required columns. A table with no data keeps one blank row (CreateTable); a sheet
    ' written as a range with only its header has none; a table the sheet does not have is skipped.
    Dim wb As Object, file As String, saved As String, text As String
    saved = gLogPath
    file = Environ$("TEMP") & "\j2d_empty_test_" & Format$(Now, "hhnnss") & Int(Rnd() * 1000) & ".log"
    gLogPath = file
    Set wb = Workbooks.Add
    PutRows wb, "RawTypesSpec", Array(Array("Table", "Column", "Type", "Code", "Required", "Source", "Allowed"), _
        Array("Changelog", "Key", "key", "J2D-DATA-007", "yes", "key", ""), _
        Array("IssueLinks", "Blocker", "key", "J2D-DATA-007", "yes", "x", ""), _
        Array("Absent", "A", "text", "J2D-DATA-013", "yes", "x", ""))
    PutRows wb, "Changelog", Array(Array("Key", "IssueId"))
    PutRows wb, "IssueLinks", Array(Array("Blocker", "Blocked"), Array("", ""))
    MakeTable wb, "IssueLinks", 2, 2
    Application.DisplayAlerts = False
    wb.Worksheets(1).Delete
    Check "types", "tables with no row raise no alarm", "(empty)", "", TypeProblems(wb)
    wb.Close False
    Application.DisplayAlerts = True
    text = FileText(file)
    Check "types", "none of them was read: no WARN, no fill line, a summary of zero", "tables 0", True, _
          InStr(text, " WARN ") = 0 And InStr(text, " INFO types ") = 0 And InStr(text, " INFO column types | tables 0 | columns 0 | checked 0 | problems 0" & vbCrLf) > 0
    gLogPath = saved
    Kill file
End Sub

Public Sub T_TypeBlocks()
    ' The first filled cells are looked for by blocks of rows, the filled ones counted by CountA: a column filled only far down is found,
    ' on a sheet written as a range (no table), at its true cell.
    Dim wb As Object, ws As Object, data() As Variant, r As Long, file As String, saved As String, text As String
    saved = gLogPath
    file = Environ$("TEMP") & "\j2d_blocks_test_" & Format$(Now, "hhnnss") & Int(Rnd() * 1000) & ".log"
    gLogPath = file
    Set wb = Workbooks.Add
    PutRows wb, "RawTypesSpec", Array(Array("Table", "Column", "Type", "Code", "Required", "Source", "Allowed"), _
        Array("Many", "Value", "key", "J2D-DATA-007", "no", "src.value", ""))
    PutRows wb, "Many", Array(Array("Value", "Other"), Array("", "x"))
    Set ws = wb.Worksheets("Many")
    ReDim data(1 To 4500, 1 To 1)
    For r = 4100 To 4500
        data(r, 1) = "bad"
    Next r
    ws.Range(ws.Cells(2, 1), ws.Cells(4501, 1)).Value = data
    Application.DisplayAlerts = False
    wb.Worksheets(1).Delete
    Check "types", "a column filled from row 4100 is found", "field", True, InStr(TypeProblems(wb), " | Column types: 1 column(s), first Many.Value expected key found text") = 1
    wb.Close False
    Application.DisplayAlerts = True
    text = FileText(file)
    Check "types", "its first cell is the true one", "A4101", True, _
          InStr(text, " WARN J2D-DATA-007 | Many.Value | expected key | found text | source src.value | first filled A4101" & vbCrLf) > 0
    Check "types", "the filled cells are all counted, the kinds are those of the first 20", "401", True, _
          InStr(text, " INFO J2D-DATA-007 | Many.Value | rows 4500, filled 401, empty 4099 | kinds in first 20: text x 20 | value shape String:aaa | cell format General" & vbCrLf) > 0
    Check "types", "and the fill line says it", "401", True, InStr(text, " INFO types Many | rows 4500 | filled: Value 401" & vbCrLf) > 0
    gLogPath = saved
    Kill file
End Sub

Public Sub T_TypeFillLines()
    ' The fill counts of a wide table go on a next line: the log cuts a line at 300 characters (raw_types.fill_lines does the same).
    Dim wb As Object, spec() As Variant, hdr() As Variant, one() As Variant, c As Long, file As String, saved As String, text As String
    Dim first As String, second As String
    saved = gLogPath
    file = Environ$("TEMP") & "\j2d_fill_test_" & Format$(Now, "hhnnss") & Int(Rnd() * 1000) & ".log"
    gLogPath = file
    ReDim spec(0 To 22): ReDim hdr(0 To 21): ReDim one(0 To 21)
    spec(0) = Array("Table", "Column", "Type", "Code", "Required", "Source", "Allowed")
    For c = 0 To 21
        hdr(c) = "WideColumn" & Format$(c + 1, "00"): one(c) = "x"
        spec(c + 1) = Array("Wide", hdr(c), "text", "J2D-DATA-013", "no", "src", "")
    Next c
    Set wb = Workbooks.Add
    PutRows wb, "RawTypesSpec", spec
    PutRows wb, "Wide", Array(hdr, one, one)
    MakeTable wb, "Wide", 3, 22
    Application.DisplayAlerts = False
    wb.Worksheets(1).Delete
    Check "types", "a wide table with all its columns well typed: no field", "(empty)", "", TypeProblems(wb)
    wb.Close False
    Application.DisplayAlerts = True
    text = FileText(file)
    first = "types Wide | rows 2 | filled: WideColumn01 2"
    For c = 2 To 15
        first = first & ", WideColumn" & Format$(c, "00") & " 2"
    Next c
    second = "types Wide | filled (continued): WideColumn16 2"
    For c = 17 To 22
        second = second & ", WideColumn" & Format$(c, "00") & " 2"
    Next c
    Check "types", "the first line holds 15 columns", "268 characters", True, InStr(text, " INFO " & first & vbCrLf) > 0
    Check "types", "the next line holds the 7 others", "continued", True, InStr(text, " INFO " & second & vbCrLf) > 0
    Check "types", "no line is cut", "...", False, InStr(text, "...") > 0
    gLogPath = saved
    Kill file
End Sub

Public Sub T_WireTypes()
    ' What Jira really sent, by JSON type: names and types, never a value (python: mapping.wire_lines, the same words).
    Dim issues As New Collection, raw As New Collection, lines() As String, ids As Variant
    ResetDateKey
    Check "wire", "an object", "{}", "object", WireWord(JsonParse("{}"))
    Check "wire", "an array", "[]", "array", WireWord(JsonParse("[]"))
    Check "wire", "a string", "x", "string", WireWord("x")
    Check "wire", "a number: epoch milliseconds", "1761134538011", "number", WireWord(1761134538011#)
    Check "wire", "a boolean", "True", "boolean", WireWord(True)
    Check "wire", "null", "Null", "null", WireWord(Null)
    Check "wire", "absent", "Empty", "absent", WireWord(Empty)
    issues.Add JsonParse(J("{'id':'1','key':'ABC-1','fields':{'created':'2026-01-01T00:00:00.000+0000','resolutiondate':null,'status':{'name':'Done'}," & _
                           "'customfield_3':'Feature','customfield_2':[{'state':'active'}],'customfield_1':null}}"))
    issues.Add JsonParse(J("{'id':'2','key':'ABC-2','fields':{'created':'2026-01-02T00:00:00.000+0000','resolutiondate':null,'status':{'name':'Open'}," & _
                           "'customfield_1':5}}"))
    raw.Add Array("10001", "5", 1761134538011#, "n", "a", "status", "Status", "jira", "1", "To Do", "2", "Done")
    ids = Array("customfield_2", "customfield_1", "", "", "", "", "")
    lines = Split(WireTypesLines(issues, raw, ids, "customfield_3"), vbLf)
    Check "wire", "three lines", "3", 2, UBound(lines)
    Check "wire", "the issue: a null stays null, a field nobody has is absent", "line 1", _
          "wire types issue | id string, key string, fields.created string, fields.resolutiondate null, fields.status.name string, fields.parent.key absent", lines(0)
    Check "wire", "the fields of Config: the first value present, a field not set says so", "line 2", _
          "wire types configured fields | WorkTypeField string, SprintField array, StoryPointsField number, PIField not set, BusinessValueField not set, " & _
          "TeamField not set, TargetStartField not set, TargetEndField not set", lines(1)
    Check "wire", "the changelog: a date sent as a number says number", "line 3", _
          "wire types changelog | id string, created number, fieldId string, from string, to string", lines(2)
    raw.Add Array("10002", "6", 1761134538011#, "n", "a", "status", "Status", "jira", Null, "To Do", Null, "Done")
    Check "wire", "a row with nulls after a filled one changes nothing", "line 3", _
          "wire types changelog | id string, created number, fieldId string, from string, to string", Split(WireTypesLines(issues, raw, ids, ""), vbLf)(2)
    Dim nulls As New Collection
    nulls.Add Array("10002", "6", 1761134538011#, "n", "a", "status", "Status", "jira", Null, "To Do", Null, "Done")
    Check "wire", "all null: null", "line 3", "wire types changelog | id string, created number, fieldId string, from null, to null", _
          Split(WireTypesLines(issues, nulls, ids, ""), vbLf)(2)
    Check "wire", "no history", "line 3", "wire types changelog | no history", Split(WireTypesLines(issues, New Collection, ids, ""), vbLf)(2)
    Check "wire", "a work type field that is not set", "line 2", True, _
          InStr(Split(WireTypesLines(issues, raw, ids, ""), vbLf)(1), "wire types configured fields | WorkTypeField not set, SprintField array") = 1
End Sub

Private Sub MakeTable(ByVal wb As Object, ByVal logicalName As String, ByVal nRows As Long, ByVal nCols As Long)
    Dim ws As Object
    Set ws = wb.Worksheets(SheetName(logicalName))
    ws.ListObjects.Add(1, ws.Range(ws.Cells(1, 1), ws.Cells(nRows, nCols)), , 1).Name = logicalName
End Sub

Private Function TypesBook(ByVal clean As Boolean) As Object
    ' A scratch workbook: the sheet RawTypesSpec (one column of several types), Issues as a table and Changelog as a plain range, as the
    ' real workbook has them (the sheet Changelog is written by WriteTable, no table), wrong or not.
    Dim wb As Object, d As Variant
    Set wb = Workbooks.Add
    d = DateSerial(2026, 11, 14) + TimeSerial(9, 30, 0)
    PutRows wb, "RawTypesSpec", Array(Array("Table", "Column", "Type", "Code", "Required", "Source", "Allowed"), _
        Array("Changelog", "CreatedUtc", "datetime", "J2D-DATA-004", "yes", "histories[].created", ""), _
        Array("Changelog", "IssueId", "id", "J2D-DATA-006", "yes", "issueChangeLogs[].issueId", ""), _
        Array("Changelog", "FromString", "text", "J2D-DATA-013", "no", "histories[].items[].fromString", ""), _
        Array("Changelog", "Missing", "text", "J2D-DATA-013", "no", "x", ""), _
        Array("Issues", "Key", "key", "J2D-DATA-007", "yes", "key", ""), _
        Array("Issues", "State", "enum", "J2D-DATA-011", "no", "fields.sprint.state", "future, active, closed"), _
        Array("Issues", "Points", "number", "J2D-DATA-009", "no", "fields.points", ""), _
        Array("Issues", "Open", "boolean", "J2D-DATA-010", "yes", "computed", ""), _
        Array("Other", "A", "text", "J2D-DATA-013", "no", "x", ""))
    If clean Then
        PutRows wb, "Changelog", Array(Array("CreatedUtc", "IssueId", "FromString", "Missing"), Array(d, "'10001", "To Do", "x"), Array(d + 1, "'10002", "Done", "y"))
        PutRows wb, "Issues", Array(Array("Key", "State", "Points", "Open"), Array("ABC-1", "active", 5, True), Array("ABC-2", "closed", 3, False))
    Else
        PutRows wb, "Changelog", Array(Array("CreatedUtc", "IssueId", "FromString"), Array(1761134538011#, 10001, ""), Array(1761134538012#, 10002, ""))
        PutRows wb, "Issues", Array(Array("Key", "State", "Points", "Open"), Array("ABC-1", "", "5 points", ""), Array("ABC-2", "paused", 3, ""))
    End If
    MakeTable wb, "Issues", 3, 4
    Application.DisplayAlerts = False
    wb.Worksheets(1).Delete
    Application.DisplayAlerts = True
    Set TypesBook = wb
End Function

Public Sub T_TypeProblems()
    ' The first filled cell of each column against its type: a WARN line per problem with its code and what helps to debug (names, counts,
    ' a shape, never a value), an INFO line per table, a summary, and the field of the status; nothing at all when all is well.
    Dim wb As Object, text As String, field As String, file As String, saved As String
    saved = gLogPath
    file = Environ$("TEMP") & "\j2d_types_test_" & Format$(Now, "hhnnss") & Int(Rnd() * 1000) & ".log"
    gLogPath = file
    Set wb = TypesBook(False)
    field = TypeProblems(wb)
    text = FileText(file)
    Check "types", "the status field names the first five columns and says there are more", "6 problems", _
          " | Column types: 6 column(s), first Changelog.CreatedUtc expected datetime found number; Changelog.IssueId expected id found number; " & _
          "Changelog.Missing expected text found missing; Issues.State expected enum found text; Issues.Points expected number found text; ...", field
    Check "types", "epoch milliseconds in a date column: the v0.7.0 failure, the WARN line", "J2D-DATA-004", True, _
          InStr(text, " WARN J2D-DATA-004 | Changelog.CreatedUtc | expected datetime | found number | source histories[].created | first filled A2" & vbCrLf) > 0
    Check "types", "... its counts, kinds, shape and cell format", "J2D-DATA-004", True, _
          InStr(text, " INFO J2D-DATA-004 | Changelog.CreatedUtc | rows 2, filled 2, empty 0 | kinds in first 20: number x 2 | value shape Double:9999999999999 | cell format General" & vbCrLf) > 0
    Check "types", "... and its hint, which names the cause", "J2D-DATA-004", True, _
          InStr(text, " INFO J2D-DATA-004 | Changelog.CreatedUtc | hint: 13 digits: epoch milliseconds, Jira sent a number and the date was not converted (modText.ParseJiraDate)" & vbCrLf) > 0
    Check "types", "a number in an id column", "J2D-DATA-006", True, _
          InStr(text, " WARN J2D-DATA-006 | Changelog.IssueId | expected id | found number | source issueChangeLogs[].issueId | first filled B2" & vbCrLf) > 0 And _
          InStr(text, " INFO J2D-DATA-006 | Changelog.IssueId | rows 2, filled 2, empty 0 | kinds in first 20: number x 2 | value shape Double:99999 | cell format General" & vbCrLf) > 0 And _
          InStr(text, " INFO J2D-DATA-006 | Changelog.IssueId | hint: Excel turned the text into a number (leading apostrophe lost) or Jira sent a number: keep ids as text" & vbCrLf) > 0
    Check "types", "a column the sheet lacks", "J2D-DATA-013", True, _
          InStr(text, " WARN J2D-DATA-013 | Changelog.Missing | expected text | found missing | source x | first filled -" & vbCrLf) > 0 And _
          InStr(text, " INFO J2D-DATA-013 | Changelog.Missing | rows 2, filled 0, empty 2 | kinds in first 20: - | value shape Empty" & vbCrLf) > 0 And _
          InStr(text, " INFO J2D-DATA-013 | Changelog.Missing | hint: the column is not in the sheet: the sheet and RawTypesSpec differ (an old workbook?)" & vbCrLf) > 0
    Check "types", "the first FILLED cell is read, not the first row", "B3", True, _
          InStr(text, " WARN J2D-DATA-011 | Issues.State | expected enum | found text | source fields.sprint.state | first filled B3" & vbCrLf) > 0 And _
          InStr(text, " INFO J2D-DATA-011 | Issues.State | rows 2, filled 1, empty 1 | kinds in first 20: text x 1 | value shape String:aaaaaa | cell format General" & vbCrLf) > 0 And _
          InStr(text, " INFO J2D-DATA-011 | Issues.State | hint: not one of future, active, closed: Jira added a value or the field is another one" & vbCrLf) > 0
    Check "types", "a text in a number column, the kinds are counted", "J2D-DATA-009", True, _
          InStr(text, " WARN J2D-DATA-009 | Issues.Points | expected number | found text | source fields.points | first filled C2" & vbCrLf) > 0 And _
          InStr(text, " INFO J2D-DATA-009 | Issues.Points | rows 2, filled 2, empty 0 | kinds in first 20: number x 1, text x 1 | value shape String:9 aaaaaa | cell format General" & vbCrLf) > 0 And _
          InStr(text, " INFO J2D-DATA-009 | Issues.Points | hint: a text where a number is expected: Jira sent text, or the custom field is of another type (check its id in Config)" & vbCrLf) > 0
    Check "types", "a required column that is empty", "J2D-DATA-010", True, _
          InStr(text, " WARN J2D-DATA-010 | Issues.Open | expected boolean | found empty | source computed | first filled -" & vbCrLf) > 0 And _
          InStr(text, " INFO J2D-DATA-010 | Issues.Open | rows 2, filled 0, empty 2 | kinds in first 20: - | value shape Empty" & vbCrLf) > 0 And _
          InStr(text, " INFO J2D-DATA-010 | Issues.Open | hint: a required column is empty: the extraction wrote nothing there (check the field and its id in Config)" & vbCrLf) > 0
    Check "types", "an optional column that is empty says nothing", "FromString", False, InStr(text, "WARN J2D-DATA-013 | Changelog.FromString") > 0
    Check "types", "a table of the spec the workbook lacks is not read", "Other", False, InStr(text, "Other.A") > 0
    Check "types", "the filled cells of each column, in one line per table", "INFO", True, _
          InStr(text, " INFO types Changelog | rows 2 | filled: CreatedUtc 2, IssueId 2, FromString 0" & vbCrLf) > 0 And _
          InStr(text, " INFO types Issues | rows 2 | filled: Key 2, State 1, Points 2, Open 0" & vbCrLf) > 0
    Check "types", "the summary", "INFO", True, InStr(text, " INFO column types | tables 2 | columns 8 | checked 5 | problems 6" & vbCrLf) > 0
    Check "types", "per table: its problems, then its fill line (the order of raw_types.log_lines)", "order", True, _
          InStr(text, " WARN J2D-DATA-013 | Changelog.Missing") < InStr(text, " INFO types Changelog | rows 2") And _
          InStr(text, " INFO types Changelog | rows 2") < InStr(text, " WARN J2D-DATA-011 | Issues.State") And _
          InStr(text, " WARN J2D-DATA-010 | Issues.Open") < InStr(text, " INFO types Issues | rows 2")
    Check "types", "the check is one step of the log", "BEGIN END", True, InStr(text, " BEGIN column types" & vbCrLf) > 0 And InStr(text, " END column types | ") > 0
    Application.DisplayAlerts = False
    wb.Close False
    Kill file
    gLogPath = file
    Set wb = TypesBook(True)
    Check "types", "all well: no field", "(empty)", "", TypeProblems(wb)
    text = FileText(file)
    Check "types", "all well: no WARN line, a summary with no problem", "problems 0", True, _
          InStr(text, " WARN ") = 0 And InStr(text, " INFO column types | tables 2 | columns 8 | checked 8 | problems 0" & vbCrLf) > 0
    wb.Close False
    Application.DisplayAlerts = True
    Kill file
    Check "types", "it never raises: no workbook", "Nothing", "", TypeProblems(Nothing)
    gLogPath = saved
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
    Set d = JsonParse(J("{'a':[1,2],'b':{'c':'x','d':{'e':7}},'f':false}"))
    Check "json", "JGet on an array is Empty (no Exists on a Collection)", "a", True, IsEmpty(JGet(d("a"), "c"))
    Check "json", "JGet on Nothing is Empty", "Nothing", True, IsEmpty(JGet(Nothing, "c"))
    Check "json", "JPath through an array is Empty", "a.c", True, IsEmpty(JPath(d, "a.c"))
    Check "json", "JPath reads a value after two objects", "b.d.e", 7, JPath(d, "b.d.e")
    Check "json", "JPath returns an object as an object", "b.d", "Dictionary", TypeName(JPath(d, "b.d"))
    Check "json", "JPath keeps False", "f", False, JPath(d, "f")
    Dim missing As Variant
    missing = JGet(d, "zz")
    Check "json", "a missing key is not added by JGet", "zz", 3, d.Count
    Check "json", "JPath of a scalar start", "x", True, IsEmpty(JPath(5, "a.b"))
    Check "json", "AsList returns an array as it is", "a", 2, AsList(d("a")).Count
    Check "json", "AsList of an object is an empty list", "b", 0, AsList(d("b")).Count
    Set d = JsonParse(J("{'a':1,'A':2,'b':{'x':3}}"))
    Check "json", "keys are case-sensitive (Dictionary compare mode 0)", "a, A", "1|2", d("a") & "|" & d("A")
    Check "json", "an object holds three keys", "count", 3, d.Count
    Check "json", "the compare mode is binary", "CompareMode", 0, d.CompareMode
    Check "json", "a nested object is binary too", "b", 0, d("b").CompareMode
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
    Check "jira", "page guard message", "maxPages=1", True, (InStr(msg, "did not end after 1 pages") > 0)
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

' ================= Parameters (sheets Workflows, Transitions) =================

Public Sub T_ParamsTables()
    ' Hermetic: the tables are written on the scratch sheet, with ASCII fixtures.
    Dim ws As Object, rows As New Collection, t As Collection, cats As Object
    Set ws = ThisWorkbook.Worksheets("TestScratch")
    rows.Add Array("Team", "Nouveau", "Backlog", 0, 0, 0)
    rows.Add Array("Team", "En analyse", "Active", 1, 0, 0)
    rows.Add Array("Team", "Termine", "Done", 0, 1, 0)
    rows.Add Array("Team", "Deprioritise", "Backlog", Null, 0, 1)
    rows.Add Array("Feature", "En analyse", "Wait", 1, 0, 0)
    WriteTable ws, Array("Workflow", "Status", "Category", "IsStart", "IsEnd", "IsFree"), rows
    Set t = ReadTable(ws)
    Check "params", "one dictionary per row", "5 data rows", 5, t.Count
    Check "params", "headers are the keys", "row 2 Status", "En analyse", t(2)("Status")
    Check "params", "numbers stay numbers", "row 2 IsStart", 1, t(2)("IsStart")
    Check "params", "empty cell is empty", "row 4 IsStart", Empty, t(4)("IsStart")
    Set cats = StatusCategories(ws, "Team")
    Check "params", "Team has 4 statuses", "StatusCategories", 4, cats.Count
    Check "params", "category of a status", "Team En analyse", "Active", cats("En analyse")
    Check "params", "same status, other workflow", "Feature En analyse", "Wait", StatusCategories(ws, "Feature")("En analyse")
    Check "params", "unknown workflow is empty", "Other", 0, StatusCategories(ws, "Other").Count
    Check "params", "start status", "Team IsStart", "En analyse", FlaggedStatus(ws, "Team", "IsStart")
    Check "params", "end status", "Team IsEnd", "Termine", FlaggedStatus(ws, "Team", "IsEnd")
    Check "params", "free status", "Team IsFree", "Deprioritise", FlaggedStatus(ws, "Team", "IsFree")
    Check "params", "no flagged status gives empty text", "Feature IsEnd", "", FlaggedStatus(ws, "Feature", "IsEnd")
    Check "params", "flag on", "Team Termine IsEnd", True, StatusHasFlag(ws, "Team", "Termine", "IsEnd")
    Check "params", "flag off", "Team Nouveau IsEnd", False, StatusHasFlag(ws, "Team", "Nouveau", "IsEnd")
    Check "params", "flag of an unknown status", "Team Nope IsEnd", False, StatusHasFlag(ws, "Team", "Nope", "IsEnd")
    Set rows = New Collection
    rows.Add Array("Team", "Nouveau", "Backlog", 0, 0, 0)
    rows.Add Array("", "ignored: first cell is empty", "", 0, 0, 0)
    WriteTable ws, Array("Workflow", "Status", "Category", "IsStart", "IsEnd", "IsFree"), rows
    Check "params", "rows without a first cell are ignored", "2 rows written", 1, ReadTable(ws).Count
    ws.Cells.ClearContents
    Check "params", "empty sheet gives no rows", "", 0, ReadTable(ws).Count
    WriteTable ws, Array("Workflow", "Status"), New Collection
    Check "params", "header only gives no rows", "", 0, ReadTable(ws).Count
End Sub

Public Sub T_ParamsOnWorkbook()
    ' The sheets written by the generator from params\*.md. Accents are built with ChrW (VBA source is ASCII).
    Dim pret As String, depri As String, deployer As String, termine As String, wsW As Object
    pret = "Pr" & ChrW(234) & "t"
    depri = "D" & ChrW(233) & "prioris" & ChrW(233)
    deployer = ChrW(192) & " d" & ChrW(233) & "ployer"
    termine = "Termin" & ChrW(233)
    Set wsW = ThisWorkbook.Worksheets(SheetName("Workflows"))
    Check "params", "Team start status", "Team IsStart", "En analyse", FlaggedStatus(wsW, "Team", "IsStart")
    Check "params", "Team end status", "Team IsEnd", termine, FlaggedStatus(wsW, "Team", "IsEnd")
    Check "params", "Team category", "Team En analyse", "Active", CategoryOf("Team", "En analyse")
    Check "params", "Team waiting status", "Team Pret", "Wait", CategoryOf("Team", pret)
    Check "params", "backlog status", "Team Depriorise", "Backlog", CategoryOf("Team", depri)
    Check "params", "Feature has its own test status", "Feature En test / recette", "Active", CategoryOf("Feature", "En test / recette")
    Check "params", "Team does not know it", "Team En test / recette", "Unclassified", CategoryOf("Team", "En test / recette")
    Check "params", "unknown status", "Team Inconnu", "Unclassified", CategoryOf("Team", "Inconnu")
    Check "params", "forward move", "Team Nouveau > En analyse", True, IsAllowed("Team", "Nouveau", "En analyse")
    Check "params", "back move", "Team En analyse < Pret", True, IsAllowed("Team", pret, "En analyse")
    Check "params", "skipping a step is not allowed", "Team Nouveau > Pret", False, IsAllowed("Team", "Nouveau", pret)
    Check "params", "free status from any status", "Team Nouveau > Depriorise", True, IsAllowed("Team", "Nouveau", depri)
    Check "params", "exit from the free status", "Team Depriorise > Annule", True, IsAllowed("Team", depri, "Annul" & ChrW(233))
    Check "params", "Team test to deploy", "Team En test > A deployer", True, IsAllowed("Team", "En test", deployer)
    Check "params", "Feature test to deploy", "Feature En test / recette > A deployer", True, IsAllowed("Feature", "En test / recette", deployer)
    Check "params", "Feature has no plain En test", "Feature En test > A deployer", False, IsAllowed("Feature", "En test", deployer)
End Sub

Public Sub T_JiraConfig()
    ' The proposal sheets from fake Jira answers: 3 projects of 4 keys (two pages), one workflow shared by two projects
    ' (written once), a project whose issue types follow two workflows, a third project whose workflows the account may
    ' not read (401), both category enums, an error that leaves the proposal sheets as they were.
    Dim cfg As Object, keys As New Collection, base As String, summary As String, nm As Variant, e As Variant, wf As String
    Dim wfB As String, statusesB As String, rowsW As New Collection, rowsM As New Collection
    Dim wsT As Object, wsS As Object, wsW As Object, wsM As Object, wsF As Object, previews As String, raised As Long
    Dim many As New Collection, one As New Collection, i As Long, wsU As Object, wsP As Object, wsG As Object
    Set cfg = TestConfig()
    base = cfg("BaseUrl")
    keys.Add "PRG": keys.Add "TA": keys.Add "TB": keys.Add "XX"
    Application.DisplayAlerts = False
    For Each nm In Array("JiraIssueTypes", "JiraStatuses", "JiraWorkflows", "JiraTransitions", "JiraFields", "JiraProjectWorkflows", _
                         "JiraProjectFields", "JiraFieldUsage")
        If SheetExists(ThisWorkbook, CStr(nm)) Then ThisWorkbook.Worksheets(SheetName(CStr(nm))).Delete
    Next nm
    Application.DisplayAlerts = True
    wf = "{'statuses':[{'statusReference':'r1','name':'Open','statusCategory':'TODO'}," & _
         "{'statusReference':'r2','name':'Doing','statusCategory':'IN_PROGRESS'}," & _
         "{'statusReference':'r3','name':'Closed','statusCategory':'DONE'},{'statusReference':'r4','name':'Parked','statusCategory':'TODO'}" & _
         "MORE_STATUSES]," & _
         "'workflows':[{'name':'WF-A','queryContext':[{'project':'P','issueTypes':[TYPES]}]," & _
         "'statuses':[{'statusReference':'r1'},{'statusReference':'r2'},{'statusReference':'r3'},{'statusReference':'r4'}]," & _
         "'transitions':[{'type':'INITIAL','name':'Create','toStatusReference':'r1','links':[]}," & _
         "{'type':'DIRECTED','name':'Start','id':'11','toStatusReference':'r2','links':[{'fromStatusReference':'r1'},{'fromStatusReference':'r4'}]}," & _
         "{'type':'DIRECTED','name':'Close','toStatusReference':'r3','links':[{'fromStatusReference':'r2'},{'fromStatusReference':'zz'}]}," & _
         "{'type':'GLOBAL','name':'Park','toStatusReference':'r4','links':[]}]}MORE_WORKFLOWS]}"
    ' TA: its stories follow WF-A (also the workflow of PRG), its sub-tasks follow WF-B.
    statusesB = ",{'statusReference':'r5','name':'Triage','statusCategory':'TODO'},{'statusReference':'r6','name':'Fixed','statusCategory':'DONE'}"
    wfB = ",{'name':'WF-B','queryContext':[{'project':'P','issueTypes':['10005']}]," & _
          "'statuses':[{'statusReference':'r5'},{'statusReference':'r6'}]," & _
          "'transitions':[{'type':'INITIAL','name':'Create','toStatusReference':'r5','links':[]}," & _
          "{'type':'DIRECTED','name':'Fix','toStatusReference':'r6','links':[{'fromStatusReference':'r5'}]}]}"
    FakeReset
    FakeAdd "GET", ProjectSearchUrl(base, keys, 1, 0), 200, J("{'startAt':0,'isLast':false,'values':[{'id':'10100','key':'PRG'," & _
        "'style':'classic','issueTypes':[{'id':'10001','name':'Feature','subtask':false,'hierarchyLevel':1}]}]}")
    FakeAdd "GET", ProjectSearchUrl(base, keys, 1, 1), 200, J("{'startAt':1,'isLast':true,'values':[" & _
        "{'id':'10101','key':'TA','style':'next-gen','issueTypes':[{'id':'10002','name':'Story','subtask':false,'hierarchyLevel':0}," & _
        "{'id':'10005','name':'Sub-task','subtask':true,'hierarchyLevel':-1}]}," & _
        "{'id':'10102','key':'TB','style':'classic','issueTypes':[{'id':'10002','name':'Story','subtask':false,'hierarchyLevel':0}]}]}")
    FakeAdd "POST", base & "/rest/api/3/workflows/preview", 200, _
        J(Replace(Replace(Replace(wf, "TYPES", "'10001'"), "MORE_STATUSES", ""), "MORE_WORKFLOWS", ""))
    FakeAdd "POST", base & "/rest/api/3/workflows/preview", 200, _
        J(Replace(Replace(Replace(wf, "TYPES", "'10002'"), "MORE_STATUSES", statusesB), "MORE_WORKFLOWS", wfB))
    FakeAdd "POST", base & "/rest/api/3/workflows/preview", 401, J("{'errorMessages':['no permission']}")
    For Each nm In Array("PRG", "TA", "TB")
        FakeAdd "GET", StatusesUrl(base, CStr(nm)), 200, J("[{'id':'1','name':'Story','subtask':false,'statuses':[" & _
            "{'id':'5','name':'Doing','statusCategory':{'key':'indeterminate','name':'In progress'}}]}]")
    Next nm
    ' The create screens (createmeta, paged) and one page of issues per project (which fields are filled).
    FakeAdd "GET", CreateMetaUrl(base, "PRG", "10001", 0), 200, J("{'startAt':0,'total':1,'fields':[{'fieldId':'summary'," & _
        "'name':'Summary','required':true,'schema':{'type':'string'}}]}")
    FakeAdd "GET", CreateMetaUrl(base, "TA", "10002", 0), 200, J("{'startAt':0,'total':2,'fields':[{'fieldId':'labels'," & _
        "'name':'Labels','required':false,'schema':{'type':'array'}}]}")
    FakeAdd "GET", CreateMetaUrl(base, "TA", "10002", 1), 200, J("{'startAt':1,'total':2,'fields':[{'fieldId':'duedate'," & _
        "'name':'Due date','required':false,'schema':{'type':'date'}}]}")
    FakeAdd "GET", CreateMetaUrl(base, "TA", "10005", 0), 200, J("{'total':1,'results':[{'fieldId':'parent'," & _
        "'name':'Parent','required':true,'schema':{'type':'issuelink'}}]}")
    FakeAdd "GET", CreateMetaUrl(base, "TB", "10002", 0), 403, J("{'errorMessages':['no permission']}")
    For Each nm In Array("PRG", "TA", "TB")
        FakeAdd "GET", JiraSearchUrl(base, ProjectsJql(ColOf(CStr(nm))), "", JIRA_PAGE_SIZE, "*all", "names"), 200, _
            J("{'isLast':true,'names':{'summary':'Summary','labels':'Labels'},'issues':[" & _
              "{'key':'A-1','fields':{'summary':'s','labels':[],'duedate':null,'assignee':{'name':'x'}}}," & _
              "{'key':'A-2','fields':{'summary':'','labels':['l'],'duedate':null,'assignee':null}}]}")
    Next nm
    FakeAdd "GET", base & "/rest/api/3/field", 200, J("[{'id':'customfield_1','name':'Team','custom':true,'schema':{'type':'team'," & _
        "'custom':'com.atlassian.teams:rm-teams-custom-field-team'}},{'id':'status','name':'Status','custom':false," & _
        "'schema':{'type':'status','system':'status'}}]")
    summary = ReadJiraConfiguration(cfg, keys, ThisWorkbook)
    Set wsT = ThisWorkbook.Worksheets(SheetName("JiraIssueTypes")): Set wsS = ThisWorkbook.Worksheets(SheetName("JiraStatuses"))
    Set wsW = ThisWorkbook.Worksheets(SheetName("JiraWorkflows")): Set wsM = ThisWorkbook.Worksheets(SheetName("JiraTransitions"))
    Set wsF = ThisWorkbook.Worksheets(SheetName("JiraFields"))
    Set wsU = ThisWorkbook.Worksheets(SheetName("JiraProjectWorkflows")): Set wsP = ThisWorkbook.Worksheets(SheetName("JiraProjectFields"))
    Set wsG = ThisWorkbook.Worksheets(SheetName("JiraFieldUsage"))
    Check "jiraconfig", "a workflow shared by two projects names both", "WF-A", "WF-A|Open|Backlog|||0|TODO|1|PRG, TA", RowText(wsW, 2, 9)
    Check "jiraconfig", "a workflow of one project names it", "WF-B", "WF-B|Triage|Backlog|||0|TODO|1|TA", RowText(wsW, 6, 9)
    Check "jiraconfig", "a move has its id, its scope and its projects", "Start", "WF-A|Open|Doing|Start|11|DIRECTED|PRG, TA", RowText(wsM, 2, 7)
    Check "jiraconfig", "summary: a project, its workflow, its issue types, its counts", "PRG", "PRG|WF-A|Feature|4|4", RowText(wsU, 2, 5)
    Check "jiraconfig", "summary: the shared workflow is listed for each project", "TA", "TA|WF-A|Story|4|4", RowText(wsU, 3, 5)
    Check "jiraconfig", "summary: the second workflow of a project", "TA", "TA|WF-B|Sub-task|2|1", RowText(wsU, 4, 5)
    Check "jiraconfig", "summary: a workflow that could not be read stays visible", "TB", "TB||Story|0|0", RowText(wsU, 5, 5)
    Check "jiraconfig", "create screen of the first project", "PRG", "PRG|Feature|summary|Summary|Yes|string", RowText(wsP, 2, 6)
    Check "jiraconfig", "create screen read page by page", "TA Story", "TA|Story|labels|Labels|No|array;TA|Story|duedate|Due date|No|date", _
          RowText(wsP, 3, 6) & ";" & RowText(wsP, 4, 6)
    Check "jiraconfig", "create screen answered with results", "TA Sub-task", "TA|Sub-task|parent|Parent|Yes|issuelink", RowText(wsP, 5, 6)
    Check "jiraconfig", "no right on the create screen: no row, and the status says so", "TB", 5, LastRow(wsP)
    Check "jiraconfig", "the status tells the create screen could not be read", "403 on TB", True, _
          InStr(summary, " | Create screen not read: 1 project(s), ") > 0
    Check "jiraconfig", "fields filled: counted in the issues read", "summary", "PRG|summary|Summary|1|2", RowText(wsG, 2, 5)
    Check "jiraconfig", "fields filled: an empty list is not filled", "labels", "PRG|labels|Labels|1|2", RowText(wsG, 3, 5)
    Check "jiraconfig", "fields filled: null is not filled; a field without a name", "duedate", "PRG|duedate||0|2", RowText(wsG, 4, 5)
    Check "jiraconfig", "fields filled: an object is filled, null is not", "assignee", "PRG|assignee||1|2", RowText(wsG, 5, 5)
    Check "jiraconfig", "fields filled: 4 fields for each of 3 projects", "rows", 13, LastRow(wsG)
    Check "jiraconfig", "summary counts", "3 of 4 keys, 2 workflows, 2 fields", True, _
          InStr(summary, "Jira configuration read | Projects: 3 | Workflows: 2 | Fields: 2 | ") = 1
    Check "jiraconfig", "a key that Jira does not know is told", "XX", True, InStr(summary, " | Not found in Jira: 1 project key(s) of the sheet Projects") > 0
    Check "jiraconfig", "a project whose workflows cannot be read is told", "401 on TB", True, InStr(summary, " | Workflows not read: 1 project(s), ") > 0
    Check "jiraconfig", "issue type of the first page", "PRG", "PRG|Feature|WF-A|No|1|classic", RowText(wsT, 2, 6)
    Check "jiraconfig", "one project, two workflows: the story", "TA Story", "TA|Story|WF-A|No|0|next-gen", RowText(wsT, 3, 6)
    Check "jiraconfig", "one project, two workflows: the sub-task", "TA Sub-task", "TA|Sub-task|WF-B|Yes|-1|next-gen", RowText(wsT, 4, 6)
    Check "jiraconfig", "no workflow when it could not be read", "TB Story", "TB|Story||No|0|classic", RowText(wsT, 5, 6)
    Check "jiraconfig", "a workflow shared by two projects is written once", "header + WF-A 4 + WF-B 2", 7, LastRow(wsW)
    Check "jiraconfig", "the second workflow of a project has its own statuses", "WF-B", _
          "WF-B|Triage|Backlog|||0|TODO|1;WF-B|Fixed|Done||1|0|DONE|", RowText(wsW, 6, 8) & ";" & RowText(wsW, 7, 8)
    Check "jiraconfig", "status of creation: Initial 1; to do -> Backlog", "Open", "WF-A|Open|Backlog|||0|TODO|1", RowText(wsW, 2, 8)
    Check "jiraconfig", "in progress: Active or Wait is left to a human; the start is the first status in progress from the creation", "Doing", "WF-A|Doing||1||0|IN_PROGRESS|", RowText(wsW, 3, 8)
    Check "jiraconfig", "done -> Done, and the end", "Closed", "WF-A|Closed|Done||1|0|DONE|", RowText(wsW, 4, 8)
    Check "jiraconfig", "a GLOBAL transition makes its status free", "Parked", "WF-A|Parked|Backlog|||1|TODO|", RowText(wsW, 5, 8)
    Check "jiraconfig", "one move per link of a DIRECTED transition", "Start from Open and Parked", _
          "WF-A|Open|Doing|Start;WF-A|Parked|Doing|Start", RowText(wsM, 2, 4) & ";" & RowText(wsM, 3, 4)
    Check "jiraconfig", "a reference the answer does not list is shown, not hidden", "zz", "WF-A|zz|Closed|Close", RowText(wsM, 5, 4)
    Check "jiraconfig", "the moves of the second workflow follow", "WF-B", "WF-B|Triage|Fixed|Fix", RowText(wsM, 6, 4)
    Check "jiraconfig", "4 + 1 moves, the shared workflow written once", "rows", 6, LastRow(wsM)
    Check "jiraconfig", "statuses per issue type with the category key", "PRG", "PRG|Story|Doing|indeterminate", RowText(wsS, 2, 4)
    Check "jiraconfig", "one status row per project", "3 projects", 4, LastRow(wsS)
    Check "jiraconfig", "a custom field: id, name, custom, type, custom type", "Team", _
          "customfield_1|Team|Yes|team|com.atlassian.teams:rm-teams-custom-field-team", RowText(wsF, 2, 5)
    Check "jiraconfig", "a system field shows its system name", "status", "status|Status|No|status|status", RowText(wsF, 3, 5)
    For Each e In gHttpLog
        If InStr(e(1), "/workflows/preview") > 0 Then previews = previews & e(2) & ";"
    Next e
    Check "jiraconfig", "one preview per project, with its id and its issue types", "request bodies", _
          J("{'projectId':'10100','issueTypeIds':['10001']};{'projectId':'10101','issueTypeIds':['10002','10005']};" & _
            "{'projectId':'10102','issueTypeIds':['10002']};"), previews
    ' An error after some answers: nothing is written, the proposals of the run before stay.
    FakeReset
    FakeAdd "GET", ProjectSearchUrl(base, keys, 1, 0), 200, J("{'isLast':true,'values':[{'id':'10100','key':'PRG','issueTypes':[]}]}")
    FakeAdd "GET", StatusesUrl(base, "PRG"), 400, J("{'errorMessages':['bad']}")
    On Error Resume Next
    ReadJiraConfiguration cfg, keys, ThisWorkbook
    raised = Err.Number - vbObjectError
    On Error GoTo 0
    Check "jiraconfig", "an error on the statuses call stops the run", "HTTP 400", 602, raised
    Check "jiraconfig", "an error leaves the proposal sheets as they were", "JiraWorkflows", 7, LastRow(wsW)
    On Error Resume Next
    ReadJiraConfiguration cfg, New Collection, ThisWorkbook
    raised = Err.Number - vbObjectError
    On Error GoTo 0
    Check "jiraconfig", "no project in the Projects sheet is refused", "no key", 620, raised
    ' A project/search that never says isLast stops after 100 pages.
    one.Add "K"
    FakeReset
    For i = 0 To 100
        FakeAdd "GET", ProjectSearchUrl(base, one, 1, i), 200, J("{'isLast':false,'values':[{'key':'K'}]}")
    Next i
    On Error Resume Next
    ReadProjects cfg, one
    raised = Err.Number - vbObjectError
    On Error GoTo 0
    Check "jiraconfig", "paging is bounded", "isLast never true", 621, raised
    FakeOff
    ' Pure helpers: the two category enums, the 50 keys and 25 issue types per call, the rights.
    Check "jiraconfig", "category key new -> Backlog", "new", "Backlog", ProposedCategory("new")
    Check "jiraconfig", "category key indeterminate -> human choice", "indeterminate", "", ProposedCategory("indeterminate")
    Check "jiraconfig", "category key done -> Done", " done ", "Done", ProposedCategory(" done ")
    For i = 1 To 51
        many.Add "K" & i
    Next i
    Check "jiraconfig", "keys 51 and up go to the next call", "51 keys", _
          base & "/rest/api/3/project/search?keys=K51&expand=issueTypes&startAt=0&maxResults=50", ProjectSearchUrl(base, many, 51, 0)
    Check "jiraconfig", "a key is URL-encoded", "A B", True, InStr(ProjectSearchUrl(base, ColOf("A B"), 1, 0), "keys=A%20B&") > 0
    Check "jiraconfig", "issue types 26 and up go to the next preview", "27 ids", J("{'projectId':'7','issueTypeIds':['K26','K27']}"), _
          PreviewBody("7", Head(many, 27), 26)
    Check "jiraconfig", "401, 403 and 404 mean no right to read the workflows; 500 does not", "401|403|404|500", _
          "1110", IIf(IsNotAllowed(401), 1, 0) & IIf(IsNotAllowed(403), 1, 0) & IIf(IsNotAllowed(404), 1, 0) & IIf(IsNotAllowed(500), 1, 0)
    ' Workflows are told apart by id: two project-scoped workflows with the same name are both written, the same id once.
    AddWorkflowRows JsonParse(J("{'statuses':[{'statusReference':'a','name':'S'}],'workflows':[" & _
        "{'id':'1','name':'W','statuses':[{'statusReference':'a'}]},{'id':'2','name':'W','statuses':[{'statusReference':'a'}]}," & _
        "{'id':'1','name':'W','statuses':[{'statusReference':'a'}]}]}")), CreateObject("Scripting.Dictionary"), _
        CreateObject("Scripting.Dictionary"), rowsW, rowsM
    Check "jiraconfig", "same name, two ids: both written; the same id once", "ids 1, 2, 1", 2, rowsW.Count
End Sub

Private Function ColOf(ParamArray items() As Variant) As Collection
    Dim c As New Collection, item As Variant
    For Each item In items
        c.Add item
    Next item
    Set ColOf = c
End Function

Private Function Head(ByVal items As Collection, ByVal n As Long) As Collection
    Dim c As New Collection, i As Long
    For i = 1 To n
        c.Add items(i)
    Next i
    Set Head = c
End Function

' ================= Fake Jira (localhost; started by the Python tests, never part of the offline run) =================
' The M_ tests talk to the fake Jira of src\jira2dashboard\mockjira. Python calls MockSetup with its URL and the
' sizes of its dataset, then RunMockTests. Faults (429, 503) are requested over HTTP: POST /__mock__/fail.

Public Sub MockSetup(ByVal baseUrl As String, ByVal email As String, ByVal token As String, _
                     ByVal issueCount As Long, ByVal changeCount As Long)
    mMockUrl = baseUrl: mMockEmail = email: mMockToken = token
    mMockIssues = issueCount: mMockChanges = changeCount
End Sub

Public Function RunMockTests() As String
    ' M_JiraConfig resets the fake Jira (its faults): it runs before the proxy tests, whose requests a scenario reads after.
    RunMockTests = RunTests(Array("M_SearchAll", "M_Changelogs", "M_FlowFromMock", "M_Retry429", "M_ErrorsNotRetried", _
                                  "M_JiraConfig", "M_Credentials", "M_ThroughProxy", "M_ProxyWithoutWindowsAuth", _
                                  "M_AutoProxyPac", "M_FilterFromConfig"))
End Function

Private Function MockConfig() As Object
    Dim c As Object
    Set c = JiraConfig(mMockUrl, mMockEmail, mMockToken, "", "", False)
    c("RetryBaseMs") = 5: c("RetryMaxMs") = 20              ' retries must not slow the test
    Set MockConfig = c
End Function

Private Function MockProjects() As Collection
    ' The projects come from the Projects sheet of the workbook, the Program first (the fake Jira holds them).
    Set MockProjects = ProjectKeys(ThisWorkbook.Worksheets(SheetName("Projects")))
End Function

Private Sub MockFault(ByVal cfg As Object, ByVal path As String, ByVal status As Long, ByVal times As Long, _
                      Optional ByVal method As String = "GET")
    Dim body As String
    body = "{""method"":""" & method & """,""path"":""" & path & """,""status"":" & status & ",""times"":" & times & ",""retryAfter"":""0""}"
    HttpRequest cfg, "POST", mMockUrl & "/__mock__/fail", body
End Sub

Private Sub MockReset(ByVal cfg As Object)
    HttpRequest cfg, "POST", mMockUrl & "/__mock__/reset", "{}"
End Sub

Private Function MockSearch(ByVal cfg As Object, ByRef names As Object) As Collection
    Set MockSearch = JiraSearchAll(cfg, ProjectsJql(MockProjects()), names)
End Function

Public Sub M_SearchAll()
    Dim cfg As Object, names As Object, issues As Collection, issue As Variant, keys As Object, pages As Long, e As Variant
    Set cfg = MockConfig()
    MockReset cfg
    Set gHttpLog = New Collection
    Set issues = MockSearch(cfg, names)
    Set keys = CreateObject("Scripting.Dictionary")
    For Each issue In issues
        keys(JGet(issue, "key")) = True
    Next issue
    For Each e In gHttpLog
        If InStr(e(1), "/search/jql") > 0 Then pages = pages + 1
    Next e
    Check "mock", "every issue is read", "all pages", mMockIssues, issues.Count
    Check "mock", "no duplicate key", "unique keys", mMockIssues, keys.Count
    Check "mock", "pages = issues / 100 rounded up", "page size 100", Int((mMockIssues + 99) / 100), pages
    Check "mock", "field names come with expand=names", "status", "Status", names("status")
    Check "mock", "custom field name", "customfield_10016", "Story point estimate", names("customfield_10016")
    Check "mock", "status name is read", "first issue", True, Len(JPath(issues(1), "fields.status.name")) > 0
    Check "mock", "created is a date (offset converted)", "first issue", True, VarType(ParseJiraDate(JPath(issues(1), "fields.created"))) = vbDate
    Check "mock", "unresolved issue has a null resolution", "an open issue", True, FirstOpenHasNullResolution(issues)
    Check "mock", "only the Program project holds Features", "type vs project", 0, RoleMismatches(issues)
End Sub

Private Function RoleMismatches(ByVal issues As Collection) As Long
    ' Features are in the Program project of the Projects sheet and nowhere else; the Program holds nothing but Features.
    Dim issue As Variant, program As String
    program = ProgramProject(ThisWorkbook.Worksheets(SheetName("Projects")))
    For Each issue In issues
        If (JPath(issue, "fields.issuetype.name") = "Feature") <> (JPath(issue, "fields.project.key") = program) Then RoleMismatches = RoleMismatches + 1
    Next issue
End Function

Public Sub M_FilterFromConfig()
    ' The Filter of the Config sheet is a JQL restriction; the fake Jira applies it, so only Features come back.
    Dim cfg As Object, names As Object, issues As Collection, issue As Variant, notFeature As Long
    Set cfg = MockConfig()
    Set issues = JiraSearchAll(cfg, ProjectsJql(MockProjects(), "issuetype = Feature AND updated >= -3650d"), names)
    For Each issue In issues
        If JPath(issue, "fields.issuetype.name") <> "Feature" Then notFeature = notFeature + 1
    Next issue
    Check "mock", "the filter keeps Features only", "issuetype = Feature", 0, notFeature
    Check "mock", "some Features come back", "count", True, issues.Count > 0
    Check "mock", "fewer issues than without the filter", "count", True, issues.Count < mMockIssues
    Check "mock", "a filter the fake Jira cannot read is a 400", "assignee = x", 400, _
          HttpRequest(cfg, "GET", JiraSearchUrl(mMockUrl, ProjectsJql(MockProjects(), "assignee = x"), ""))("Status")
End Sub

Private Function FirstOpenHasNullResolution(ByVal issues As Collection) As Boolean
    Dim issue As Variant
    For Each issue In issues
        If JPath(issue, "fields.status.name") = "Nouveau" Then
            FirstOpenHasNullResolution = IsNull(JPath(issue, "fields.resolution")) Or IsEmpty(JPath(issue, "fields.resolution"))
            Exit Function
        End If
    Next issue
End Function

Public Sub M_Changelogs()
    Dim cfg As Object, names As Object, issues As Collection, keys As New Collection, issue As Variant
    Dim rows As Collection, r As Variant, bad As Long
    Set cfg = MockConfig()
    Set issues = MockSearch(cfg, names)
    For Each issue In issues
        keys.Add JGet(issue, "key")
    Next issue
    Set rows = JiraChangelogsAll(cfg, keys)
    For Each r In rows
        If r(5) <> "status" Or Len(CStr(r(11) & "")) = 0 Then bad = bad + 1
    Next r
    Check "mock", "every status change is read", "all issues", mMockChanges, rows.Count
    Check "mock", "rows are status changes with a destination", "fieldId and toString", 0, bad
    Check "mock", "issueId is an id, not a key", "first row", True, IsNumeric(rows(1)(0))
End Sub

Public Sub M_FlowFromMock()
    ' E2 + E3 together: categories from the Workflows sheet, issues and changes from the fake Jira, then intervals.
    Dim cfg As Object, names As Object, issues As Collection, keys As New Collection, issue As Variant
    Dim info As Object, idToKey As Object, changes As Collection, cats As Object, out As Collection, iv As Variant
    Dim bad As Long, open_ As Long
    Set cfg = MockConfig()
    Set issues = MockSearch(cfg, names)
    Set info = CreateObject("Scripting.Dictionary")
    Set idToKey = CreateObject("Scripting.Dictionary")
    For Each issue In issues
        info(JGet(issue, "key")) = Array(ParseJiraDate(JPath(issue, "fields.created")), JPath(issue, "fields.status.name"))
        idToKey(CStr(JGet(issue, "id"))) = JGet(issue, "key")
        keys.Add JGet(issue, "key")
    Next issue
    Set changes = StatusChangesFromChangelog(JiraChangelogsAll(cfg, keys), idToKey)
    Set cats = StatusCategories(ThisWorkbook.Worksheets(SheetName("Workflows")), "Team")
    Set out = BuildStatusIntervals(info, changes, DateSerial(2026, 11, 3), cats)
    For Each iv In out
        If Len(CStr(iv(9) & "")) > 0 Then bad = bad + 1
        If iv(7) Then open_ = open_ + 1
    Next iv
    Check "mock", "changes are keyed by issue key", "status changes", mMockChanges, changes.Count
    Check "mock", "one interval per change plus one open interval per issue", "intervals", mMockChanges + mMockIssues, out.Count
    Check "mock", "exactly one open interval per issue", "open intervals", mMockIssues, open_
    Check "mock", "search and changelog agree (no mismatch flag)", "check column", 0, bad
End Sub

Public Sub M_Retry429()
    Dim cfg As Object, resp As Object
    Set cfg = MockConfig()
    MockReset cfg
    MockFault cfg, "/rest/api/3/field", 429, 2
    Set resp = HttpRequest(cfg, "GET", mMockUrl & "/rest/api/3/field")
    Check "mock", "429 twice then success", "Retry-After 0", 200, resp("Status")
    Check "mock", "three attempts were needed", "2 faults", 3, resp("Attempts")
    MockFault cfg, "/rest/api/3/field", 429, 10
    cfg("MaxRetries") = 2
    Set resp = HttpRequest(cfg, "GET", mMockUrl & "/rest/api/3/field")
    Check "mock", "retries are bounded", "MaxRetries 2", 429, resp("Status")
    Check "mock", "bounded attempts", "MaxRetries 2", 3, resp("Attempts")
    MockReset cfg
End Sub

Public Sub M_ErrorsNotRetried()
    Dim cfg As Object, resp As Object
    Set cfg = MockConfig()
    MockReset cfg
    Set resp = HttpRequest(cfg, "GET", JiraSearchUrl(mMockUrl, "ORDER BY key ASC", ""))
    Check "mock", "unbounded JQL is a 400", "ORDER BY only", 400, resp("Status")
    Check "mock", "400 is not retried", "attempts", 1, resp("Attempts")
    Check "mock", "the message says why", "text", True, InStr(resp("Text"), "Unbounded") > 0
    Set resp = HttpRequest(cfg, "GET", mMockUrl & "/rest/api/3/search?jql=project%3DTA")
    Check "mock", "the removed search is a 410", "/rest/api/3/search", 410, resp("Status")
    Check "mock", "410 is not retried", "attempts", 1, resp("Attempts")
    On Error Resume Next
    CheckStatus resp, "search"
    Check "mock", "CheckStatus raises with the status", "410", True, InStr(Err.Description, "HTTP 410") > 0
    Err.Clear
    On Error GoTo 0
End Sub

Public Sub M_Credentials()
    ' Jira Basic auth: base64("email:token"). The PowerShell script trims the token; so must the client.
    Dim cfg As Object, pasted As Object, resp As Object
    Set cfg = MockConfig()
    Set resp = HttpRequest(cfg, "GET", mMockUrl & "/rest/api/3/myself")
    Check "mock", "the account is accepted", "email + token", 200, resp("Status")
    Set pasted = JiraConfig(mMockUrl, " " & mMockEmail & " ", mMockToken & vbCrLf, "", "", False)
    Set resp = HttpRequest(pasted, "GET", mMockUrl & "/rest/api/3/myself")
    Check "mock", "a token pasted with a line break still works", "token + CRLF", 200, resp("Status")
    pasted("Authorization") = BasicAuthHeader(mMockEmail, mMockToken & vbCrLf)      ' what an untrimmed client would send
    Set resp = HttpRequest(pasted, "GET", mMockUrl & "/rest/api/3/myself")
    Check "mock", "an untrimmed token is refused like Jira does", "token + CRLF, no trim", 401, resp("Status")
    Check "mock", "401 is plain text, not JSON", "body", "Unauthorized", resp("Text")
    Check "mock", "401 is not retried", "attempts", 1, resp("Attempts")
    pasted("Authorization") = BasicAuthHeader(mMockEmail, "wrong")
    Set resp = HttpRequest(pasted, "GET", mMockUrl & "/rest/api/3/myself")
    Check "mock", "a wrong token is refused", "wrong", 401, resp("Status")
End Sub

Private Function ProxyConfig(ByVal windowsAuth As String) As Object
    ' The connection comes from a Config sheet, as in production. BaseUrl is the Jira site, ProxyServer is the fake,
    ' typed with http:// like in the PowerShell script. The site name cannot be resolved, so nothing can reach a real
    ' server: an answer proves that the request went through the proxy setting.
    Dim c As Object
    ConfigRows ThisWorkbook.Worksheets("TestScratch"), Array("BaseUrl", "http://jira.mock.invalid", _
        "Email", mMockEmail, "ApiToken", mMockToken & vbCrLf, "ProxyServer", mMockUrl, _
        "ProxyWindowsAuth", windowsAuth, "TimeoutSeconds", "30", "MaxRetries", "2")
    Set c = JiraConfigFromSheet(ThisWorkbook.Worksheets("TestScratch"))
    c("RetryBaseMs") = 5: c("RetryMaxMs") = 20
    Set ProxyConfig = c
End Function

Public Sub M_ThroughProxy()
    ' The proxy asks for Windows authentication (407 until the request carries Proxy-Authorization). With
    ' ProxyWindowsAuth = Yes, WinHTTP answers the challenge with the Windows session and the request goes through.
    Dim cfg As Object, resp As Object, names As Object, issues As Collection
    Set cfg = ProxyConfig("Yes")
    Check "mock", "the proxy address typed with http:// is accepted", "ProxyServer", Mid$(mMockUrl, 8), cfg("ProxyServer")
    Set resp = HttpRequest(cfg, "GET", cfg("BaseUrl") & "/rest/api/3/myself")
    Check "mock", "the site is reached through the proxy", "GET /myself", 200, resp("Status")
    Set issues = JiraSearchAll(cfg, ProjectsJql(MockProjects()), names)
    Check "mock", "a full search works through the proxy", "all pages", mMockIssues, issues.Count
End Sub

Public Sub M_AutoProxyPac()
    ' ProxyServer empty, and the Windows settings give a proxy script (PAC) served by the fake Jira that sends every
    ' address to the fake as a proxy. The site name cannot be resolved: an answer proves that the detected proxy is used.
    Dim c As Object, resp As Object
    Set gProxyFake = FakeProxySettings(False, mMockUrl & "/proxy.pac", "", "")
    ConfigRows ThisWorkbook.Worksheets("TestScratch"), Array("BaseUrl", "http://jira.mock.invalid", _
        "Email", mMockEmail, "ApiToken", mMockToken, "ProxyServer", "", _
        "ProxyWindowsAuth", "Yes", "TimeoutSeconds", "30", "MaxRetries", "2")
    Set c = JiraConfigFromSheet(ThisWorkbook.Worksheets("TestScratch"))
    c("RetryBaseMs") = 5: c("RetryMaxMs") = 20
    Check "mock", "the proxy of the PAC script", "ProxyServer empty", Mid$(mMockUrl, 8), c("ProxyServer")
    Check "mock", "found by the script", "PAC", "proxy script (PAC): a proxy", c("ProxyMode")
    Set resp = HttpRequest(c, "GET", c("BaseUrl") & "/rest/api/3/myself")
    Check "mock", "the site is reached through the detected proxy", "GET /myself", 200, resp("Status")
End Sub

Public Sub M_ProxyWithoutWindowsAuth()
    ' ProxyWindowsAuth = No: the Windows session is never sent, so a proxy that wants it answers 407, and the
    ' error message says which Config cells to check. A 407 is not retried.
    Dim cfg As Object, resp As Object, msg As String
    Set cfg = ProxyConfig("No")
    Set resp = HttpRequest(cfg, "GET", cfg("BaseUrl") & "/rest/api/3/myself")
    Check "mock", "no Windows authentication gives 407", "ProxyWindowsAuth No", 407, resp("Status")
    Check "mock", "407 is not retried", "attempts", 1, resp("Attempts")
    On Error Resume Next
    CheckStatus resp, "myself"
    msg = Err.Description: Err.Clear
    On Error GoTo 0
    Check "mock", "the message points to the Config cells", "HTTP 407", True, InStr(msg, "ProxyWindowsAuth") > 0
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
    CovWrite ThisWorkbook.Worksheets(SheetName("CoverageHits"))
    mTotalPlanned = 0
    RunWithCoverage = s
End Function

Public Sub M_JiraConfig()
    ' Round trip: the fake Jira is built from the parameters of the repo, so the proposals give them back (every status of
    ' the sheet Workflows with its free flag, every move of the sheet Transitions); a 401 on the preview keeps the rest.
    Dim cfg As Object, summary As String, row As Variant, want As Object, got As Object, program As String, wrong As Long
    Set cfg = MockConfig()
    MockReset cfg
    summary = ReadJiraConfiguration(cfg, MockProjects(), ThisWorkbook)
    Check "mock", "the configuration is read", "summary", True, _
          InStr(summary, "Jira configuration read | Projects: " & MockProjects().Count & " | Workflows: 2 | ") = 1
    Set want = CreateObject("Scripting.Dictionary"): Set got = CreateObject("Scripting.Dictionary")
    For Each row In ReadTable(ThisWorkbook.Worksheets(SheetName("Workflows")))
        want(row("Workflow") & "|" & row("Status") & "|" & row("IsFree")) = True
    Next row
    For Each row In ReadTable(ThisWorkbook.Worksheets(SheetName("JiraWorkflows")))
        got(row("Workflow") & "|" & row("Status") & "|" & row("IsFree")) = True
    Next row
    Check "mock", "JiraWorkflows gives back the statuses and the free flags of Workflows", "Workflow|Status|IsFree", _
          SortedKeys(want), SortedKeys(got)
    Set want = CreateObject("Scripting.Dictionary"): Set got = CreateObject("Scripting.Dictionary")
    For Each row In ReadTable(ThisWorkbook.Worksheets(SheetName("Transitions")))
        want(row("Workflow") & "|" & row("From") & "|" & row("To")) = True
    Next row
    For Each row In ReadTable(ThisWorkbook.Worksheets(SheetName("JiraTransitions")))
        got(row("Workflow") & "|" & row("From") & "|" & row("To")) = True
    Next row
    Check "mock", "JiraTransitions gives back the moves of Transitions", "Workflow|From|To", SortedKeys(want), SortedKeys(got)
    program = ProgramProject(ThisWorkbook.Worksheets(SheetName("Projects")))
    For Each row In ReadTable(ThisWorkbook.Worksheets(SheetName("JiraIssueTypes")))
        If (row("Project") = program) <> (row("Workflow") = "Feature") Or Len(row("Workflow") & "") = 0 Then wrong = wrong + 1
    Next row
    Check "mock", "the Program types follow Feature, the team types Team", "JiraIssueTypes", 0, wrong
    MockFault cfg, "/rest/api/3/workflows/preview", 401, 100, "POST"
    summary = ReadJiraConfiguration(cfg, MockProjects(), ThisWorkbook)
    Check "mock", "without the right to read workflows the run says so", "401 on the preview", True, _
          InStr(summary, " | Workflows not read: " & MockProjects().Count & " project(s), ") > 0
    Check "mock", "and writes no workflow", "JiraWorkflows", 1, LastRow(ThisWorkbook.Worksheets(SheetName("JiraWorkflows")))
    Check "mock", "but the statuses of each issue type", "JiraStatuses", True, LastRow(ThisWorkbook.Worksheets(SheetName("JiraStatuses"))) > 1
    MockReset cfg
End Sub

Private Function SortedKeys(ByVal d As Object) As String
    ' The keys of a Dictionary, sorted and joined with ";" (two sets compare as two texts).
    Dim k As Variant, items() As String, i As Long, j As Long, t As String
    If d.Count = 0 Then Exit Function
    ReDim items(0 To d.Count - 1)
    For Each k In d.Keys
        items(i) = CStr(k): i = i + 1
    Next k
    For i = 1 To UBound(items)
        t = items(i): j = i - 1
        Do While j >= 0
            If items(j) <= t Then Exit Do
            items(j + 1) = items(j): j = j - 1
        Loop
        items(j + 1) = t
    Next i
    SortedKeys = Join(items, ";")
End Function

' ---------- Moved from the delivered modules: no caller there, the tests below still use them ----------

' (from modJira)
Public Function JiraSearchAll(ByVal cfg As Object, ByVal jql As String, ByRef names As Object, _
                              Optional ByVal maxPages As Long = MAX_SEARCH_PAGES, _
                              Optional ByVal fields As String = "*all", Optional ByVal expand As String = "names") As Collection
    ' Reads every page. names receives fieldId -> display name (from expand=names).
    Cov "JIR06"
    Dim issues As New Collection, token As String, page As Long, isLast As Boolean
    If names Is Nothing Then Set names = CreateObject("Scripting.Dictionary")
    Do
        page = page + 1
        token = JiraSearchPage(cfg, jql, token, issues, names, fields, expand, isLast)
        Application.StatusBar = "Jira search: " & issues.Count & " issues read (page " & page & ")"
        If isLast Then Exit Do
        If page >= maxPages Then Cov "JIR07": Err.Raise vbObjectError + 601, "JiraSearchAll", "Jira search did not end after " & maxPages & " pages: narrow the extraction with the Filter of the Config sheet."
    Loop
    Application.StatusBar = False
    Set JiraSearchAll = issues
End Function

' (from modJira)
Public Function JiraChangelogsAll(ByVal cfg As Object, ByVal issueKeys As Collection, _
                                  Optional ByVal chunkSize As Long = CHANGELOG_CHUNK) As Collection
    ' One row per changed item: Array(issueId, changeId, created, authorName, authorId,
    '                                 fieldId, field, fieldtype, from, fromString, to, toString)
    Cov "JIR10"
    Dim rows As New Collection, nextFirst As Long
    nextFirst = 1
    Do While nextFirst <= issueKeys.Count
        nextFirst = JiraChangelogChunk(cfg, issueKeys, nextFirst, rows, chunkSize)
        Application.StatusBar = "Jira changelogs: " & (nextFirst - 1) & " / " & issueKeys.Count & " issues"
    Loop
    Application.StatusBar = False
    Set JiraChangelogsAll = rows
End Function

' (from modJira)
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

' (from modJira)
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

' (from modJira)
Public Function ExtractionJqlFromWorkbook() As String
    ExtractionJqlFromWorkbook = ExtractionJql(ThisWorkbook.Worksheets(SheetName("Projects")), ThisWorkbook.Worksheets(SheetName("Config")))
End Function

' (from modParams)
Public Function FlaggedStatus(ByVal wsWorkflows As Object, ByVal workflow As String, ByVal flagColumn As String) As String
    ' First status of the workflow whose flag column (IsStart, IsEnd or IsFree) is 1. Empty text if none.
    Cov "PRM04"
    Dim row As Variant
    For Each row In ReadTable(wsWorkflows)
        If CStr(row("Workflow")) = workflow Then
            If Val(CStr(row(flagColumn) & "")) = 1 Then Cov "PRM05": FlaggedStatus = CStr(row("Status")): Exit Function
        End If
    Next row
End Function

' (from modParams)
Public Function TransitionAllowed(ByVal wsWorkflows As Object, ByVal wsTransitions As Object, ByVal workflow As String, _
                                  ByVal fromStatus As String, ByVal toStatus As String) As Boolean
    ' Allowed if a row From -> To exists, or if the target is a free status (IsFree = 1: reachable from any status).
    Cov "PRM07"
    Dim row As Variant
    For Each row In ReadTable(wsTransitions)
        If CStr(row("Workflow")) = workflow And CStr(row("From")) = fromStatus And CStr(row("To")) = toStatus Then
            Cov "PRM08": TransitionAllowed = True: Exit Function
        End If
    Next row
    If StatusHasFlag(wsWorkflows, workflow, toStatus, "IsFree") Then Cov "PRM09": TransitionAllowed = True
End Function

' (from modParams)
Public Function TeamOfProject(ByVal wsProjects As Object, ByVal projectKey As String) As String
    ' Team name of a project. Empty text for the Program and for an unknown project.
    Cov "PRM20"
    Dim row As Variant
    For Each row In ReadTable(wsProjects)
        If TrimWhite(CStr(row("Project"))) = projectKey Then TeamOfProject = TrimWhite(CStr(row("Team") & "")): Exit Function
    Next row
End Function

' (from modParams)
Public Function CategoryOf(ByVal workflow As String, ByVal status As String) As String
    ' Category of a status in this workbook's Workflows sheet. A status that is not listed is "Unclassified".
    Cov "PRM11"
    Dim cats As Object
    Set cats = StatusCategories(ThisWorkbook.Worksheets(SheetName(SHEET_WORKFLOWS)), workflow)
    If cats.Exists(status) Then CategoryOf = cats(status) Else Cov "PRM12": CategoryOf = "Unclassified"
End Function

' (from modParams)
Public Function IsAllowed(ByVal workflow As String, ByVal fromStatus As String, ByVal toStatus As String) As Boolean
    ' Transition check on this workbook's sheets Workflows and Transitions.
    Cov "PRM13"
    IsAllowed = TransitionAllowed(ThisWorkbook.Worksheets(SheetName(SHEET_WORKFLOWS)), ThisWorkbook.Worksheets(SheetName(SHEET_TRANSITIONS)), _
                                  workflow, fromStatus, toStatus)
End Function
