# Tested VBA modules and test harness

**Test status:** last run on 2026-11-03 on Windows 11, Python 3.14.7, pywin32 312,
Excel 16.0 Microsoft 365 (French UI). Result: project compiles (`COMPILED_OK`), **143 checks OK, 0 KO**
(135 offline, 8 network against httpbin.org), **coverage 65/65 markers = 100 %**.
Jira itself was **not** called: Jira behaviour is tested with recorded-shape fake responses.

Rules for Copilot:
- Extract these modules **mechanically** with `tools/cookbook_extract.py` (in `cookbook_python.md`); it checks
  every MD5. The generator reads the `.bas` files. Never retype or rewrite them.
- The source is pure ASCII on purpose. Keep the `\u00e9`-style escapes inside JSON fixtures as they are.
- Add new code in **new** modules (for example `modExtract`, `modChecks`, `modMetrics`) that call these.
- Every new procedure and every branch gets a unique `Cov "XXX00"` marker (see modCoverage).
- Module load order matters only for readability; VBA resolves names at compile time.

| Module | Purpose |
|---|---|
| `modCoverage` | Coverage markers. `Cov "ID"` costs almost nothing when coverage is off. |
| `modJson` | JSON parser and writer (objects -> Scripting.Dictionary, arrays -> Collection, null -> Null), safe accessors `JGet`, `JPath`, `AsList`, `IsTrue`. |
| `modText` | UTF-8, Base64, Basic auth header, URL encoding, Jira date parsing to UTC, any field value to a cell value (ADF rich text included). |
| `modHttp` | WinHttp with named proxy + Windows auth to the proxy, asynchronous wait with DoEvents, retry on 429/503 with Retry-After and jittered backoff, a fake for offline tests. |
| `modJira` | Jira Cloud: config, project JQL, `/rest/api/3/search/jql` pagination with `nextPageToken`, `POST /rest/api/3/changelog/bulkfetch` in chunks of 1000 with pagination, wide issue rows, fast block writing. |
| `modFlow` | Status intervals (traceable base of Flow Time / Load / Efficiency) with data checks FROM_MISMATCH, BEFORE_CREATED, CURRENT_MISMATCH. |
| `modBackground` | Option B: watch an export folder written by the user's own PowerShell script, then import the pages. Excel never starts a process. |
| `modButtons` | Button and Alt+F8 entry points: Public Subs without arguments, one call + a message. |
| `modTests` | The 29 offline tests + 4 network tests. Writes each TestResults row **as soon as its check runs** (TestId, Area, Description, Input, Expected, Actual, OK/KO, RunAt) and moves a progress bar (shapes `pbBack` / `pbFill`, label `J10`) after each test. |

Required workbook sheets for the tests: `TestResults`, `TestScratch`, `CoverageHits`, `CoverageMap`.
On `TestResults` the generator adds 3 buttons (J2, J4, J6) and the progress bar (J8, label J10); see `run_vba_tests.py`.
The tests also create folders `tmp_files` and `tmp_export` next to the workbook (never `%TEMP%`).

## Get the modules

From the project folder (see `cookbook_python.md` section 2):

```text
py -3.14 tools/cookbook_extract.py docs/cookbook_python.md docs/vba-modules-tested.md
```

It writes `vba/<module>.bas` (ASCII, LF line endings) and checks each MD5 given in the headings below.
The test harness (`run_vba_tests.py`, `_run_tests_child.py`) and the machine check are in `cookbook_python.md`.

## Modules

**modCoverage**: 37 lines, 0 coverage markers.

### `vba/modCoverage.bas`  (MD5 `0b03cf2cb8d0660ed22beadabc9655a4`)

```vba
Option Explicit
' VBA code coverage by markers. Each procedure and each branch calls Cov "<ID>" once.
' The Python generator lists every marker ID in the sheet "CoverageMap" at build time.
' Tests switch coverage on, run, then CovWrite fills "CoverageHits". A formula gives the percentage.

Public gCovOn As Boolean
Private mHits As Object

Public Sub Cov(ByVal markerId As String)
    If Not gCovOn Then Exit Sub                 ' near-zero cost in production
    If mHits Is Nothing Then Set mHits = CreateObject("Scripting.Dictionary")
    mHits(markerId) = mHits(markerId) + 1
End Sub

Public Sub CovStart()
    Set mHits = CreateObject("Scripting.Dictionary")
    gCovOn = True
End Sub

Public Sub CovStop()
    gCovOn = False
End Sub

Public Sub CovWrite(ByVal ws As Object)
    ' Writes MarkerId | Hits for every marker that was reached.
    Dim data() As Variant, k As Variant, r As Long
    ws.Cells.ClearContents
    ws.Range("A1:B1").Value = Array("MarkerId", "Hits")
    If mHits Is Nothing Then Exit Sub
    If mHits.Count = 0 Then Exit Sub
    ReDim data(1 To mHits.Count, 1 To 2)
    For Each k In mHits.Keys
        r = r + 1
        data(r, 1) = k: data(r, 2) = mHits(k)
    Next k
    ws.Range(ws.Cells(2, 1), ws.Cells(mHits.Count + 1, 2)).Value = data
End Sub
```

**modJson**: 260 lines, 18 coverage markers.

### `vba/modJson.bas`  (MD5 `9aa255d28cbabe546261339e4073a217`)

```vba
Option Explicit
' JSON parser and writer. No references needed (late binding).
' Objects -> Scripting.Dictionary (keys case-sensitive), arrays -> Collection, null -> Null.
' Numbers use Val(), which ignores the Windows locale (French uses a comma).

Private mText As String
Private mPos As Long
Private mLen As Long
Private mNextBackslash As Long

Public Function JsonParse(ByVal text As String) As Variant
    Cov "JSN01"
    mText = text: mPos = 1: mLen = Len(text): mNextBackslash = 0
    Dim result As Variant
    ParseValue result
    SkipWhitespace
    If mPos <= mLen Then Cov "JSN02": Fail "Unexpected text after JSON"
    If IsObject(result) Then Set JsonParse = result Else JsonParse = result
End Function

Private Sub ParseValue(ByRef target As Variant)
    SkipWhitespace
    If mPos > mLen Then Cov "JSN03": Fail "Unexpected end of JSON"
    Select Case Mid$(mText, mPos, 1)
        Case "{": Set target = ParseObject()
        Case "[": Set target = ParseArray()
        Case """": target = ParseString()
        Case "t": ExpectWord "true": target = True
        Case "f": ExpectWord "false": target = False
        Case "n": ExpectWord "null": target = Null
        Case Else: target = ParseNumber()
    End Select
End Sub

Private Function ParseObject() As Object
    Cov "JSN04"
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 0
    mPos = mPos + 1
    SkipWhitespace
    If Mid$(mText, mPos, 1) = "}" Then mPos = mPos + 1: Set ParseObject = d: Exit Function
    Do
        SkipWhitespace
        If Mid$(mText, mPos, 1) <> """" Then Cov "JSN05": Fail "Expected a key"
        Dim key As String
        key = ParseString()
        SkipWhitespace
        If Mid$(mText, mPos, 1) <> ":" Then Cov "JSN06": Fail "Expected ':'"
        mPos = mPos + 1
        ParseMember d, key
        SkipWhitespace
        Select Case Mid$(mText, mPos, 1)
            Case ",": mPos = mPos + 1
            Case "}": mPos = mPos + 1: Exit Do
            Case Else: Cov "JSN07": Fail "Expected ',' or '}'"
        End Select
    Loop
    Set ParseObject = d
End Function

Private Sub ParseMember(ByVal d As Object, ByVal key As String)
    Dim v As Variant
    ParseValue v
    If IsObject(v) Then Set d(key) = v Else d(key) = v
End Sub

Private Function ParseArray() As Collection
    Cov "JSN08"
    Dim c As New Collection
    mPos = mPos + 1
    SkipWhitespace
    If Mid$(mText, mPos, 1) = "]" Then mPos = mPos + 1: Set ParseArray = c: Exit Function
    Do
        ParseItem c
        SkipWhitespace
        Select Case Mid$(mText, mPos, 1)
            Case ",": mPos = mPos + 1
            Case "]": mPos = mPos + 1: Exit Do
            Case Else: Cov "JSN09": Fail "Expected ',' or ']'"
        End Select
    Loop
    Set ParseArray = c
End Function

Private Sub ParseItem(ByVal c As Collection)
    Dim v As Variant
    ParseValue v
    c.Add v
End Sub

Private Function ParseString() As String
    Dim out As String, q As Long, esc As String
    mPos = mPos + 1
    Do
        q = InStr(mPos, mText, """")
        If q = 0 Then Cov "JSN10": Fail "Unterminated string"
        If mNextBackslash < mPos And mNextBackslash <> -1 Then
            mNextBackslash = InStr(mPos, mText, "\")
            If mNextBackslash = 0 Then mNextBackslash = -1
        End If
        If mNextBackslash = -1 Or mNextBackslash > q Then
            ParseString = out & Mid$(mText, mPos, q - mPos)
            mPos = q + 1
            Exit Function
        End If
        Cov "JSN11"
        out = out & Mid$(mText, mPos, mNextBackslash - mPos)
        esc = Mid$(mText, mNextBackslash + 1, 1)
        mPos = mNextBackslash + 2
        Select Case esc
            Case """", "\", "/": out = out & esc
            Case "b": out = out & Chr$(8)
            Case "f": out = out & Chr$(12)
            Case "n": out = out & vbLf
            Case "r": out = out & vbCr
            Case "t": out = out & vbTab
            Case "u": out = out & ChrW$(CLng("&H" & Mid$(mText, mPos, 4))): mPos = mPos + 4
            Case Else: Cov "JSN12": Fail "Bad escape \" & esc
        End Select
    Loop
End Function

Private Function ParseNumber() As Double
    Dim startPos As Long
    startPos = mPos
    Do While mPos <= mLen
        If InStr("+-0123456789.eE", Mid$(mText, mPos, 1)) = 0 Then Exit Do
        mPos = mPos + 1
    Loop
    If mPos = startPos Then Cov "JSN13": Fail "Unexpected character '" & Mid$(mText, mPos, 1) & "'"
    ParseNumber = Val(Mid$(mText, startPos, mPos - startPos))
End Function

Private Sub ExpectWord(ByVal word As String)
    If Mid$(mText, mPos, Len(word)) <> word Then Cov "JSN14": Fail "Expected " & word
    mPos = mPos + Len(word)
End Sub

Private Sub SkipWhitespace()
    Do While mPos <= mLen
        Select Case AscW(Mid$(mText, mPos, 1))
            Case 32, 9, 10, 13: mPos = mPos + 1
            Case Else: Exit Do
        End Select
    Loop
End Sub

Private Sub Fail(ByVal message As String)
    Err.Raise vbObjectError + 513, "modJson", message & " at position " & mPos
End Sub

' ---------- Writer ----------

Public Function JsonStringify(ByVal value As Variant) As String
    If IsObject(value) Then
        If TypeName(value) = "Dictionary" Then
            Cov "JSN15": JsonStringify = StringifyObject(value)
        ElseIf TypeName(value) = "Collection" Then
            Cov "JSN16": JsonStringify = StringifyArray(value)
        Else
            Cov "JSN17": Err.Raise vbObjectError + 514, "modJson", "Cannot write " & TypeName(value)
        End If
    ElseIf IsNull(value) Or IsEmpty(value) Then
        JsonStringify = "null"
    ElseIf VarType(value) = vbBoolean Then
        JsonStringify = IIf(value, "true", "false")
    ElseIf IsNumeric(value) And VarType(value) <> vbString Then
        JsonStringify = Trim$(Str$(value))
    ElseIf IsArray(value) Then
        Cov "JSN18": JsonStringify = StringifyVbArray(value)
    Else
        JsonStringify = """" & JsonEscape(CStr(value)) & """"
    End If
End Function

Private Function StringifyObject(ByVal d As Object) As String
    Dim parts() As String, i As Long, k As Variant
    If d.Count = 0 Then StringifyObject = "{}": Exit Function
    ReDim parts(0 To d.Count - 1)
    For Each k In d.Keys
        parts(i) = """" & JsonEscape(CStr(k)) & """:" & JsonStringify(d(k))
        i = i + 1
    Next k
    StringifyObject = "{" & Join(parts, ",") & "}"
End Function

Private Function StringifyArray(ByVal c As Collection) As String
    Dim parts() As String, i As Long
    If c.Count = 0 Then StringifyArray = "[]": Exit Function
    ReDim parts(1 To c.Count)
    For i = 1 To c.Count
        parts(i) = JsonStringify(c(i))
    Next i
    StringifyArray = "[" & Join(parts, ",") & "]"
End Function

Private Function StringifyVbArray(ByVal a As Variant) As String
    Dim parts() As String, i As Long
    If UBound(a) < LBound(a) Then StringifyVbArray = "[]": Exit Function
    ReDim parts(LBound(a) To UBound(a))
    For i = LBound(a) To UBound(a)
        parts(i) = JsonStringify(a(i))
    Next i
    StringifyVbArray = "[" & Join(parts, ",") & "]"
End Function

Public Function JsonEscape(ByVal s As String) As String
    Dim i As Long, ch As String, code As Long, out As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCr, "\r")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)
        If code >= 0 And code < 32 Then out = out & "\u" & Right$("000" & Hex$(code), 4) Else out = out & ch
    Next i
    JsonEscape = out
End Function

' ---------- Safe access helpers ----------

Public Function JGet(ByVal node As Variant, ByVal key As String) As Variant
    ' Returns node(key), or Empty if node is not a Dictionary or the key is missing.
    JGet = Empty
    If Not IsObject(node) Then Exit Function
    If TypeName(node) <> "Dictionary" Then Exit Function
    If Not node.Exists(key) Then Exit Function
    If IsObject(node(key)) Then Set JGet = node(key) Else JGet = node(key)
End Function

Public Function JPath(ByVal node As Variant, ByVal path As String) As Variant
    ' JPath(issue, "fields.status.name"). Returns Empty when any step is missing.
    Dim parts() As String, i As Long, cur As Variant
    parts = Split(path, ".")
    If IsObject(node) Then Set cur = node Else cur = node
    For i = LBound(parts) To UBound(parts)
        If IsObject(JGet(cur, parts(i))) Then
            Set cur = JGet(cur, parts(i))
        Else
            cur = JGet(cur, parts(i))
        End If
    Next i
    If IsObject(cur) Then Set JPath = cur Else JPath = cur
End Function

Public Function AsList(ByVal value As Variant) As Collection
    ' Returns value if it is a Collection (JSON array), else an empty Collection. Safe for For Each.
    If IsObject(value) Then
        If TypeName(value) = "Collection" Then Set AsList = value: Exit Function
    End If
    Set AsList = New Collection
End Function

Public Function IsTrue(ByVal value As Variant) As Boolean
    ' True only for a real Boolean True. Safe with Null and Empty.
    If VarType(value) = vbBoolean Then IsTrue = value
End Function
```

**modText**: 144 lines, 15 coverage markers.

### `vba/modText.bas`  (MD5 `d1bc0e69487658ed6ec5e78903738c16`)

```vba
Option Explicit
' Text helpers: UTF-8, Base64, URL encoding, Jira dates, field values to cells.

Public Function Utf8Encode(ByVal text As String) As Byte()
    Cov "TXT01"
    Dim st As Object, empty_() As Byte
    If Len(text) = 0 Then Utf8Encode = empty_: Exit Function
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8": st.Open
    st.WriteText text
    st.Position = 0: st.Type = 1: st.Position = 3   ' skip the UTF-8 BOM
    Utf8Encode = st.Read
    st.Close
End Function

Public Function Utf8Decode(ByVal bytes As Variant) As String
    Cov "TXT02"
    Dim st As Object, upper As Long
    On Error Resume Next
    upper = UBound(bytes)                              ' fails on an uninitialised (empty) array
    If Err.Number <> 0 Then Err.Clear: Exit Function   ' expected: leave no stale error
    On Error GoTo 0
    If upper < LBound(bytes) Then Exit Function
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1: st.Open
    st.Write bytes
    st.Position = 0: st.Type = 2: st.Charset = "utf-8"
    Utf8Decode = st.ReadText
    st.Close
End Function

Public Function Base64Encode(ByVal bytes As Variant) As String
    Cov "TXT03"
    Dim node As Object
    Set node = CreateObject("MSXML2.DOMDocument.6.0").createElement("b64")
    node.DataType = "bin.base64"
    node.nodeTypedValue = bytes
    Base64Encode = Replace(Replace(node.Text, vbLf, ""), vbCr, "")
End Function

Public Function BasicAuthHeader(ByVal email As String, ByVal apiToken As String) As String
    Cov "TXT04"
    BasicAuthHeader = "Basic " & Base64Encode(Utf8Encode(email & ":" & apiToken))
End Function

Public Function UrlEncode(ByVal text As String) As String
    Cov "TXT05"
    Dim b() As Byte, i As Long, out As String
    If Len(text) = 0 Then Exit Function
    b = Utf8Encode(text)
    For i = LBound(b) To UBound(b)
        Select Case b(i)
            Case 48 To 57, 65 To 90, 97 To 122, 45, 46, 95, 126   ' 0-9 A-Z a-z - . _ ~
                out = out & Chr$(b(i))
            Case Else
                out = out & "%" & Right$("0" & Hex$(b(i)), 2)
        End Select
    Next i
    UrlEncode = out
End Function

Public Function ParseJiraDate(ByVal value As Variant) As Variant
    ' "2026-10-01T09:12:33.123+0200", "...+02:00", "...Z" -> Date in UTC.
    ' "2026-10-01" -> that date. Empty or invalid -> Empty (never raises).
    Dim s As String, d As Date, offsetMin As Long, tail As String, sign As Long
    ParseJiraDate = Empty
    If IsNull(value) Or IsEmpty(value) Then Cov "TXT06": Exit Function
    s = CStr(value)
    If Len(s) < 10 Then Cov "TXT07": Exit Function
    If Mid$(s, 5, 1) <> "-" Or Mid$(s, 8, 1) <> "-" Then Cov "TXT08": Exit Function
    d = DateSerial(CInt(Left$(s, 4)), CInt(Mid$(s, 6, 2)), CInt(Mid$(s, 9, 2)))
    If Len(s) = 10 Then Cov "TXT09": ParseJiraDate = d: Exit Function
    Cov "TXT10"
    d = d + TimeSerial(CInt(Mid$(s, 12, 2)), CInt(Mid$(s, 15, 2)), CInt(Mid$(s, 18, 2)))
    tail = Mid$(s, 20)
    Do While Len(tail) > 0 And InStr("+-Zz", Left$(tail, 1)) = 0
        tail = Mid$(tail, 2)                              ' skip ".123"
    Loop
    If Len(tail) > 0 And UCase$(Left$(tail, 1)) <> "Z" Then
        sign = IIf(Left$(tail, 1) = "-", -1, 1)
        tail = Replace(Mid$(tail, 2), ":", "")
        offsetMin = sign * (CLng(Left$(tail, 2)) * 60 + CLng(Mid$(tail, 3, 2)))
    End If
    ParseJiraDate = d - offsetMin / 1440#
End Function

Public Function FieldToCell(ByVal value As Variant) As Variant
    ' Turn any Jira field value into something a cell can hold.
    If IsObject(value) Then
        If TypeName(value) = "Collection" Then
            Cov "TXT11": FieldToCell = JoinCollection(value)
        Else
            Cov "TXT12": FieldToCell = ObjectToText(value)
        End If
    ElseIf IsNull(value) Or IsEmpty(value) Then
        Cov "TXT13": FieldToCell = Empty
    ElseIf VarType(value) = vbString Then
        FieldToCell = Left$(value, 32767)
    Else
        FieldToCell = value
    End If
End Function

Private Function JoinCollection(ByVal c As Collection) As String
    Dim parts() As String, i As Long
    If c.Count = 0 Then Exit Function
    ReDim parts(1 To c.Count)
    For i = 1 To c.Count
        parts(i) = CStr(FieldToCell(c(i)))
    Next i
    JoinCollection = Left$(Join(parts, "; "), 32767)
End Function

Private Function ObjectToText(ByVal d As Object) As String
    Dim k As Variant
    If d.Exists("type") Then
        If Not IsObject(d("type")) Then
            If d("type") = "doc" Then Cov "TXT14": ObjectToText = Left$(AdfToText(d), 32767): Exit Function
        End If
    End If
    For Each k In Array("displayName", "value", "name", "key", "id")
        If d.Exists(k) Then
            If Not IsObject(d(k)) And Not IsNull(d(k)) Then ObjectToText = CStr(d(k)): Exit Function
        End If
    Next k
    Cov "TXT15"
    ObjectToText = Left$(JsonStringify(d), 32767)
End Function

Public Function AdfToText(ByVal node As Variant) As String
    ' Atlassian Document Format (rich text, e.g. description) -> plain text.
    Dim out As String, child As Variant, nodeType As String
    If Not IsObject(node) Then Exit Function
    If TypeName(node) <> "Dictionary" Then Exit Function
    nodeType = CStr(JGet(node, "type") & "")
    If node.Exists("text") Then out = CStr(node("text") & "")
    If node.Exists("content") Then
        For Each child In node("content")
            out = out & AdfToText(child)
        Next child
        If nodeType = "paragraph" Or nodeType = "heading" Or nodeType = "listItem" Then out = out & vbLf
    End If
    AdfToText = out
End Function
```

**modHttp**: 154 lines, 7 coverage markers.

### `vba/modHttp.bas`  (MD5 `74263efe377c649e5c4c27b2ab9bdf19`)

```vba
Option Explicit
' HTTP with WinHttp: named proxy, Windows (integrated) auth to the proxy, asynchronous wait that
' keeps Excel responsive (DoEvents), retry on 429 / 503 with Retry-After and exponential backoff.
' Test hook: when gHttpFake is set, no network call is made and queued fake responses are returned.

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Public gHttpFake As Object      ' Dictionary: "METHOD url" -> Collection of Array(status, body, retryAfter)
Public gHttpLog As Collection   ' one Array(method, url, body) per attempt

Private Const PROXY_SETTING_PROXY As Long = 2      ' HTTPREQUEST_PROXYSETTING_PROXY
Private Const AUTOLOGON_ALWAYS As Long = 0         ' AutoLogonPolicy_Always: send Windows credentials

Public Function NewHttpConfig() As Object
    Cov "HTP01"
    Dim c As Object
    Set c = CreateObject("Scripting.Dictionary")
    c("ProxyServer") = ""            ' "proxy.client.local:8080"; empty = WinHTTP default (no proxy)
    c("ProxyBypass") = ""            ' "*.client.local;<local>"
    c("ProxyWindowsAuth") = True     ' send the Windows session credentials to the proxy
    c("ResolveMs") = 10000: c("ConnectMs") = 15000: c("SendMs") = 30000: c("ReceiveMs") = 120000
    c("MaxRetries") = 4: c("RetryBaseMs") = 2000: c("RetryMaxMs") = 30000
    c("Authorization") = ""
    Set NewHttpConfig = c
End Function

Public Function HttpRequest(ByVal cfg As Object, ByVal method As String, ByVal url As String, _
                            Optional ByVal body As String = "") As Object
    ' Returns a Dictionary: Status (0 = network error), Text, RetryAfter, Error, Attempts.
    Cov "HTP02"
    Dim attempt As Long, resp As Object
    For attempt = 0 To cfg("MaxRetries")
        Set resp = HttpSendOnce(cfg, method, url, body)
        resp("Attempts") = attempt + 1
        If Not IsRetryable(resp) Then Exit For
        If attempt = cfg("MaxRetries") Then Cov "HTP03": Exit For
        Cov "HTP04"
        WaitMs RetryDelayMs(cfg, attempt, resp("RetryAfter"))
    Next attempt
    Set HttpRequest = resp
End Function

Private Function IsRetryable(ByVal resp As Object) As Boolean
    ' Only call HttpRequest with POST for read-only endpoints (search, bulkfetch).
    If resp("Status") = 429 Then IsRetryable = True: Exit Function
    If resp("Status") = 503 And Len(resp("RetryAfter")) > 0 Then IsRetryable = True
End Function

Public Function RetryDelayMs(ByVal cfg As Object, ByVal attempt As Long, ByVal retryAfter As String) As Double
    Cov "HTP05"
    Dim backoff As Double, fromHeader As Double
    backoff = cfg("RetryBaseMs") * 2 ^ attempt
    If backoff > cfg("RetryMaxMs") Then backoff = cfg("RetryMaxMs")
    backoff = backoff * Rnd()                                   ' full jitter
    If Len(retryAfter) > 0 Then fromHeader = Val(retryAfter) * 1000# + Rnd() * 500#
    RetryDelayMs = IIf(fromHeader > backoff, fromHeader, backoff)
End Function

Private Function HttpSendOnce(ByVal cfg As Object, ByVal method As String, ByVal url As String, _
                              ByVal body As String) As Object
    Dim r As Object, req As Object, raw As Variant
    Set r = CreateObject("Scripting.Dictionary")
    r("Status") = 0: r("Text") = "": r("RetryAfter") = "": r("Error") = ""
    If gHttpLog Is Nothing Then Set gHttpLog = New Collection
    gHttpLog.Add Array(method, url, body)
    If Not gHttpFake Is Nothing Then Cov "HTP06": Set HttpSendOnce = FakeResponse(r, method, url): Exit Function

    On Error GoTo NetError
    Set req = CreateObject("WinHttp.WinHttpRequest.5.1")
    req.Open method, url, True                                  ' True = asynchronous
    If Len(cfg("ProxyServer")) > 0 Then req.SetProxy PROXY_SETTING_PROXY, cfg("ProxyServer"), cfg("ProxyBypass")
    If cfg("ProxyWindowsAuth") Then req.SetAutoLogonPolicy AUTOLOGON_ALWAYS
    req.SetTimeouts cfg("ResolveMs"), cfg("ConnectMs"), cfg("SendMs"), cfg("ReceiveMs")
    req.SetRequestHeader "Accept", "application/json"
    If Len(cfg("Authorization")) > 0 Then req.SetRequestHeader "Authorization", cfg("Authorization")
    If Len(body) > 0 Then
        req.SetRequestHeader "Content-Type", "application/json; charset=utf-8"
        req.Send Utf8Encode(body)
    Else
        req.Send
    End If
    Do While Not req.WaitForResponse(0)                         ' returns at once; keeps Excel alive
        DoEvents
        Sleep 20
    Loop
    r("Status") = req.Status
    raw = req.ResponseBody
    If IsArray(raw) Then r("Text") = Utf8Decode(raw)   ' always decode as UTF-8, whatever the header says
    r("RetryAfter") = SafeHeader(req, "Retry-After")
    Set HttpSendOnce = r
    Exit Function
NetError:
    Cov "HTP07"
    r("Status") = 0
    r("Error") = Err.Number & " " & Err.Description
    Set HttpSendOnce = r
End Function

Private Function SafeHeader(ByVal req As Object, ByVal name As String) As String
    On Error Resume Next
    SafeHeader = req.GetResponseHeader(name)
    Err.Clear                                   ' a missing header is normal: leave no stale error
End Function

Public Sub WaitMs(ByVal ms As Double)
    ' Waits without freezing Excel. Timer wraps at midnight: handled.
    Dim t0 As Double, elapsed As Double
    t0 = Timer
    Do
        DoEvents
        Sleep 20
        elapsed = (Timer - t0) * 1000#
        If elapsed < 0 Then elapsed = elapsed + 86400000#
    Loop While elapsed < ms
End Sub

' ---------- Test hook ----------

Public Sub FakeReset()
    Set gHttpFake = CreateObject("Scripting.Dictionary")
    Set gHttpLog = New Collection
End Sub

Public Sub FakeAdd(ByVal method As String, ByVal url As String, ByVal status As Long, _
                   ByVal body As String, Optional ByVal retryAfter As String = "")
    Dim key As String
    key = method & " " & url
    If gHttpFake Is Nothing Then FakeReset
    If Not gHttpFake.Exists(key) Then gHttpFake.Add key, New Collection
    gHttpFake(key).Add Array(status, body, retryAfter)
End Sub

Public Sub FakeOff()
    Set gHttpFake = Nothing
End Sub

Private Function FakeResponse(ByVal r As Object, ByVal method As String, ByVal url As String) As Object
    Dim key As String, q As Collection, item As Variant
    key = method & " " & url
    If gHttpFake.Exists(key) Then Set q = gHttpFake(key)
    If q Is Nothing Then
        r("Status") = 404: r("Text") = "{""fake"":""no response queued""}"
    ElseIf q.Count = 0 Then
        r("Status") = 404: r("Text") = "{""fake"":""queue empty""}"
    Else
        item = q(1): q.Remove 1
        r("Status") = item(0): r("Text") = item(1): r("RetryAfter") = item(2)
    End If
    Set FakeResponse = r
End Function
```

**modJira**: 198 lines, 13 coverage markers.

### `vba/modJira.bas`  (MD5 `cb0c3fd8aac83eb525628adb55aa5302`)

```vba
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
```

**modFlow**: 81 lines, 7 coverage markers.

### `vba/modFlow.bas`  (MD5 `d709521a710f21e3319df21038c8e463`)

```vba
Option Explicit
' Status intervals: the traceable base for Flow Time, Flow Load and Flow Efficiency.
' One row per (issue, status visit): when it entered, when it left, how long, which change caused it.

Public Function StatusChangesFromChangelog(ByVal changelogRows As Collection, ByVal idToKey As Object) As Collection
    ' Keeps fieldId = "status". Output rows: Array(issueKey, changeId, createdUtc, fromString, toString)
    Cov "FLW01"
    Dim out As New Collection, r As Variant, key As String
    For Each r In changelogRows
        If CStr(r(5) & "") = "status" Then
            key = CStr(r(0) & "")
            If idToKey.Exists(key) Then key = idToKey(key)
            out.Add Array(key, r(1), ParseJiraDate(r(2)), r(9), r(11))
        End If
    Next r
    Set StatusChangesFromChangelog = out
End Function

Public Function BuildStatusIntervals(ByVal issueInfo As Object, ByVal statusChanges As Collection, _
                                     ByVal asOfUtc As Date, ByVal categories As Object) As Collection
    ' issueInfo: key -> Array(createdUtc, currentStatus).  categories: status -> "Active" / "Wait" / "Done" / "Removed" / "Backlog" (unknown = "Unclassified").
    ' Output rows: Array(key, seq, status, category, enteredUtc, leftUtc, durationDays, isOpen, changeId, check)
    Cov "FLW02"
    Dim byIssue As Object, ch As Variant, key As Variant, out As New Collection
    Set byIssue = CreateObject("Scripting.Dictionary")
    For Each ch In statusChanges
        If Not byIssue.Exists(ch(0)) Then byIssue.Add ch(0), New Collection
        byIssue(ch(0)).Add ch
    Next ch
    For Each key In issueInfo.Keys
        If byIssue.Exists(key) Then
            AddIssueIntervals out, CStr(key), issueInfo(key), SortByDate(byIssue(key)), asOfUtc, categories
        Else
            Cov "FLW03"
            AddIssueIntervals out, CStr(key), issueInfo(key), New Collection, asOfUtc, categories
        End If
    Next key
    Set BuildStatusIntervals = out
End Function

Private Sub AddIssueIntervals(ByVal out As Collection, ByVal key As String, ByVal info As Variant, _
                              ByVal changes As Collection, ByVal asOfUtc As Date, ByVal categories As Object)
    Dim status As String, entered As Variant, seq As Long, ch As Variant, check As String
    entered = info(0)
    If changes.Count = 0 Then
        status = CStr(info(1) & "")
    Else
        status = CStr(changes(1)(3) & "")                     ' fromString of the first change
    End If
    For Each ch In changes
        seq = seq + 1
        check = ""
        If CStr(ch(3) & "") <> status Then Cov "FLW04": check = "FROM_MISMATCH(prev=" & status & ")"
        If ch(2) < entered Then Cov "FLW05": check = check & IIf(Len(check) > 0, ";", "") & "BEFORE_CREATED"
        out.Add Array(key, seq, status, Category(categories, status), entered, ch(2), _
                      ch(2) - entered, False, ch(1), check)
        status = CStr(ch(4) & "")
        entered = ch(2)
    Next ch
    seq = seq + 1
    check = ""
    If status <> CStr(info(1) & "") Then Cov "FLW06": check = "CURRENT_MISMATCH(now=" & info(1) & ")"
    out.Add Array(key, seq, status, Category(categories, status), entered, Empty, asOfUtc - entered, True, Empty, check)
End Sub

Private Function Category(ByVal categories As Object, ByVal status As String) As String
    If categories.Exists(status) Then Category = categories(status) Else Cov "FLW07": Category = "Unclassified"
End Function

Private Function SortByDate(ByVal changes As Collection) As Collection
    ' Insertion sort on createdUtc (index 2). Changelog pages are date-sorted, but do not rely on it.
    Dim sorted As New Collection, ch As Variant, i As Long, placed As Boolean
    For Each ch In changes
        placed = False
        For i = 1 To sorted.Count
            If ch(2) < sorted(i)(2) Then sorted.Add ch, Before:=i: placed = True: Exit For
        Next i
        If Not placed Then sorted.Add ch
    Next ch
    Set SortByDate = sorted
End Function
```

**modBackground**: 98 lines, 5 coverage markers.

### `vba/modBackground.bas`  (MD5 `1d098ce38925b8dc55bcf54404cf314f`)

```vba
Option Explicit
' Option B: the PowerShell extraction runs OUTSIDE Excel (started by the user, as today).
' Excel never starts a process: antivirus / ASR rules block Office from launching PowerShell
' (tested: a hidden powershell.exe started from VBA was blocked).
' The script writes page-*.json files and finally "done.flag" in an export folder.
' Excel only watches that folder with Application.OnTime, then imports the files. Excel stays usable.

Public gExportOutDir As String
Public gExportNextCheck As Date
Public gExportStatus As String       ' "", "WATCHING", "DONE", "TIMEOUT"
Private mExportDeadline As Date

Public Sub WatchExportFolder(ByVal outDir As String, Optional ByVal timeoutMinutes As Long = 60)
    ' Call this from a button after starting the PowerShell script yourself.
    Cov "BKG01"
    gExportOutDir = outDir
    gExportStatus = "WATCHING"
    mExportDeadline = Now + timeoutMinutes / 1440#
    ScheduleCheck
End Sub

Public Sub ScheduleCheck()
    gExportNextCheck = Now + TimeSerial(0, 0, 5)
    Application.OnTime gExportNextCheck, "CheckBackgroundExport"
End Sub

Public Sub CheckBackgroundExport()
    ' Called by Application.OnTime. Keep it short: it runs on Excel's UI thread.
    If gExportStatus <> "WATCHING" Then Exit Sub
    If IsExportDone(gExportOutDir) Then
        Cov "BKG02"
        gExportStatus = "DONE"
        Application.StatusBar = "Jira export finished: click Import."
    ElseIf Now > mExportDeadline Then
        Cov "BKG03"
        gExportStatus = "TIMEOUT"
        Application.StatusBar = "Jira export not finished in time. Check the PowerShell window."
    Else
        Application.StatusBar = "Waiting for the Jira export files..."
        ScheduleCheck
    End If
End Sub

Public Sub CancelBackgroundCheck()
    On Error Resume Next
    Application.OnTime gExportNextCheck, "CheckBackgroundExport", , False
    Err.Clear
    gExportStatus = ""
    Application.StatusBar = False
End Sub

Public Function IsExportDone(ByVal outDir As String) As Boolean
    IsExportDone = Len(Dir$(outDir & "\done.flag")) > 0
End Function

Public Function ReadUtf8File(ByVal path As String) As String
    Cov "BKG04"
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8": st.Open
    st.LoadFromFile path
    ReadUtf8File = st.ReadText
    st.Close
End Function

Public Function ListFiles(ByVal outDir As String, ByVal pattern As String) As Collection
    ' Files sorted by name (page-0001.json, page-0002.json ...).
    Dim c As New Collection, f As String, i As Long, placed As Boolean
    f = Dir$(outDir & "\" & pattern)
    Do While Len(f) > 0
        placed = False
        For i = 1 To c.Count
            If f < c(i) Then c.Add f, Before:=i: placed = True: Exit For
        Next i
        If Not placed Then c.Add f
        f = Dir$()
    Loop
    Set ListFiles = c
End Function

Public Function ImportSearchPages(ByVal outDir As String, ByRef names As Object) As Collection
    ' Reads page-*.json files written by the script (each = one /search/jql response) into issues.
    Cov "BKG05"
    Dim issues As New Collection, f As Variant, data As Variant, issue As Variant, k As Variant
    If names Is Nothing Then Set names = CreateObject("Scripting.Dictionary")
    For Each f In ListFiles(outDir, "page-*.json")
        Set data = JsonParse(ReadUtf8File(outDir & "\" & f))
        For Each issue In AsList(JGet(data, "issues"))
            issues.Add issue
        Next issue
        If IsObject(JGet(data, "names")) Then
            For Each k In data("names").Keys
                If Not names.Exists(k) Then names(k) = data("names")(k)
            Next k
        End If
    Next f
    Set ImportSearchPages = issues
End Function
```

**modTests**: 609 lines, 0 coverage markers.

### `vba/modTests.bas`  (MD5 `39b9f1f5a44449b645af87721cce5b17`)

```vba
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
```

**modButtons**: 51 lines, UI entry points, never run by the tests.

### `vba/modButtons.bas`  (MD5 `3bbe95b810cc0722eabed5c2bdbabbb2`)

```vba
Option Explicit
' Entry points for the buttons. A macro started by a button or by Alt+F8 must be a Public Sub
' without arguments: Functions and Subs with parameters do not appear in the Alt+F8 list.
' Keep these Subs trivial (one call, then a message). The automated tests never call them:
' a MsgBox in a hidden Excel would block the run.
' The runs use DoEvents, so Excel stays usable and buttons stay clickable: a busy flag blocks
' a second click (it would clear the results in the middle of a run).

Private mBusy As Boolean

Public Sub BtnRunAllTests()
    RunFromButton False
End Sub

Public Sub BtnRunNetworkTests()
    RunFromButton True
End Sub

Public Sub BtnClearResults()
    If mBusy Then Beep: Exit Sub
    ClearResults
End Sub

Private Sub ClearResults()
    ThisWorkbook.Worksheets("TestResults").Range("A:H").ClearContents
    ThisWorkbook.Worksheets("CoverageHits").Cells.ClearContents
    ProgressReset
End Sub

Private Sub RunFromButton(ByVal includeNetwork As Boolean)
    Dim summary As String
    If mBusy Then Beep: Exit Sub
    mBusy = True
    On Error GoTo Failed
    ClearResults
    Application.Cursor = xlWait
    summary = RunWithCoverage(includeNetwork)
    Application.Cursor = xlDefault
    Application.StatusBar = False
    Application.Calculate
    mBusy = False
    MsgBox "Tests: " & summary & vbCrLf & "Coverage: " & _
           Format$(ThisWorkbook.Worksheets("CoverageMap").Range("E2").Value2, "0.0%"), _
           vbInformation, "Test results"
    Exit Sub
Failed:
    Application.Cursor = xlDefault
    Application.StatusBar = False
    mBusy = False
    MsgBox "Tests stopped by an error: " & Err.Number & " - " & Err.Description, vbExclamation, "Test results"
End Sub
```
