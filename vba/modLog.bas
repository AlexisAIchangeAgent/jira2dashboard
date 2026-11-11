Option Explicit
' Local log of the runs, to debug them later: one line per run, per step and per error, in the folder logs next to the
' workbook (<folder of the .xlsm>\logs\jira2dashboard.log: the workbook runs from a folder not known in advance). When the
' workbook has no folder (never saved, or opened from a web address) or the folder cannot be written, the log goes to
' %LOCALAPPDATA%\jira2dashboard\logs (%TEMP% when LOCALAPPDATA is not set). Nothing leaves the computer.
' Each line is opened, written and closed at once: a closed line is in the hands of Windows and survives a crash of
' Excel. After a crash (no VBA error is raised then), the last BEGIN without its END names the step that was running.
' The log never holds the API token, an e-mail address, a URL or an answer of Jira: callers pass step names and numbers,
' and the descriptions of the errors that quote the answer of Jira (602, 607) are not written.
' The logger never raises: a line that cannot be written is dropped.
' Every error line starts with a support code, J2D-<AREA>-<NNN> (ErrorCode below; the table with the cause and the check of
' each code is src/jira2dashboard/error_codes.py, shown in the sheet ErrorCodes): the person reads the code to the support.
' Level (Config LogLevel, read by the caller): Steps = everything; Errors = no BEGIN/END lines.
' A run that does not end (Excel closed or crashed) leaves a flag in the registry (HKCU, VB and VBA Program Settings):
' the next run says so in the log.

Public gLogPath As String                          ' empty = the default file; the tests point it to a file of their own
Private Const LOG_MAX_BYTES As Long = 1000000      ' above this size a run starts a new file (the old one becomes .1.log)
Private Const LOG_MAX_TEXT As Long = 300           ' longer texts are cut
Public gLogOpenFake As Boolean                      ' tests: LogOpen records its command in gLogOpenCommand and starts nothing
Public gLogOpenCommand As String
Public gLogFallbackFile As String                 ' tests: the file used when gLogPath cannot be written
Private mFallback As Boolean                       ' True once the folder next to the workbook could not be written (until the next run)
Private mErrorsOnly As Boolean                     ' True while a run asked the level Errors
Private Const STEPS_SHEET As String = "Steps"
Private mStepSheet As Object
Private mStepRows As Collection                     ' rows of the steps begun and not ended (a stack: steps can nest)
Private mStepStarts As Collection                   ' their start times (Timer)
Private mStepNext As Long                           ' the next free row
Private mLastCode As String                         ' the support code of the last error of the run (shown in the sheet Steps)

Public Function LogFolderFrom(ByVal workbookPath As String, ByVal localAppData As String, ByVal temp As String) As String
    ' logs next to the workbook; the folder of the user when the workbook has no folder of its own.
    Cov "LOG01"
    If Len(workbookPath) > 0 And LCase$(Left$(workbookPath, 4)) <> "http" Then Cov "LOG19": LogFolderFrom = workbookPath & "\logs": Exit Function
    If Len(localAppData) = 0 Then Cov "LOG02": localAppData = temp
    LogFolderFrom = localAppData & "\jira2dashboard\logs"
End Function

Public Function LogFile() As String
    If mFallback Then LogFile = FallbackFile(): Exit Function
    If Len(gLogPath) > 0 Then LogFile = gLogPath: Exit Function
    LogFile = LogFolderFrom(ThisWorkbook.Path, Environ$("LOCALAPPDATA"), Environ$("TEMP")) & "\jira2dashboard.log"
End Function

Private Function FallbackFile() As String
    If Len(gLogFallbackFile) > 0 Then FallbackFile = gLogFallbackFile: Exit Function
    FallbackFile = LogFolderFrom("", Environ$("LOCALAPPDATA"), Environ$("TEMP")) & "\jira2dashboard.log"
End Function

Public Sub LogLine(ByVal kind As String, ByVal text As String)
    ' "2026-11-04 21:05:33.123 KIND text", plain ASCII on one line. Three tries (another Excel may be writing), then dropped.
    Dim f As Integer, file As String, attempt As Long, line As String
    On Error Resume Next
    file = LogFile()
    line = Stamp() & " " & kind & " " & OneLine(text)
    For attempt = 1 To 3
        Err.Clear
        f = FreeFile
        Open file For Append Access Write Lock Write As #f
        If Err.Number = 0 Then
            Cov "LOG03"
            Print #f, line
            Close #f
            Err.Clear
            Exit Sub
        End If
        If Err.Number = 76 Then Cov "LOG04": MakeFolders file     ' path not found: the first line on this computer
    Next attempt
    Cov "LOG05"
    Err.Clear
    If Not mFallback And (Len(gLogPath) = 0 Or Len(gLogFallbackFile) > 0) Then
        Cov "LOG20"                                          ' the folder next to the workbook is read-only: the folder of the user
        mFallback = True
        LogLine kind, text
    End If
End Sub

Public Sub LogBegin(ByVal stepName As String)
    StepsBegin stepName                                ' the sheet Steps shows every step, whatever the level
    If mErrorsOnly Then Exit Sub
    LogLine "BEGIN", stepName
End Sub

Public Sub LogEnd(ByVal stepName As String)
    StepsEnd stepName
    If mErrorsOnly Then Exit Sub
    LogLine "END", stepName
End Sub

Public Sub LogRunStart(ByVal how As String, Optional ByVal level As String = "Steps")
    ' The first line of a run: a new file when the log is too big, then what is needed to read the lines that follow,
    ' and a warning when the previous run never ended.
    Cov "LOG06"
    Dim bits As String, previous As String
    On Error Resume Next
    #If Win64 Then
        bits = "64-bit"
    #Else
        bits = "32-bit"
    #End If
    mFallback = False                                  ' try the folder next to the workbook again
    mLastCode = ""
    StepsReset
    RotateLog
    mErrorsOnly = (StrComp(level, "Errors", vbTextCompare) = 0)
    LogLine "RUN", how & " | Excel " & Application.Version & " build " & Application.Build & " " & bits & " | " & ThisWorkbook.Name
    previous = GetSetting("jira2dashboard", FlagSection(), "RunInProgress", "")
    If Len(previous) > 0 Then
        Cov "LOG12"
        LogLine "RUN", "previous run did not end (started " & previous & "): Excel was closed or stopped during it"
    End If
    SaveSetting "jira2dashboard", FlagSection(), "RunInProgress", Stamp()
    Err.Clear
End Sub

Public Sub LogRunEnd(ByVal result As String)
    ' The last line of a run: "end ok", "end cancelled", "end stopped". Clears the flag of the run in progress.
    Cov "LOG13"
    On Error Resume Next
    LogLine "RUN", "end " & result
    StepsClose result
    mErrorsOnly = False
    DeleteSetting "jira2dashboard", FlagSection(), "RunInProgress"
    Err.Clear
End Sub

' ---- The sheet Steps: the steps of the run in progress, with their status and time ----------------------------------
' One row per step (LogBegin adds it as "running", LogEnd closes it as "done" with its time, in seconds); a run that stops
' leaves its running rows as "error" or "cancelled". The sheet is created at the first run and emptied at each run start.
' Whatever happens here is ignored: the sheet is a view of the run, never a reason to stop it.


Private Sub StepsReset()
    Cov "LOG21"
    Dim shown As Object
    On Error Resume Next
    Set mStepRows = New Collection: Set mStepStarts = New Collection
    Set mStepSheet = Nothing
    Set shown = Application.ActiveSheet
    If SheetExists(ThisWorkbook, STEPS_SHEET) Then
        Set mStepSheet = ThisWorkbook.Worksheets(STEPS_SHEET)
    Else
        Cov "LOG22"
        Set mStepSheet = ThisWorkbook.Worksheets.Add(, ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        mStepSheet.Name = STEPS_SHEET
        mStepSheet.Columns(1).ColumnWidth = 70
        mStepSheet.Columns(2).ColumnWidth = 12
        mStepSheet.Columns(3).ColumnWidth = 18
        mStepSheet.Columns(3).NumberFormat = "0.0"
        mStepSheet.Columns(4).ColumnWidth = 20
        shown.Activate                              ' the sheet shown before stays shown
    End If
    If mStepSheet Is Nothing Then Err.Clear: Exit Sub
    mStepSheet.Cells.ClearContents
    mStepSheet.Range("A1:D1").Value = Array("Step", "Status", "Time elapsed (s)", "Error code")
    mStepSheet.Range("A1:D1").Font.Bold = True
    mStepNext = 2
    Err.Clear
End Sub

Private Sub StepsBegin(ByVal stepName As String)
    On Error Resume Next
    If mStepSheet Is Nothing Then Exit Sub
    mStepSheet.Cells(mStepNext, 1).Value = SafeCell(stepName)
    mStepSheet.Cells(mStepNext, 2).Value = "running"
    mStepRows.Add mStepNext
    mStepStarts.Add Timer
    mStepNext = mStepNext + 1
    Err.Clear
End Sub

Private Sub StepsEnd(ByVal text As String)
    ' Closes the last step begun. The text of the END replaces the name when it says more ("search page 3: 100 issues").
    Dim r As Long, started As Double, current As String
    On Error Resume Next
    If mStepSheet Is Nothing Then Exit Sub
    If mStepRows.Count = 0 Then Cov "LOG23": Exit Sub          ' an END whose BEGIN was not seen (another run, or a new sheet)
    r = mStepRows(mStepRows.Count): started = mStepStarts(mStepStarts.Count)
    mStepRows.Remove mStepRows.Count: mStepStarts.Remove mStepStarts.Count
    current = CStr(mStepSheet.Cells(r, 1).Value)
    If Len(text) > Len(current) And Left$(text, Len(current)) = current Then
        Cov "LOG25"
        mStepSheet.Cells(r, 1).Value = SafeCell(text)
    End If
    StepsFinish r, started, "done"
    Err.Clear
End Sub

Private Sub StepsClose(ByVal result As String)
    ' The run ended: a step still running stopped with it.
    Dim i As Long, status As String
    On Error Resume Next
    If mStepSheet Is Nothing Then Exit Sub
    If mStepRows.Count = 0 Then Exit Sub
    Cov "LOG24"
    If LCase$(Left$(result, 2)) = "ok" Then
        status = "done"
    ElseIf LCase$(Left$(result, 9)) = "cancelled" Then
        status = "cancelled"
    Else
        status = "error"
    End If
    For i = mStepRows.Count To 1 Step -1
        StepsFinish CLng(mStepRows(i)), CDbl(mStepStarts(i)), status
    Next i
    Set mStepRows = New Collection: Set mStepStarts = New Collection
    Err.Clear
End Sub

Private Sub StepsFinish(ByVal r As Long, ByVal started As Double, ByVal status As String)
    Dim seconds As Double
    seconds = Timer - started
    If seconds < 0 Then seconds = seconds + 86400            ' past midnight
    mStepSheet.Cells(r, 2).Value = status
    mStepSheet.Cells(r, 3).Value = Round(seconds, 1)
    If status = "error" And Len(mLastCode) > 0 Then Cov "LOG31": mStepSheet.Cells(r, 4).Value = mLastCode
End Sub

Private Function FlagSection() As String
    ' The tests write to a log file of their own: their flag is apart too.
    If Len(gLogPath) = 0 Then FlagSection = "Log" Else FlagSection = "LogTests"
End Function

Public Sub LogError(ByVal where As String, ByVal number As Long, ByVal description As String)
    ' The caller passes the values it saved from Err: any On Error statement clears Err.
    Cov "LOG07"
    Dim text As String, code As String
    code = ErrorCode(number, description)
    mLastCode = code
    text = code & " | " & where & " | " & ErrorNumber(number) & " | "
    If number = vbObjectError + 602 Or number = vbObjectError + 607 Or number = vbObjectError + 513 Then
        Cov "LOG08"
        text = text & "(description not logged: it quotes the answer of Jira)"
    Else
        text = text & description
    End If
    LogLine "ERROR", text
End Sub

Public Function LogHint() As String
    ' The sentence added to the status of a stopped run: where the log is.
    Cov "LOG14"
    LogHint = " See the log: " & LogFile()
End Function

Public Sub LogOpen(Optional ByVal enabled As Boolean = True)
    ' After an error, shows the log in Notepad when a person is at the keyboard (Excel started by the user, not by a test
    ' or a script). Best effort: never raises. gLogOpenFake lets the tests read the command instead of starting Notepad.
    Dim cmd As String
    If Not enabled Then Cov "LOG18": Exit Sub                  ' Config OpenLogOnError = No
    cmd = NotepadCommand(LogFile())
    If gLogOpenFake Then Cov "LOG15": gLogOpenCommand = cmd: Exit Sub
    On Error Resume Next
    If Not Application.UserControl Then Cov "LOG16": Exit Sub
    Shell cmd, vbNormalFocus
    Err.Clear
End Sub

Public Function NotepadCommand(ByVal file As String) As String
    Cov "LOG17"
    NotepadCommand = "notepad.exe """ & file & """"
End Function

Public Function ErrorNumber(ByVal number As Long) As String
    ' The errors raised by this project read vbObjectError+n; the others keep their number.
    If number >= vbObjectError And number < vbObjectError + 65536 Then
        Cov "LOG09"
        ErrorNumber = "vbObjectError+" & CStr(number - vbObjectError)
    Else
        ErrorNumber = CStr(number)
    End If
End Function

Public Function ErrorCode(ByVal number As Long, Optional ByVal description As String = "") As String
    ' The support code of an error: J2D-<AREA>-<NNN>, never reused, never renumbered. AREA: CFG (parameters), JIRA (what Jira
    ' answered: the number is the HTTP status for 4xx/5xx), NET (no answer: connection, proxy, certificate, time out; the number
    ' is the WinHTTP error), DATA (JSON), FLOW (tables and charts), ROVO, SYS (Excel, the run itself), VBA (an error of VBA
    ' or Excel: its number). The table with the cause and the check of each code: src/jira2dashboard/error_codes.py.
    Dim n As Long, status As Long, net As Long
    If number >= vbObjectError And number < vbObjectError + 65536 Then
        n = number - vbObjectError
        Select Case n
            Case 600: ErrorCode = "J2D-CFG-001"
            Case 608: ErrorCode = "J2D-CFG-002"
            Case 603: ErrorCode = "J2D-CFG-003"
            Case 604: ErrorCode = "J2D-CFG-004"
            Case 620: ErrorCode = "J2D-CFG-006"
            Case 701: ErrorCode = "J2D-CFG-010"
            Case 704: ErrorCode = "J2D-CFG-011"
            Case 703: ErrorCode = "J2D-CFG-012"
            Case 601: ErrorCode = "J2D-JIRA-010"
            Case 606: ErrorCode = "J2D-JIRA-011"
            Case 505: ErrorCode = "J2D-JIRA-012"
            Case 621: ErrorCode = "J2D-JIRA-013"
            Case 607: ErrorCode = "J2D-JIRA-020"
            Case 513: ErrorCode = "J2D-DATA-001"
            Case 514: ErrorCode = "J2D-DATA-002"
            Case 702: ErrorCode = "J2D-FLOW-001"
            Case 705: ErrorCode = "J2D-FLOW-002"
            Case 706: ErrorCode = "J2D-FLOW-003"
            Case 707: ErrorCode = "J2D-FLOW-004"
            Case 708: ErrorCode = "J2D-FLOW-005"
            Case 709: ErrorCode = "J2D-FLOW-006"
            Case 630: ErrorCode = "J2D-ROVO-001"
            Case 631: ErrorCode = "J2D-ROVO-002"
            Case 632: ErrorCode = "J2D-ROVO-003"
            Case 610: ErrorCode = "J2D-SYS-001"
            Case 611: ErrorCode = "J2D-SYS-002"
            Case 602                                           ' two causes: the BaseUrl check, and an answer of Jira that is not 200
                If Left$(description, 5) <> "HTTP " Then
                    ErrorCode = "J2D-CFG-005"
                Else
                    status = CLng(Val(Mid$(description, 6)))
                    If status > 0 Then
                        Cov "LOG26"
                        ErrorCode = "J2D-JIRA-" & Format$(status, "000")
                    Else
                        Cov "LOG27"
                        net = WinHttpNumber(description)
                        If net > 0 Then ErrorCode = "J2D-NET-" & CStr(net) Else ErrorCode = "J2D-NET-001"
                    End If
                End If
            Case Else
                Cov "LOG30"
                ErrorCode = "J2D-GEN-" & Format$(n, "000")
        End Select
    Else
        net = WinHttpOf(number)
        If net > 0 Then
            Cov "LOG28"
            ErrorCode = "J2D-NET-" & CStr(net)
        Else
            Cov "LOG29"
            If number < 0 Then ErrorCode = "J2D-VBA-" & Hex$(number) Else ErrorCode = "J2D-VBA-" & CStr(number)
        End If
    End If
End Function

Private Function WinHttpOf(ByVal number As Long) As Long
    ' A WinHTTP error comes as 0x80072Exx (HRESULT) or as the plain 12xxx: its number, 0 for any other error.
    Dim low As Long
    low = number And &HFFFF&
    If number = low Or (number And &HFFFF0000) = &H80070000 Then
        If low >= 12000 And low <= 12200 Then WinHttpOf = low
    End If
End Function

Private Function WinHttpNumber(ByVal description As String) As Long
    ' The text of an HTTP 0 reads "... attempt(s). <error number> <description>": the number is the one of WinHTTP.
    Dim p As Long, v As Double
    p = InStr(description, "attempt(s). ")
    If p = 0 Then Exit Function
    v = Val(Mid$(description, p + 12))
    If v < -2147483648# Or v > 2147483647# Then Exit Function
    WinHttpNumber = WinHttpOf(CLng(v))
End Function

Private Function Stamp() As String
    Dim t As Single
    t = Timer
    Stamp = Format$(Now, "yyyy-mm-dd hh:nn:ss") & "." & Format$(Int((t - Int(t)) * 1000), "000")
End Function

Private Function OneLine(ByVal text As String) As String
    ' Line breaks and tabs become spaces, any other character outside printable ASCII becomes ?.
    Dim i As Long, c As Long
    If Len(text) > LOG_MAX_TEXT Then Cov "LOG10": text = Left$(text, LOG_MAX_TEXT) & "..."
    For i = 1 To Len(text)
        c = AscW(Mid$(text, i, 1))
        If c = 9 Or c = 10 Or c = 13 Then
            Mid$(text, i, 1) = " "
        ElseIf c < 32 Or c > 126 Then
            Mid$(text, i, 1) = "?"
        End If
    Next i
    OneLine = text
End Function

Private Sub MakeFolders(ByVal file As String)
    ' Creates each missing folder of the path (MkDir fails harmlessly on a folder that exists). No Dir$: a caller may be
    ' in the middle of a Dir$ loop.
    Dim parts() As String, i As Long, path As String
    On Error Resume Next
    parts = Split(file, "\")
    path = parts(0)
    For i = 1 To UBound(parts) - 1
        path = path & "\" & parts(i)
        MkDir path
    Next i
    Err.Clear
End Sub

Private Sub RotateLog()
    Dim file As String, size As Long
    On Error Resume Next
    file = LogFile()
    size = -1
    size = FileLen(file)                             ' stays -1 when the file does not exist yet
    If size <= LOG_MAX_BYTES Then Err.Clear: Exit Sub
    Cov "LOG11"
    Kill Left$(file, Len(file) - 4) & ".1.log"
    Name file As Left$(file, Len(file) - 4) & ".1.log"
    Err.Clear
End Sub
