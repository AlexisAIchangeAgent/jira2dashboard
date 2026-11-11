Option Explicit
' Local log of the runs, to debug them later: one line per run, per step and per error, in
' %LOCALAPPDATA%\jira2dashboard\logs\jira2dashboard.log (%TEMP% when LOCALAPPDATA is not set). Nothing leaves the computer.
' Each line is opened, written and closed at once: a closed line is in the hands of Windows and survives a crash of
' Excel. After a crash (no VBA error is raised then), the last BEGIN without its END names the step that was running.
' The log never holds the API token, an e-mail address, a URL or an answer of Jira: callers pass step names and numbers,
' and the descriptions of the errors that quote the answer of Jira (602, 607) are not written.
' The logger never raises: a line that cannot be written is dropped.
' Level (Config LogLevel, read by the caller): Steps = everything; Errors = no BEGIN/END lines.
' A run that does not end (Excel closed or crashed) leaves a flag in the registry (HKCU, VB and VBA Program Settings):
' the next run says so in the log.

Public gLogPath As String                          ' empty = the default file; the tests point it to a file of their own
Private Const LOG_MAX_BYTES As Long = 1000000      ' above this size a run starts a new file (the old one becomes .1.log)
Private Const LOG_MAX_TEXT As Long = 300           ' longer texts are cut
Public gLogOpenFake As Boolean                      ' tests: LogOpen records its command in gLogOpenCommand and starts nothing
Public gLogOpenCommand As String
Private mErrorsOnly As Boolean                     ' True while a run asked the level Errors

Public Function LogFolderFrom(ByVal localAppData As String, ByVal temp As String) As String
    Cov "LOG01"
    If Len(localAppData) = 0 Then Cov "LOG02": localAppData = temp
    LogFolderFrom = localAppData & "\jira2dashboard\logs"
End Function

Public Function LogFile() As String
    If Len(gLogPath) > 0 Then LogFile = gLogPath: Exit Function
    LogFile = LogFolderFrom(Environ$("LOCALAPPDATA"), Environ$("TEMP")) & "\jira2dashboard.log"
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
End Sub

Public Sub LogBegin(ByVal stepName As String)
    If mErrorsOnly Then Exit Sub
    LogLine "BEGIN", stepName
End Sub

Public Sub LogEnd(ByVal stepName As String)
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
    mErrorsOnly = False
    DeleteSetting "jira2dashboard", FlagSection(), "RunInProgress"
    Err.Clear
End Sub

Private Function FlagSection() As String
    ' The tests write to a log file of their own: their flag is apart too.
    If Len(gLogPath) = 0 Then FlagSection = "Log" Else FlagSection = "LogTests"
End Function

Public Sub LogError(ByVal where As String, ByVal number As Long, ByVal description As String)
    ' The caller passes the values it saved from Err: any On Error statement clears Err.
    Cov "LOG07"
    Dim text As String
    text = where & " | " & ErrorCode(number) & " | "
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

Public Sub LogOpen()
    ' After an error, shows the log in Notepad when a person is at the keyboard (Excel started by the user, not by a test
    ' or a script). Best effort: never raises. gLogOpenFake lets the tests read the command instead of starting Notepad.
    Dim cmd As String
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

Public Function ErrorCode(ByVal number As Long) As String
    ' The errors raised by this project read vbObjectError+n; the others keep their number.
    If number >= vbObjectError And number < vbObjectError + 65536 Then
        Cov "LOG09"
        ErrorCode = "vbObjectError+" & CStr(number - vbObjectError)
    Else
        ErrorCode = CStr(number)
    End If
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
