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
