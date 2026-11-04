"""Patch 1: clear swallowed errors, language-safe booleans in tests, tests for missed markers."""
from pathlib import Path

D = Path(__file__).parent


def rep(name, old, new):
    p = D / name
    s = p.read_text(encoding="ascii").replace("\r\n", "\n")
    assert s.count(old) == 1, (name, old[:60])
    p.write_text(s.replace(old, new), encoding="ascii", newline="\n")


rep("modHttp.bas", r"""    On Error Resume Next
    SafeHeader = req.GetResponseHeader(name)
End Function""", r"""    On Error Resume Next
    SafeHeader = req.GetResponseHeader(name)
    Err.Clear                                   ' a missing header is normal: leave no stale error
End Function""")

rep("modText.bas", r"""    upper = UBound(bytes)                              ' fails on an uninitialised (empty) array
    If Err.Number <> 0 Then Exit Function""", r"""    upper = UBound(bytes)                              ' fails on an uninitialised (empty) array
    If Err.Number <> 0 Then Err.Clear: Exit Function   ' expected: leave no stale error""")

rep("modTests.bas", r"""out(3)(2) & "|" & out(3)(3) & "|" & out(3)(6) & "|" & out(3)(7)""",
    r"""out(3)(2) & "|" & out(3)(3) & "|" & out(3)(6) & "|" & YesNo(out(3)(7))""")
rep("modTests.bas", r"""out(4)(2) & "|" & out(4)(6) & "|" & out(4)(7)""",
    r"""out(4)(2) & "|" & out(4)(6) & "|" & YesNo(out(4)(7))""")
rep("modTests.bas", r"""Private Function CountOf(""", r"""Private Function YesNo(ByVal b As Boolean) As String
    ' Never write Boolean & "": French VBA gives "Vrai"/"Faux".
    YesNo = IIf(b, "True", "False")
End Function

Private Function CountOf(""")

rep("modTests.bas", r"""        On Error Resume Next
        Application.Run "modTests." & nm""", r"""        On Error Resume Next
        Err.Clear
        Application.Run "modTests." & nm""")

rep("modTests.bas", r"""For Each bad In Array("{", "{'a':1}x",""",
    r"""For Each bad In Array("", "[1,", "{'a':1 'b':2}", "{", "{'a':1}x",""")

rep("modTests.bas", r"""        "T_StatusChanges", "T_RetryDelay", "T_SafeAccess", "T_FilesUtf8")""",
    r"""        "T_StatusChanges", "T_RetryDelay", "T_SafeAccess", "T_FilesUtf8", "T_StringifyBadObject", _
        "T_TooManyPages", "T_BackgroundExport")""")

rep("modTests.bas", r"""' ================= Network (public test API, not Jira) =================""",
    r"""Public Sub T_StringifyBadObject()
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
    ' Real Option B run: a hidden PowerShell writes a page and done.flag; Excel is never blocked.
    Dim folder As String, script As String, st As Object, t0 As Double
    folder = Environ$("TEMP") & "\xlsm_vba_bg"
    If Len(Dir$(folder, vbDirectory)) = 0 Then MkDir folder
    script = folder & "\fake_export.ps1"
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8": st.Open
    st.WriteText "param([string]$OutDir)" & vbCrLf & "Start-Sleep -Seconds 2" & vbCrLf & _
        "Set-Content -Path (Join-Path $OutDir 'page-0001.json') -Value '{""ok"":true}' -Encoding UTF8" & vbCrLf & _
        "New-Item -ItemType File -Path (Join-Path $OutDir 'done.flag') -Force | Out-Null" & vbCrLf
    st.SaveToFile script, 2: st.Close
    StartBackgroundExport script, folder, "", "-NoProfile -ExecutionPolicy Bypass", 10
    Check "background", "state after start", "start", "RUNNING", gExportStatus
    Check "background", "not done at once", "start", False, IsExportDone(folder)
    t0 = Timer
    Do While Not IsExportDone(folder) And Timer - t0 < 30
        WaitMs 250
    Loop
    CheckBackgroundExport
    Check "background", "done flag seen", "after wait", "DONE", gExportStatus
    Check "background", "page imported", "page-0001.json", True, IsTrue(JsonParse(ReadUtf8File(folder & "\page-0001.json"))("ok"))
    CancelBackgroundCheck
    StartBackgroundExport folder & "\does_not_exist.ps1", folder & "\empty", "", "-NoProfile", 0
    WaitMs 1100
    CheckBackgroundExport
    Check "background", "timeout detected", "timeout 0 min", "TIMEOUT", gExportStatus
    CancelBackgroundCheck
End Sub

' ================= Network (public test API, not Jira) =================""")
print("patched")
