"""Patch 2: Option B test without any child process (antivirus blocks Office -> PowerShell)."""
from pathlib import Path

D = Path(__file__).parent
p = D / "modTests.bas"
s = p.read_text(encoding="ascii").replace("\r\n", "\n")
start = s.index("Public Sub T_BackgroundExport()")
end = s.index("' ================= Network (public test API, not Jira) =================")
new = r"""Public Sub T_BackgroundExport()
    ' Option B without any process: simulate what the external PowerShell script writes.
    Dim folder As String, names As Object, issues As Collection
    folder = Environ$("TEMP") & "\xlsm_vba_bg"
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

"""
p.write_text(s[:start] + new + s[end:], encoding="ascii", newline="\n")
print("patched")
