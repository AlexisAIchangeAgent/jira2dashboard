"""Patch 6: write each result row at once (no block at the end) and drive a progress bar (shapes pbBack/pbFill)."""
from pathlib import Path

p = Path(__file__).parent / "modTests.bas"
s = p.read_text(encoding="ascii").replace("\r\n", "\n")


def rep(old, new):
    global s
    assert s.count(old) == 1, old[:60]
    s = s.replace(old, new)


rep("""Private mRows As Collection
Private mCurrent As String
""", """Private mRows As Collection
Private mCurrent As String
Private mTotalPlanned As Long      ' tests planned for this run (progress bar total)
Private mDone As Long              ' tests finished in this run
Private mNextRow As Long           ' next free row in TestResults
""")

rep("""    mRows.Add Array(mCurrent & "." & n, area, description, inputText, Repr(expected), Repr(actual), _
                    IIf(SameValue(expected, actual), "OK", "KO"), Now)""",
    """    AddRow Array(mCurrent & "." & n, area, description, inputText, Repr(expected), Repr(actual), _
                 IIf(SameValue(expected, actual), "OK", "KO"), Now)""")

start = s.index("Public Function RunAllTests() As String")
end = s.index("' ================= JSON =================")
s = s[:start] + r'''Public Function RunAllTests() As String
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

''' + s[end:]

rep("""Public Function RunWithCoverage(ByVal includeNetwork As Boolean) As String
    Dim s As String
    CovStart""", """Public Function RunWithCoverage(ByVal includeNetwork As Boolean) As String
    Dim s As String
    mTotalPlanned = CountOfNames(TestNames()) + IIf(includeNetwork, CountOfNames(NetworkTestNames()), 0)
    mDone = 0
    ProgressDraw 0, mTotalPlanned
    CovStart""")
rep("""    CovWrite ThisWorkbook.Worksheets("CoverageHits")
    RunWithCoverage = s""", """    CovWrite ThisWorkbook.Worksheets("CoverageHits")
    mTotalPlanned = 0
    RunWithCoverage = s""")
assert "WriteResults" not in s
p.write_text(s, encoding="ascii", newline="\n")

b = Path(__file__).parent / "modButtons.bas"
t = b.read_text(encoding="ascii").replace("\r\n", "\n")
old = """    ThisWorkbook.Worksheets("CoverageHits").Cells.ClearContents
End Sub"""
assert t.count(old) == 1
t = t.replace(old, """    ThisWorkbook.Worksheets("CoverageHits").Cells.ClearContents
    ProgressReset
End Sub""")
b.write_text(t, encoding="ascii", newline="\n")
print("patched")
