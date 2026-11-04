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
