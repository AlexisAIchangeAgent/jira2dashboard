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

Public Function KeyChangeCounts(ByVal changelogRows As Collection, ByVal idToKey As Object) As Object
    ' Moves between projects: a changelog row of the field Key. Returns issueKey -> number of changes (absent = none).
    Cov "FLW08"
    Dim out As Object, r As Variant, key As String
    Set out = CreateObject("Scripting.Dictionary")
    For Each r In changelogRows
        If CStr(r(6) & "") = "Key" Or CStr(r(5) & "") = "issuekey" Then
            key = CStr(r(0) & "")
            If idToKey.Exists(key) Then key = idToKey(key)
            If out.Exists(key) Then out(key) = out(key) + 1 Else out(key) = 1
        End If
    Next r
    Set KeyChangeCounts = out
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
