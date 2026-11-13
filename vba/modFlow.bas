Option Explicit
' Status intervals: the traceable base for Flow Time, Flow Load and Flow Efficiency.
' One row per (issue, status visit): when it entered, when it left, how long, which change caused it.

Private Const CHANGELOG_TEXT_MAX As Long = 255
Private Const CHANGELOG_PERSON_FIELDS As String = "assignee,reporter,creator"

Public Function StatusChangesFromChangelog(ByVal changelogRows As Collection, ByVal idToKey As Object, _
                                           Optional ByRef skipped As Long = 0, Optional ByRef shapes As String = "") As Collection
    ' Keeps fieldId = "status". Output rows: Array(issueKey, changeId, createdUtc, fromString, toString).
    ' A change whose date cannot be read is left out and counted in `skipped` (like the Python path, mapping.status_changes):
    ' it cannot be placed in time, and its Empty date made a negative DATE of the duration, which Excel refuses to write.
    ' `shapes` tells what the unreadable values look like (DateShape, no content), to find why they are not read.
    Cov "FLW01"
    Dim out As New Collection, r As Variant, key As String, at As Variant, seen As Object, shape As String, k As Variant
    skipped = 0: shapes = ""
    For Each r In changelogRows
        If CStr(r(5) & "") = "status" Then
            key = CStr(r(0) & "")
            If idToKey.Exists(key) Then key = idToKey(key)
            at = ParseJiraDate(r(2))
            If IsEmpty(at) Then
                Cov "FLW09": skipped = skipped + 1
                If seen Is Nothing Then Set seen = CreateObject("Scripting.Dictionary")
                shape = DateShape(r(2))
                seen(shape) = seen(shape) + 1
            Else
                out.Add Array(key, r(1), at, r(9), r(11))
            End If
        End If
    Next r
    If skipped > 0 Then
        For Each k In seen.Keys
            shapes = shapes & IIf(Len(shapes) > 0, "; ", "") & k & " x " & seen(k)
        Next k
    End If
    Set StatusChangesFromChangelog = out
End Function

Public Function DateShape(ByVal value As Variant) As String
    ' What a value that is not a date looks like, with none of its content (the log must not hold what Jira answered): the
    ' type, then each digit as 9, each letter as A or a, any other ASCII character as it is, 40 characters at most.
    ' "2026-10-01T09:12" gives "String:9999-99-99A99:99"; an epoch in milliseconds "String:9999999999999".
    Dim s As String, i As Long, n As Long, c As Long, out As String, tail As String
    If IsObject(value) Then Cov "FLW11": DateShape = "Object": Exit Function
    If IsNull(value) Or IsEmpty(value) Then Cov "FLW12": DateShape = TypeName(value): Exit Function
    s = CStr(value)
    n = Len(s)
    If n > 40 Then Cov "FLW13": n = 40: tail = "..."
    For i = 1 To n
        c = AscW(Mid$(s, i, 1))
        If c >= 48 And c <= 57 Then
            out = out & "9"
        ElseIf (c >= 65 And c <= 90) Then
            out = out & "A"
        ElseIf c >= 97 And c <= 122 Then
            out = out & "a"
        ElseIf c >= 32 And c < 127 Then
            out = out & Chr$(c)
        Else
            out = out & "?"
        End If
    Next i
    DateShape = TypeName(value) & ":" & out & tail
End Function

Public Function ChangelogHeader() As Variant
    ' The columns of the sheet Changelog (python: formulas.CHANGELOG_RAW).
    ChangelogHeader = Array("Key", "IssueId", "ChangeId", "CreatedUtc", "CreatedText", "FieldId", "Field", "FieldType", "From", _
                            "FromString", "To", "ToString")
End Function

Private Function ScrubMentions(ByVal text As String) As String
    ' "[~accountid:5b10...]" (and the old "[~name]") become "[user]": a mention is a person.
    Dim a As Long, b As Long
    ScrubMentions = text
    a = InStr(ScrubMentions, "[~")
    Do While a > 0
        b = InStr(a, ScrubMentions, "]")
        If b = 0 Then Exit Do
        Cov "FLW17"
        ScrubMentions = Left$(ScrubMentions, a - 1) & "[user]" & Mid$(ScrubMentions, b + 1)
        a = InStr(a + 6, ScrubMentions, "[~")
    Loop
End Function

Private Function CutText(ByVal value As Variant) As String
    ' A text of the changelog: the mentions of a person hidden, then cut at 255 characters (a description can hold thousands).
    CutText = ScrubMentions(CStr(value & ""))
    If Len(CutText) > CHANGELOG_TEXT_MAX Then Cov "FLW16": CutText = Left$(CutText, CHANGELOG_TEXT_MAX)
End Function

Public Function ChangelogRows(ByVal raw As Collection, ByVal idToKey As Object, Optional ByVal userFields As Object = Nothing) As Collection
    ' One row per changelog item, whatever the field (python: mapping.changelog_rows): Key | IssueId | ChangeId | CreatedUtc (Empty when
    ' the date cannot be read) | CreatedText (as Jira sent it) | FieldId | Field | FieldType | From | FromString | To | ToString. No author.
    ' A text is cut at 255 characters; a person field (assignee, reporter, creator) keeps no value: a name is personal data.
    Cov "FLW14"
    Dim out As New Collection, r As Variant, key As String, person As Boolean
    For Each r In raw
        key = CStr(r(0) & "")
        If idToKey.Exists(key) Then key = idToKey(key)
        person = (InStr(1, "," & CHANGELOG_PERSON_FIELDS & ",", "," & CStr(r(5) & "") & ",", vbBinaryCompare) > 0)
        If Not person And Not userFields Is Nothing Then person = userFields.Exists(CStr(r(5) & ""))     ' a field Jira types as a user
        If person Then
            Cov "FLW15"
            out.Add Array(key, CStr(r(0) & ""), r(1), ParseJiraDate(r(2)), CStr(r(2) & ""), r(5), r(6), r(7), "", "", "", "")
        Else
            out.Add Array(key, CStr(r(0) & ""), r(1), ParseJiraDate(r(2)), CStr(r(2) & ""), r(5), r(6), r(7), r(8), CutText(r(9)), r(10), CutText(r(11)))
        End If
    Next r
    Set ChangelogRows = out
End Function

Public Sub LogSkippedDates(ByVal skipped As Long, ByVal shapes As String)
    ' The code J2D-DATA-003 in the log, when status changes were left out (their date cannot be read): how many, and what the
    ' values look like (never the values: they are Jira's).
    If skipped > 0 Then Cov "FLW10": LogLine "WARN", "J2D-DATA-003 | " & skipped & " status changes ignored: their date cannot be read | created: " & shapes
End Sub

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
                      DaysBetween(entered, ch(2)), False, ch(1), check)
        status = CStr(ch(4) & "")
        entered = ch(2)
    Next ch
    seq = seq + 1
    check = ""
    If status <> CStr(info(1) & "") Then Cov "FLW06": check = "CURRENT_MISMATCH(now=" & info(1) & ")"
    out.Add Array(key, seq, status, Category(categories, status), entered, Empty, DaysBetween(entered, asOfUtc), True, Empty, check)
End Sub

Private Function DaysBetween(ByVal fromUtc As Variant, ByVal toUtc As Variant) As Double
    ' Always a Double. A date that could not be read is Empty, and Empty - Date is a DATE in VBA (negative here): Excel refuses
    ' a block holding one with "Out of memory" (error 7). Empty counts as 0, so the row stays flagged BEFORE_CREATED.
    DaysBetween = CDbl(toUtc) - CDbl(fromUtc)
End Function

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
