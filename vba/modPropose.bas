Option Explicit
' The button "Propose values from Jira". After the reading of the Jira configuration (modJiraConfig writes the proposal sheets,
' _from_Jira), it copies what fits into the parameter sheets (_from_user), and only into EMPTY cells: what the person typed
' is never changed. Every cell it writes is yellow and the first one of a row has the note "To confirm": the person sees what
' to check. It proposes:
'   Iterations   the iterations that the sprints of the scrum boards make (the sprints of the teams that start on the same
'                day, +- 1 day, are one iteration; End = the day after the latest last day: the end is excluded), the ones
'                the sheet lacks;
'   Config       the ids of the fields that Jira names once (Sprint, Team, Story Points, Target start and end, PI, Business
'                Value, Work type category) where the Value is empty. Two candidates for one parameter: nothing proposed, the
'                person chooses. The lists of issue types (SubTaskTypes: the sub-tasks of Jira; FeatureTypes: the types of the
'                Program project with the highest hierarchy level; ParentRequiredTypes and PointsRequiredTypes: the types of the
'                sheet that no Team project has are taken out) and BlocksLinkType (the link type whose outward text is "blocks"),
'                where the Value is empty or still the delivered default (the column Example): see SetListValue.
' It no longer copies statuses or moves into the sheets Workflows and Transitions (AL, 2026-11-18): the metrics read the statuses, their categories, the start
' and the end of each flow, the workflow of each issue type and its moves from the _from_Jira sheets (flow/status_rules.py); Workflows only holds the exceptions
' the person types (Wait, Removed, a Done status that is no delivery), Transitions the moves the person adds to Jira's.
' This module makes no call to Jira: it reads and writes the sheets of the workbook.

Public Const TO_CONFIRM_NOTE As String = "To confirm: proposed from Jira"

Public Function ProposeValues(ByVal cfg As Object, ByVal keys As Collection, ByVal wb As Object, _
                              ByVal withWorkflows As Boolean) As String
    ' ReadJiraConfiguration (the proposal sheets), then the proposals in the parameter sheets: the summary for the status cell.
    Cov "PRP01"
    Dim summary As String, sprints As Collection, noBoard As Long
    summary = ReadJiraConfiguration(cfg, keys, wb, withWorkflows)
    Set sprints = ReadSprints(cfg, keys, noBoard)                  ' read only: the agile API (Jira Software)
    WriteSprints wb, sprints
    WriteLinkTypes wb, ReadLinkTypes(cfg)
    summary = summary & " | Sprints read: " & sprints.Count
    If noBoard > 0 Then Cov "PRP15": summary = summary & " (" & noBoard & " project(s) with no scrum board or sprint that you can read)"
    ProposeValues = "Values proposed from Jira" & Mid$(summary, InStr(summary, " | ")) & ProposeFromSheets(wb)
End Function

Public Function ProposeFromSheets(ByVal wb As Object) As String
    ' The proposals of the sheets already in wb; " | Name: value" fields for the status cell.
    Cov "PRP02"
    Dim roleOf As Object, nC As Long, nI As Long
    Set roleOf = RolesOfProjects(wb)
    nC = ProposeFieldIds(wb) + ProposeTypes(wb, roleOf) + ProposeLinkType(wb)
    nI = ProposeIterations(wb)
    If nC + nI = 0 Then
        Cov "PRP03"
        ProposeFromSheets = " | Proposed: nothing new, the parameter sheets already hold what Jira gives"
        Exit Function
    End If
    ProposeFromSheets = " | To confirm, in yellow: the cells below" & _
                        " | " & SheetName("Config") & ": " & nC & " proposed" & _
                        " | " & SheetName("Iterations") & ": " & nI & " proposed"
End Function

Private Function RolesOfProjects(ByVal wb As Object) As Object
    ' project key (upper case) -> Program or Team, from the sheet Projects.
    Dim out As Object, row As Variant
    Set out = CreateObject("Scripting.Dictionary")
    If SheetExists(wb, "Projects") Then
        For Each row In ReadTable(wb.Worksheets(SheetName("Projects")))
            out(UCase$(TrimWhite(CStr(row("Project") & "")))) = TrimWhite(CStr(row("Role") & ""))
        Next row
    End If
    Set RolesOfProjects = out
End Function

Private Function LastUsedRow(ByVal ws As Object, ByVal col As Long) As Long
    LastUsedRow = ws.Cells(ws.Rows.Count, col).End(-4162).Row            ' -4162 = xlUp
    If LastUsedRow < 1 Then LastUsedRow = 1
End Function

Private Sub MarkToConfirm(ByVal cell As Object, ByVal withNote As Boolean)
    ' Light yellow; the note on the first cell of a row only (a note on every cell would bury the sheet).
    cell.Interior.Color = RGB(255, 242, 204)
    If withNote Then
        On Error Resume Next
        If cell.Comment Is Nothing Then Cov "PRP14": cell.AddComment TO_CONFIRM_NOTE
        On Error GoTo 0
    End If
End Sub

Private Function FillEmpty(ByVal ws As Object, ByVal r As Long, ByVal col As Long, ByVal value As Variant) As Long
    ' Writes value in the cell when it is empty and value is not; 1 when written.
    If col = 0 Then Exit Function
    If Len(CStr(value & "")) = 0 Then Exit Function
    If Len(TrimWhite(CStr(ws.Cells(r, col).Value & ""))) > 0 Then Exit Function
    Cov "PRP13"
    ws.Cells(r, col).Value = SafeCell(value)
    MarkToConfirm ws.Cells(r, col), False
    FillEmpty = 1
End Function

Private Function ProposeFieldIds(ByVal wb As Object) As Long
    ' The sheet Config: an empty Value of a ...Field parameter gets the id of the one Jira field that fits (by the type of the
    ' custom field, else by its name). Two fields that fit: nothing, the person chooses. The ids come from JiraFields.
    Cov "PRP10"
    Dim ws As Object, cParam As Long, cValue As Long, fields As Collection, rule As Variant, r As Long, id As String
    Dim done As Object, n As Long
    If Not SheetExists(wb, "Config") Or Not SheetExists(wb, "JiraFields") Then Exit Function
    Set ws = wb.Worksheets(SheetName("Config"))
    cParam = HeaderColumn(ws, "Parameter"): cValue = HeaderColumn(ws, "Value")
    If cParam = 0 Or cValue = 0 Then Exit Function
    Set fields = ReadTable(wb.Worksheets(SheetName("JiraFields")))
    Set done = CreateObject("Scripting.Dictionary")
    ' parameter | how to find the field | text (the type first, the name when the type says nothing): "detail" = the custom type contains the text, "name" = the name is the text
    For Each rule In Array(Array("SprintField", "detail", "com.pyxis.greenhopper.jira:gh-sprint"), _
                           Array("TeamField", "detail", "rm-teams-custom-field-team"), _
                           Array("StoryPointsField", "detail", "jsw-story-points"), _
                           Array("TargetStartField", "detail", "jpo-custom-field-baseline-start"), _
                           Array("TargetEndField", "detail", "jpo-custom-field-baseline-end"), _
                           Array("SprintField", "name", "Sprint"), _
                           Array("TeamField", "name", "Team"), _
                           Array("StoryPointsField", "name", "Story point estimate"), _
                           Array("StoryPointsField", "name", "Story Points"), _
                           Array("TargetStartField", "name", "Target start"), _
                           Array("TargetEndField", "name", "Target end"), _
                           Array("PIField", "name", "PI"), _
                           Array("BusinessValueField", "name", "Business Value"), _
                           Array("WorkTypeField", "name", "Work type category"), _
                           Array("WorkTypeField", "name", "TypeCategorie"))
        If Not done.Exists(rule(0)) Then
            n = FieldsThatFit(fields, CStr(rule(1)), CStr(rule(2)), id)
            If n = 1 Then
                r = FindRow(ws, cParam, CStr(rule(0)))
                done(rule(0)) = True
                If r > 0 Then
                    If FillEmpty(ws, r, cValue, id) = 1 Then
                        Cov "PRP11"
                        MarkToConfirm ws.Cells(r, cValue), True
                        ProposeFieldIds = ProposeFieldIds + 1
                    End If
                End If
            ElseIf n > 1 Then
                Cov "PRP12": done(rule(0)) = True                     ' two fields fit: the person chooses
            End If
        End If
    Next rule
End Function

Private Function ListKey(ByVal text As String) As String
    ' A list of names as a comparison key: no spaces, lower case ("Story, Task" = "story,task").
    ListKey = LCase$(Replace(TrimWhite(text), " ", ""))
End Function

Private Function SetListValue(ByVal ws As Object, ByVal cParam As Long, ByVal cValue As Long, ByVal cExample As Long, _
                              ByVal param As String, ByVal proposed As String) As Long
    ' The Value of a Config parameter gets what Jira says when it is EMPTY or still holds the delivered default (the same text as
    ' its Example: the person never chose it). A value the person typed, whatever it is, stays; so does a value Jira agrees with.
    ' Yellow, with the note. 1 when written.
    Dim r As Long, current As String
    If Len(proposed) = 0 Then Exit Function
    r = FindRow(ws, cParam, param)
    If r = 0 Then Exit Function
    current = TrimWhite(CStr(ws.Cells(r, cValue).Value & ""))
    If Len(current) > 0 Then
        If ListKey(current) <> ListKey(CStr(ws.Cells(r, cExample).Value & "")) Then Cov "PRP22": Exit Function
        If ListKey(current) = ListKey(proposed) Then Cov "PRP23": Exit Function
    End If
    Cov "PRP24"
    ws.Cells(r, cValue).Value = SafeCell(proposed)
    MarkToConfirm ws.Cells(r, cValue), True
    SetListValue = 1
End Function

Private Function NamesText(ByVal names As Object) As String
    ' The names of a dictionary (lower case -> name) as a list "A,B".
    Dim k As Variant
    For Each k In names.Keys
        NamesText = NamesText & IIf(Len(NamesText) > 0, ",", "") & CStr(names(k))
    Next k
End Function

Private Function KeepKnown(ByVal current As String, ByVal known As Object) As String
    ' The names of the list `current` that the dictionary knows (lower case -> name), in the order of the list.
    Dim p As Variant, key As String
    For Each p In Split(current, ",")
        key = LCase$(TrimWhite(CStr(p)))
        If known.Exists(key) Then KeepKnown = KeepKnown & IIf(Len(KeepKnown) > 0, ",", "") & CStr(known(key))
    Next p
End Function

Private Function ProposeTypes(ByVal wb As Object, ByVal roleOf As Object) As Long
    ' The sheet Config: the lists of issue types, from JiraIssueTypes (Project | IssueType | Workflow | Subtask | HierarchyLevel |
    ' ProjectStyle). SubTaskTypes: the types Jira calls sub-tasks. FeatureTypes: the types of the Program project (not sub-tasks)
    ' at its highest hierarchy level. ParentRequiredTypes and PointsRequiredTypes are rules, not facts: only the names that no Team
    ' project has are taken out of the list. Nothing is proposed from an empty list (no Team project, no sub-task type).
    Cov "PRP21"
    Dim ws As Object, cParam As Long, cValue As Long, cExample As Long, row As Variant, subs As Object, top As Object, team As Object
    Dim name As String, role As String, level As Double, best As Double, r As Long
    If Not SheetExists(wb, "Config") Or Not SheetExists(wb, "JiraIssueTypes") Then Exit Function
    Set ws = wb.Worksheets(SheetName("Config"))
    cParam = HeaderColumn(ws, "Parameter"): cValue = HeaderColumn(ws, "Value"): cExample = HeaderColumn(ws, "Example")
    If cParam = 0 Or cValue = 0 Or cExample = 0 Then Exit Function
    Set subs = CreateObject("Scripting.Dictionary"): Set top = CreateObject("Scripting.Dictionary"): Set team = CreateObject("Scripting.Dictionary")
    best = -1000
    For Each row In ReadTable(wb.Worksheets(SheetName("JiraIssueTypes")))
        name = TrimWhite(CStr(row("IssueType") & ""))
        If Len(name) > 0 Then
            If StrComp(TrimWhite(CStr(row("Subtask") & "")), "Yes", vbTextCompare) = 0 Then
                subs(LCase$(name)) = name
            Else
                role = ""
                If roleOf.Exists(UCase$(TrimWhite(CStr(row("Project") & "")))) Then role = CStr(roleOf(UCase$(TrimWhite(CStr(row("Project") & "")))))
                If StrComp(role, "Program", vbTextCompare) = 0 Then
                    level = Val(CStr(row("HierarchyLevel") & ""))
                    If level > best Then best = level: top.RemoveAll
                    If level = best Then top(LCase$(name)) = name
                ElseIf StrComp(role, "Team", vbTextCompare) = 0 Then
                    team(LCase$(name)) = name
                End If
            End If
        End If
    Next row
    ProposeTypes = SetListValue(ws, cParam, cValue, cExample, "SubTaskTypes", NamesText(subs)) + _
                   SetListValue(ws, cParam, cValue, cExample, "FeatureTypes", NamesText(top))
    If team.Count > 0 Then
        For Each row In Array("ParentRequiredTypes", "PointsRequiredTypes")
            r = FindRow(ws, cParam, CStr(row))
            If r > 0 Then
                ProposeTypes = ProposeTypes + SetListValue(ws, cParam, cValue, cExample, CStr(row), _
                               KeepKnown(CStr(ws.Cells(r, cValue).Value & ""), team))
            End If
        Next row
    End If
End Function

Private Function ProposeLinkType(ByVal wb As Object) As Long
    ' BlocksLinkType: the name of the one link type whose outward text is "blocks" (JiraLinkTypes: Name | Inward | Outward). Two
    ' of them, or none: nothing, the person chooses.
    Cov "PRP25"
    Dim ws As Object, cParam As Long, cValue As Long, cExample As Long, row As Variant, found As Long, name As String
    If Not SheetExists(wb, "Config") Or Not SheetExists(wb, "JiraLinkTypes") Then Exit Function
    Set ws = wb.Worksheets(SheetName("Config"))
    cParam = HeaderColumn(ws, "Parameter"): cValue = HeaderColumn(ws, "Value"): cExample = HeaderColumn(ws, "Example")
    If cParam = 0 Or cValue = 0 Or cExample = 0 Then Exit Function
    For Each row In ReadTable(wb.Worksheets(SheetName("JiraLinkTypes")))
        If StrComp(TrimWhite(CStr(row("Outward") & "")), "blocks", vbTextCompare) = 0 Then name = TrimWhite(CStr(row("Name") & "")): found = found + 1
    Next row
    If found > 1 Then Cov "PRP26": Exit Function
    ProposeLinkType = SetListValue(ws, cParam, cValue, cExample, "BlocksLinkType", name)
End Function

Private Function IsoDay(ByVal serial As Double) As String
    ' 2026-01-15, whatever the regional settings (the sheets read ISO text as a date).
    IsoDay = Year(serial) & "-" & Right$("0" & Month(serial), 2) & "-" & Right$("0" & Day(serial), 2)
End Function

Private Function ProposeIterations(ByVal wb As Object) As Long
    ' The sheet Iterations: one row per iteration that the sprints of JiraSprints make and that the sheet does not hold yet. The
    ' sprints of several teams that start on the same day (+- 1 day) are one iteration, whatever their length: Start = the
    ' earliest start, End = the day after the latest last day (the end is excluded), never later than the start of the next
    ' iteration (two iterations must not hold the same day: a ticket done that day would count twice). A cluster that would
    ' end before it starts is not proposed (the build refuses such a sheet: error 704). Label: the name the sprints share,
    ' else "Iteration <Start>".
    Cov "PRP16"
    Dim ws As Object, cLabel As Long, cStart As Long, cEnd As Long, last As Long, r As Long
    Dim startS As New Collection, endS As New Collection, names As New Collection, labels As Object
    Dim row As Variant, s As Variant, e As Variant, i As Long, k As Long, n As Long, count As Long
    Dim labelOf() As String, loOf() As Double, hiOf() As Double
    If Not SheetExists(wb, "Iterations") Or Not SheetExists(wb, "JiraSprints") Then Exit Function
    Set ws = wb.Worksheets(SheetName("Iterations"))
    cLabel = HeaderColumn(ws, "Iteration-label"): cStart = HeaderColumn(ws, "Start"): cEnd = HeaderColumn(ws, "End")
    If cLabel = 0 Or cStart = 0 Or cEnd = 0 Then Exit Function
    ' the sprints with dates, oldest first (insertion sort: a few dozen)
    For Each row In ReadTable(wb.Worksheets(SheetName("JiraSprints")))
        s = ParseJiraDate(row("Start")): e = ParseJiraDate(row("End"))
        If Not IsEmpty(s) And Not IsEmpty(e) Then
            For i = 1 To startS.Count
                If Int(CDbl(s)) < startS(i) Then Exit For
            Next i
            If i > startS.Count Then
                startS.Add Int(CDbl(s)): endS.Add Int(CDbl(e)): names.Add CStr(row("Sprint") & "")
            Else
                startS.Add Int(CDbl(s)), Before:=i: endS.Add Int(CDbl(e)), Before:=i: names.Add CStr(row("Sprint") & ""), Before:=i
            End If
        End If
    Next row
    If startS.Count = 0 Then Exit Function
    ReDim labelOf(1 To startS.Count): ReDim loOf(1 To startS.Count): ReDim hiOf(1 To startS.Count)
    ' the clusters: the sprints that start within a day of the first one of the cluster
    i = 1
    Do While i <= startS.Count
        count = count + 1
        loOf(count) = startS(i): hiOf(count) = endS(i): labelOf(count) = names(i)
        k = i + 1
        Do While k <= startS.Count
            If Abs(startS(k) - startS(i)) > 1 Then Exit Do
            If endS(k) > hiOf(count) Then hiOf(count) = endS(k)
            If names(k) <> labelOf(count) Then labelOf(count) = ""
            k = k + 1
        Loop
        i = k
    Loop
    For n = 1 To count
        hiOf(n) = hiOf(n) + 1                                             ' the end is excluded
        If n < count Then
            If hiOf(n) > loOf(n + 1) Then Cov "PRP19": hiOf(n) = loOf(n + 1)
        End If
        If Len(labelOf(n)) = 0 Then labelOf(n) = "Iteration " & IsoDay(loOf(n))
    Next n
    ' what the sheet already holds: the labels and the days
    Set labels = CreateObject("Scripting.Dictionary")
    last = LastUsedRow(ws, cLabel)
    For r = 2 To last
        labels(LCase$(TrimWhite(CStr(ws.Cells(r, cLabel).Value & "")))) = True
    Next r
    For n = 1 To count
        If hiOf(n) <= loOf(n) Then
            Cov "PRP20"                                                    ' dates that make no iteration: not proposed
        ElseIf Not IterationHeld(ws, cStart, cEnd, last, loOf(n)) And Not labels.Exists(LCase$(labelOf(n))) Then
            Cov "PRP17"
            last = last + 1
            labels(LCase$(labelOf(n))) = True
            FillEmpty ws, last, cLabel, labelOf(n)
            FillEmpty ws, last, cStart, IsoDay(loOf(n))
            FillEmpty ws, last, cEnd, IsoDay(hiOf(n))
            MarkToConfirm ws.Cells(last, cLabel), True
            ProposeIterations = ProposeIterations + 1
        End If
    Next n
End Function

Private Function IterationHeld(ByVal ws As Object, ByVal cStart As Long, ByVal cEnd As Long, ByVal last As Long, ByVal day As Double) As Boolean
    ' True when an iteration of the sheet (Start <= day < End) already holds this day.
    Dim r As Long, s As Variant, e As Variant
    For r = 2 To last
        s = SerialOf(ws.Cells(r, cStart).Value): e = SerialOf(ws.Cells(r, cEnd).Value)
        If Not IsEmpty(s) And Not IsEmpty(e) Then
            If day >= s And day < e Then Cov "PRP18": IterationHeld = True: Exit Function
        End If
    Next r
End Function

Private Function FieldsThatFit(ByVal fields As Collection, ByVal how As String, ByVal text As String, ByRef id As String) As Long
    ' How many fields of JiraFields (Id | Name | Custom | SchemaType | SchemaDetail) fit; id is the last one that did.
    Dim f As Variant, fits As Boolean
    For Each f In fields
        If how = "detail" Then
            fits = (InStr(1, CStr(f("SchemaDetail") & ""), text, vbTextCompare) > 0)
        Else
            fits = (StrComp(TrimWhite(CStr(f("Name") & "")), text, vbTextCompare) = 0)
        End If
        If fits Then id = CStr(f("Id") & ""): FieldsThatFit = FieldsThatFit + 1
    Next f
End Function
