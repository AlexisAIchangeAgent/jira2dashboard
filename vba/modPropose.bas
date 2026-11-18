Option Explicit
' The button "Propose values from Jira". After the reading of the Jira configuration (modJiraConfig writes the proposal sheets,
' _from_Jira), it copies what fits into the parameter sheets (_from_user), and only into EMPTY cells: what the person typed
' is never changed. Every cell it writes is yellow and the first one of a row has the note "To confirm": the person sees what
' to check. It proposes:
'   Workflows    the statuses the sheet lacks (Workflow = Feature when only the Program project uses the Jira workflow, else
'                Team), the category Jira gives for sure (to do -> Backlog, done -> Done), IsFree, and IsEnd on the only Done status
'                of a workflow that has no IsEnd yet; never IsStart, Active or Wait, Removed: those are human choices;
'   Transitions  the moves the sheet lacks (Workflow, From, To);
'   Iterations   the iterations that the sprints of the scrum boards make (the sprints of the teams that start on the same
'                day, +- 1 day, are one iteration; End = the day after the latest last day: the end is excluded), the ones
'                the sheet lacks;
'   Config       the ids of the fields that Jira names once (Sprint, Team, Story Points, Target start and end, PI, Business
'                Value, Work type category) where the Value is empty. Two candidates for one parameter: nothing proposed, the
'                person chooses. The lists of issue types (SubTaskTypes: the sub-tasks of Jira; FeatureTypes: the types of the
'                Program project with the highest hierarchy level; ParentRequiredTypes and PointsRequiredTypes: the types of the
'                sheet that no Team project has are taken out) and BlocksLinkType (the link type whose outward text is "blocks"),
'                where the Value is empty or still the delivered default (the column Example): see SetListValue.
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
    Dim roleOf As Object, nW As Long, nT As Long, nC As Long, nI As Long
    Set roleOf = RolesOfProjects(wb)
    nW = ProposeWorkflows(wb, roleOf)
    nT = ProposeTransitions(wb, roleOf)
    nC = ProposeFieldIds(wb) + ProposeTypes(wb, roleOf) + ProposeLinkType(wb)
    nI = ProposeIterations(wb)
    If nW + nT + nC + nI = 0 Then
        Cov "PRP03"
        ProposeFromSheets = " | Proposed: nothing new, the parameter sheets already hold what Jira gives"
        Exit Function
    End If
    ProposeFromSheets = " | To confirm, in yellow: the cells below" & _
                        " | " & SheetName("Workflows") & ": " & nW & " proposed" & _
                        " | " & SheetName("Transitions") & ": " & nT & " proposed" & _
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

Public Function KindOf(ByVal projectsText As String, ByVal roleOf As Object) As String
    ' The Workflow of the sheets Workflows and Transitions for a Jira workflow used by these projects ("PRG, TA"): Feature when
    ' every one of them is the Program, else Team.
    Dim p As Variant, key As String, some As Boolean
    KindOf = "Team"
    For Each p In Split(projectsText, ",")
        key = UCase$(TrimWhite(CStr(p)))
        If Len(key) > 0 Then
            If Not roleOf.Exists(key) Then Exit Function
            If StrComp(CStr(roleOf(key)), "Program", vbTextCompare) <> 0 Then Exit Function
            some = True
        End If
    Next p
    If some Then Cov "PRP04": KindOf = "Feature"
End Function

Private Function PairKey(ByVal a As String, ByVal b As String, Optional ByVal c As String = "") As String
    PairKey = LCase$(TrimWhite(a)) & "|" & LCase$(TrimWhite(b)) & "|" & LCase$(TrimWhite(c))
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

Private Function StatusSources(ByVal wb As Object, ByVal roleOf As Object) As Collection
    ' Array(kind, status, category, isFree): the statuses of JiraWorkflows (with the Jira workflow's projects), then those of
    ' JiraStatuses that no workflow gave (a project whose workflows could not be read, or IsUserProjectAdmin = False).
    Dim out As New Collection, seen As Object, row As Variant, kind As String, status As String
    Set seen = CreateObject("Scripting.Dictionary")
    If SheetExists(wb, "JiraWorkflows") Then
        For Each row In ReadTable(wb.Worksheets(SheetName("JiraWorkflows")))
            kind = KindOf(CStr(row("Project") & ""), roleOf)
            status = TrimWhite(CStr(row("Status") & ""))
            If Len(status) > 0 And Not seen.Exists(PairKey(kind, status)) Then
                seen(PairKey(kind, status)) = True
                out.Add Array(kind, status, CStr(row("Category") & ""), row("IsFree"))
            End If
        Next row
    End If
    If SheetExists(wb, "JiraStatuses") Then
        For Each row In ReadTable(wb.Worksheets(SheetName("JiraStatuses")))
            kind = KindOf(CStr(row("Project") & ""), roleOf)
            status = TrimWhite(CStr(row("Status") & ""))
            If Len(status) > 0 And Not seen.Exists(PairKey(kind, status)) Then
                seen(PairKey(kind, status)) = True
                out.Add Array(kind, status, ProposedCategory(CStr(row("JiraCategory") & "")), "")
            End If
        Next row
    End If
    Set StatusSources = out
End Function

Private Function ProposeWorkflows(ByVal wb As Object, ByVal roleOf As Object) As Long
    ' The sheet Workflows: an empty Category or IsFree of a status it has gets Jira's; a status it lacks is added at the end. IsEnd:
    ' the only Done status of a workflow that has no IsEnd = 1 yet gets 1 (two Done statuses: the person chooses).
    Cov "PRP05"
    Dim ws As Object, cWf As Long, cSt As Long, cCat As Long, cFree As Long, cEnd As Long, last As Long, r As Long, existing As Object
    Dim src As Variant, key As String, sources As Collection, endOf As Object, hasEnd As Object
    If Not SheetExists(wb, "Workflows") Then Exit Function
    Set ws = wb.Worksheets(SheetName("Workflows"))
    cWf = HeaderColumn(ws, "Workflow"): cSt = HeaderColumn(ws, "Status"): cCat = HeaderColumn(ws, "Category"): cFree = HeaderColumn(ws, "IsFree")
    cEnd = HeaderColumn(ws, "IsEnd")
    If cWf = 0 Or cSt = 0 Then Exit Function
    Set existing = CreateObject("Scripting.Dictionary")
    Set hasEnd = CreateObject("Scripting.Dictionary")
    last = LastUsedRow(ws, cSt)
    For r = 2 To last
        existing(PairKey(CStr(ws.Cells(r, cWf).Value & ""), CStr(ws.Cells(r, cSt).Value & ""))) = r
        If cEnd > 0 Then
            If Val(CStr(ws.Cells(r, cEnd).Value & "")) = 1 Then hasEnd(LCase$(TrimWhite(CStr(ws.Cells(r, cWf).Value & "")))) = True
        End If
    Next r
    Set sources = StatusSources(wb, roleOf)
    Set endOf = OnlyDoneStatus(sources)
    For Each src In sources
        key = PairKey(CStr(src(0)), CStr(src(1)))
        If existing.Exists(key) Then
            Cov "PRP06"
            r = existing(key)
            ProposeWorkflows = ProposeWorkflows + FillEmpty(ws, r, cCat, src(2)) + FillEmpty(ws, r, cFree, src(3)) + _
                               FillEnd(ws, r, cEnd, src, endOf, hasEnd)
        Else
            Cov "PRP07"
            last = last + 1
            existing(key) = last
            FillEmpty ws, last, cWf, src(0)
            FillEmpty ws, last, cSt, src(1)
            FillEmpty ws, last, cCat, src(2)
            FillEmpty ws, last, cFree, src(3)
            FillEnd ws, last, cEnd, src, endOf, hasEnd
            MarkToConfirm ws.Cells(last, cSt), True
            ProposeWorkflows = ProposeWorkflows + 1
        End If
    Next src
End Function

Private Function OnlyDoneStatus(ByVal sources As Collection) As Object
    ' kind (lower case) -> the status whose category is Done, when the workflow has exactly one; "" when it has several.
    Dim out As Object, src As Variant, kind As String
    Set out = CreateObject("Scripting.Dictionary")
    For Each src In sources
        If StrComp(TrimWhite(CStr(src(2))), "Done", vbTextCompare) = 0 Then
            kind = LCase$(CStr(src(0)))
            If out.Exists(kind) Then Cov "PRP27": out(kind) = "" Else out(kind) = CStr(src(1))
        End If
    Next src
    Set OnlyDoneStatus = out
End Function

Private Function FillEnd(ByVal ws As Object, ByVal r As Long, ByVal cEnd As Long, ByVal src As Variant, ByVal endOf As Object, _
                         ByVal hasEnd As Object) As Long
    ' IsEnd = 1 on the only Done status of a workflow that has no IsEnd yet (1 when written).
    Dim kind As String
    If cEnd = 0 Then Exit Function
    kind = LCase$(CStr(src(0)))
    If Not endOf.Exists(kind) Then Exit Function
    If hasEnd.Exists(kind) Then Exit Function
    If StrComp(CStr(endOf(kind)), CStr(src(1)), vbTextCompare) <> 0 Then Exit Function
    Cov "PRP28"
    FillEnd = FillEmpty(ws, r, cEnd, 1)
    If FillEnd = 1 Then hasEnd(kind) = True
End Function

Private Function ProposeTransitions(ByVal wb As Object, ByVal roleOf As Object) As Long
    ' The sheet Transitions: the moves it lacks (Workflow, From, To) are added at the end.
    Cov "PRP08"
    Dim ws As Object, cWf As Long, cFrom As Long, cTo As Long, last As Long, r As Long, existing As Object, row As Variant
    Dim kind As String, key As String
    If Not SheetExists(wb, "Transitions") Or Not SheetExists(wb, "JiraTransitions") Then Exit Function
    Set ws = wb.Worksheets(SheetName("Transitions"))
    cWf = HeaderColumn(ws, "Workflow"): cFrom = HeaderColumn(ws, "From"): cTo = HeaderColumn(ws, "To")
    If cWf = 0 Or cFrom = 0 Or cTo = 0 Then Exit Function
    Set existing = CreateObject("Scripting.Dictionary")
    last = LastUsedRow(ws, cFrom)
    For r = 2 To last
        existing(PairKey(CStr(ws.Cells(r, cWf).Value & ""), CStr(ws.Cells(r, cFrom).Value & ""), CStr(ws.Cells(r, cTo).Value & ""))) = True
    Next r
    For Each row In ReadTable(wb.Worksheets(SheetName("JiraTransitions")))
        kind = KindOf(CStr(row("Project") & ""), roleOf)
        key = PairKey(kind, CStr(row("From") & ""), CStr(row("To") & ""))
        If Len(TrimWhite(CStr(row("From") & ""))) > 0 And Len(TrimWhite(CStr(row("To") & ""))) > 0 And Not existing.Exists(key) Then
            Cov "PRP09"
            existing(key) = True
            last = last + 1
            FillEmpty ws, last, cWf, kind
            FillEmpty ws, last, cFrom, TrimWhite(CStr(row("From") & ""))
            FillEmpty ws, last, cTo, TrimWhite(CStr(row("To") & ""))
            MarkToConfirm ws.Cells(last, cFrom), True
            ProposeTransitions = ProposeTransitions + 1
        End If
    Next row
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
