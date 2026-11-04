Option Explicit
' Extraction of the flow data: Jira -> the raw sheets Issues, StatusIntervals and IssueLinks (docs section 2.1).
' VBA writes raw tables only; every figure (ItemFlow, DailyLoad, Metrics, DataChecks) is an Excel formula that reads them.

Public Sub ExtractFlowData()
    ' Everything comes from the workbook: Config (connection, proxy, Filter, field ids, AsOf), Projects, Workflows.
    Dim wsConfig As Object
    Set wsConfig = ThisWorkbook.Worksheets("Config")
    ExtractFlowDataTo JiraConfigFromSheet(wsConfig), ExtractionJql(ThisWorkbook.Worksheets("Projects"), wsConfig), _
                      ConfigValue(wsConfig, "WorkTypeField"), AsOfFromConfig(wsConfig), _
                      ThisWorkbook.Worksheets("Projects"), ThisWorkbook.Worksheets("Workflows"), _
                      SheetNamed("Issues"), SheetNamed("StatusIntervals"), True, FieldSettings(wsConfig), SheetNamed("IssueLinks")
End Sub

Private Function SheetNamed(ByVal name As String) As Object
    ' The sheet, created at the end of the workbook if it is not there yet.
    Dim ws As Object
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = name Then Set SheetNamed = ws: Exit Function
    Next ws
    Set SheetNamed = ThisWorkbook.Worksheets.Add(, ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    SheetNamed.Name = name
End Function

Private Function FieldParameters() As Variant
    FieldParameters = Array("StoryPointsField", "BusinessValueField", "PIField", "TeamField", "SprintField", _
                            "TargetStartField", "TargetEndField")
End Function

Public Function FieldSettings(ByVal wsConfig As Object) As Object
    ' The Jira fields of the data checks and the link type that means "blocks", read from Config (empty = not asked).
    Dim out As Object, nm As Variant
    Set out = CreateObject("Scripting.Dictionary")
    For Each nm In FieldParameters()
        out(CStr(nm)) = ConfigValue(wsConfig, CStr(nm))
    Next nm
    out("BlocksLinkType") = ConfigValue(wsConfig, "BlocksLinkType")
    Set FieldSettings = out
End Function

Private Function Setting(ByVal settings As Object, ByVal name As String) As String
    If settings Is Nothing Then Exit Function
    If settings.Exists(name) Then Setting = CStr(settings(name))
End Function

Private Function WorkTypeOf(ByVal issue As Variant, ByVal fieldId As String) As Variant
    ' The work type (TypeCategorie). A select field comes as an object {"value": "..."}; empty if not set.
    Cov "MEX02"
    Dim v As Variant, path As String
    If Len(fieldId) = 0 Then Exit Function
    path = "fields." & fieldId
    If IsObject(JPath(issue, path)) Then
        Cov "MEX03": v = JGet(JPath(issue, path), "value")
    Else
        v = JPath(issue, path)
    End If
    If Not IsNull(v) Then WorkTypeOf = v
End Function

Private Function FieldValue(ByVal issue As Variant, ByVal fieldId As String) As Variant
    ' The raw value of a field by its id; Empty when no id is configured or the field is missing or null.
    If Len(fieldId) = 0 Then Exit Function
    Dim v As Variant
    If IsObject(JPath(issue, "fields." & fieldId)) Then          ' a team or a sprint list is an object: needs Set
        Set FieldValue = JPath(issue, "fields." & fieldId)
        Exit Function
    End If
    v = JPath(issue, "fields." & fieldId)
    If Not IsNull(v) Then FieldValue = v
End Function

Private Function TextOf(ByVal v As Variant) As Variant
    ' A text field, or an object (Team: {"name": ...} or a select {"value": ...}): the text, else Empty.
    Cov "MEX04"
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If IsObject(v) Then
        TextOf = JGet(v, "name")
        If IsEmpty(TextOf) Or IsNull(TextOf) Then TextOf = JGet(v, "value")
        If IsNull(TextOf) Then TextOf = Empty
        Exit Function
    End If
    TextOf = CStr(v)
End Function

Private Function NumberOf(ByVal v As Variant) As Variant
    If IsEmpty(v) Or IsObject(v) Then Exit Function
    If IsNumeric(v) Then NumberOf = CDbl(v)
End Function

Private Function DateOf(ByVal v As Variant) As Variant
    If IsEmpty(v) Or IsObject(v) Then Exit Function
    DateOf = ParseJiraDate(CStr(v))
End Function

Private Function VersionsOf(ByVal issue As Variant) As Variant
    ' Names of the fix versions, joined; Empty when there is none.
    Dim v As Variant, out As String
    For Each v In AsList(JPath(issue, "fields.fixVersions"))
        If Len(out) > 0 Then out = out & ", "
        out = out & CStr(JGet(v, "name") & "")
    Next v
    If Len(out) > 0 Then VersionsOf = out
End Function

Private Function SprintRank(ByVal state As String) As Long
    Select Case state
        Case "active": SprintRank = 3
        Case "future": SprintRank = 2
        Case "closed": SprintRank = 1
    End Select
End Function

Private Function SprintInfo(ByVal sprints As Variant) As Variant
    ' Array(state, start, end, count). The sprint that counts is the active one, else a future one, else the last closed
    ' one; `count` is the number of sprints in the field. With no sprint: Array(Empty, Empty, Empty, 0).
    Cov "MEX05"
    Dim sp As Variant, best As Variant, rank As Long, bestRank As Long, n As Long, state As String
    best = Array(Empty, Empty, Empty, 0)
    If IsEmpty(sprints) Then SprintInfo = best: Exit Function
    For Each sp In AsList(sprints)
        n = n + 1
        state = CStr(JGet(sp, "state") & "")
        rank = SprintRank(state)
        If rank > 0 And rank >= bestRank Then
            bestRank = rank
            best = Array(state, ParseJiraDate(JGet(sp, "startDate")), ParseJiraDate(JGet(sp, "endDate")), 0)
        End If
    Next sp
    best(3) = n
    SprintInfo = best
End Function

Public Function ExtractFields(ByVal workTypeField As String, Optional ByVal settings As Object = Nothing) As String
    ' The Jira fields the extraction reads: the rest (comments, descriptions, attachments) is not downloaded.
    Dim out As String, nm As Variant
    out = "key,project,issuetype,created,status,parent,resolutiondate,fixVersions,issuelinks"
    If Len(workTypeField) > 0 Then out = out & "," & workTypeField
    For Each nm In FieldParameters()
        If Len(Setting(settings, CStr(nm))) > 0 Then out = out & "," & Setting(settings, CStr(nm))
    Next nm
    ExtractFields = out
End Function

Private Function KeyChangesOf(ByVal moved As Object, ByVal key As String) As Long
    If moved.Exists(key) Then KeyChangesOf = moved(key)
End Function

Private Function IssueRowOf(ByVal issue As Variant, ByVal workTypeField As String, ByVal settings As Object, _
                          ByVal moved As Object) As Variant
    ' One row of the sheet Issues (the columns are those of FlowFormulas, table Issues).
    Dim sprint As Variant, key As String
    key = CStr(JGet(issue, "key"))
    sprint = SprintInfo(FieldValue(issue, Setting(settings, "SprintField")))
    IssueRowOf = Array(JGet(issue, "key"), JGet(issue, "id"), JPath(issue, "fields.project.key"), _
                     JPath(issue, "fields.issuetype.name"), WorkTypeOf(issue, workTypeField), _
                     ParseJiraDate(JPath(issue, "fields.created")), JPath(issue, "fields.status.name"), _
                     TextOf(JPath(issue, "fields.parent.key")), TextOf(JPath(issue, "fields.status.statusCategory.key")), _
                     ParseJiraDate(JPath(issue, "fields.resolutiondate")), VersionsOf(issue), _
                     NumberOf(FieldValue(issue, Setting(settings, "StoryPointsField"))), _
                     TextOf(FieldValue(issue, Setting(settings, "PIField"))), _
                     NumberOf(FieldValue(issue, Setting(settings, "BusinessValueField"))), _
                     TextOf(FieldValue(issue, Setting(settings, "TeamField"))), _
                     sprint(0), sprint(1), sprint(2), sprint(3), _
                     DateOf(FieldValue(issue, Setting(settings, "TargetStartField"))), _
                     DateOf(FieldValue(issue, Setting(settings, "TargetEndField"))), _
                     KeyChangesOf(moved, key))
End Function

Private Sub AddBlockLinks(ByVal links As Collection, ByVal issue As Variant, ByVal linkType As String)
    ' One row (blocker, blocked) for every link of the type "blocks" that starts at this issue (outward side only: the
    ' other end of the same link shows it as inward, so each link is counted once).
    Cov "MEX06"
    Dim link As Variant, other As Variant
    If Len(linkType) = 0 Then Exit Sub
    For Each link In AsList(JPath(issue, "fields.issuelinks"))
        If CStr(JPath(link, "type.name") & "") = linkType Then
            other = JPath(link, "outwardIssue.key")
            If Not IsEmpty(other) And Not IsNull(other) Then links.Add Array(JGet(issue, "key"), other)
        End If
    Next link
End Sub

Public Sub ExtractFlowDataTo(ByVal cfg As Object, ByVal jql As String, ByVal workTypeField As String, _
                             ByVal asOf As Date, ByVal wsProjects As Object, ByVal wsWorkflows As Object, _
                             ByVal wsIssues As Object, ByVal wsIntervals As Object, _
                             Optional ByVal clearDerived As Boolean = False, _
                             Optional ByVal settings As Object = Nothing, Optional ByVal wsLinks As Object = Nothing)
    ' clearDerived: after everything was read from Jira, and only then, empty the sheets of an earlier run (a failed
    ' extraction therefore leaves the previous results as they were).
    ' settings: the field ids of the data checks (see FieldSettings); wsLinks: where the "blocks" links go.
    ' Issues: Key | Id | Project | Type | WorkType | CreatedUtc | Status | Parent | StatusCategory | ResolvedUtc | FixVersions |
    '   StoryPoints | PI | BV | TeamValue | SprintState | SprintStart | SprintEnd | SprintCount | TargetStart | TargetEnd | KeyChanges.
    ' StatusIntervals: Key | Workflow | Seq | Status | Category | EnteredUtc | LeftUtc | DurationDays | IsOpen | ChangeId | Check.
    ' IssueLinks: Blocker | Blocked.
    ' The Program follows the Feature workflow, every other project the Team workflow (sheet Projects).
    Cov "MEX01"
    Dim names As Object, issues As Collection, issue As Variant, rows As New Collection, keys As New Collection
    Dim idToKey As Object, raw As Collection, changes As Collection, moved As Object, program As String, project As String
    Dim wf As Variant, info As Object, mine As Collection, ch As Variant, iv As Variant, out As New Collection
    Dim links As New Collection
    Set issues = JiraSearchAll(cfg, jql, names, , ExtractFields(workTypeField, settings), "")     ' only what the metrics use
    Set idToKey = CreateObject("Scripting.Dictionary")
    program = ProgramProject(wsProjects)
    For Each issue In issues
        keys.Add JGet(issue, "key")
        idToKey(CStr(JGet(issue, "id"))) = JGet(issue, "key")
    Next issue
    Set raw = JiraChangelogsAll(cfg, keys)
    Set changes = StatusChangesFromChangelog(raw, idToKey)
    Set moved = KeyChangeCounts(raw, idToKey)
    For Each issue In issues
        rows.Add IssueRowOf(issue, workTypeField, settings, moved)
        AddBlockLinks links, issue, Setting(settings, "BlocksLinkType")
    Next issue
    For Each wf In Array("Feature", "Team")
        Set info = CreateObject("Scripting.Dictionary")
        For Each issue In issues
            project = CStr(JPath(issue, "fields.project.key"))
            If (project = program) = (wf = "Feature") Then
                info(JGet(issue, "key")) = Array(ParseJiraDate(JPath(issue, "fields.created")), JPath(issue, "fields.status.name"))
            End If
        Next issue
        Set mine = New Collection
        For Each ch In changes
            If info.Exists(ch(0)) Then mine.Add ch
        Next ch
        For Each iv In BuildStatusIntervals(info, mine, asOf, StatusCategories(wsWorkflows, CStr(wf)))
            out.Add Array(iv(0), wf, iv(1), iv(2), iv(3), iv(4), iv(5), iv(6), iv(7), iv(8), iv(9))
        Next iv
    Next wf
    If clearDerived Then ClearFlowSheets ThisWorkbook
    WriteTable wsIssues, Array("Key", "Id", "Project", "Type", "WorkType", "CreatedUtc", "Status", "Parent", "StatusCategory", _
                               "ResolvedUtc", "FixVersions", "StoryPoints", "PI", "BV", "TeamValue", "SprintState", _
                               "SprintStart", "SprintEnd", "SprintCount", "TargetStart", "TargetEnd", "KeyChanges"), rows
    WriteTable wsIntervals, Array("Key", "Workflow", "Seq", "Status", "Category", "EnteredUtc", "LeftUtc", _
                                  "DurationDays", "IsOpen", "ChangeId", "Check"), out
    If Not wsLinks Is Nothing Then WriteTable wsLinks, Array("Blocker", "Blocked"), links
End Sub
