Option Explicit
' Extraction of the flow data: Jira -> the raw sheets Issues and StatusIntervals (docs section 2.1).
' VBA writes raw tables only; every figure (ItemFlow, DailyLoad, Metrics) is an Excel formula that reads them.

Public Sub ExtractFlowData()
    ' Everything comes from the workbook: Config (connection, proxy, Filter, WorkTypeField, AsOf), Projects, Workflows.
    Dim wsConfig As Object
    Set wsConfig = ThisWorkbook.Worksheets("Config")
    ExtractFlowDataTo JiraConfigFromSheet(wsConfig), ExtractionJql(ThisWorkbook.Worksheets("Projects"), wsConfig), _
                      ConfigValue(wsConfig, "WorkTypeField"), AsOfFromConfig(wsConfig), _
                      ThisWorkbook.Worksheets("Projects"), ThisWorkbook.Worksheets("Workflows"), _
                      SheetNamed("Issues"), SheetNamed("StatusIntervals"), True
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

Public Function ExtractFields(ByVal workTypeField As String) As String
    ' The Jira fields the extraction reads: the rest (comments, descriptions, attachments) is not downloaded.
    ExtractFields = "key,project,issuetype,created,status" & IIf(Len(workTypeField) > 0, "," & workTypeField, "")
End Function

Public Sub ExtractFlowDataTo(ByVal cfg As Object, ByVal jql As String, ByVal workTypeField As String, _
                             ByVal asOf As Date, ByVal wsProjects As Object, ByVal wsWorkflows As Object, _
                             ByVal wsIssues As Object, ByVal wsIntervals As Object, _
                             Optional ByVal clearDerived As Boolean = False)
    ' clearDerived: after everything was read from Jira, and only then, empty the sheets of an earlier run (a failed
    ' extraction therefore leaves the previous results as they were).
    ' Issues: Key | Id | Project | Type | WorkType | CreatedUtc | Status.
    ' StatusIntervals: Key | Workflow | Seq | Status | Category | EnteredUtc | LeftUtc | DurationDays | IsOpen | ChangeId | Check.
    ' The Program follows the Feature workflow, every other project the Team workflow (sheet Projects).
    Cov "MEX01"
    Dim names As Object, issues As Collection, issue As Variant, rows As New Collection, keys As New Collection
    Dim idToKey As Object, changes As Collection, program As String, project As String
    Dim wf As Variant, info As Object, mine As Collection, ch As Variant, iv As Variant, out As New Collection
    Set issues = JiraSearchAll(cfg, jql, names, , ExtractFields(workTypeField), "")     ' only what the metrics use
    Set idToKey = CreateObject("Scripting.Dictionary")
    program = ProgramProject(wsProjects)
    For Each issue In issues
        rows.Add Array(JGet(issue, "key"), JGet(issue, "id"), JPath(issue, "fields.project.key"), _
                       JPath(issue, "fields.issuetype.name"), WorkTypeOf(issue, workTypeField), _
                       ParseJiraDate(JPath(issue, "fields.created")), JPath(issue, "fields.status.name"))
        keys.Add JGet(issue, "key")
        idToKey(CStr(JGet(issue, "id"))) = JGet(issue, "key")
    Next issue
    Set changes = StatusChangesFromChangelog(JiraChangelogsAll(cfg, keys), idToKey)
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
    WriteTable wsIssues, Array("Key", "Id", "Project", "Type", "WorkType", "CreatedUtc", "Status"), rows
    WriteTable wsIntervals, Array("Key", "Workflow", "Seq", "Status", "Category", "EnteredUtc", "LeftUtc", _
                                  "DurationDays", "IsOpen", "ChangeId", "Check"), out
End Sub
