Option Explicit
' Reads the configuration of Jira, read-only, and writes proposal sheets for the parameters:
'   JiraIssueTypes  the workflow of each issue type of each project of the Projects sheet;
'   JiraStatuses    the statuses of each issue type, with the Jira category (needs only the right to browse);
'   JiraWorkflows   one row per status of each workflow, same first columns as the sheet Workflows;
'   JiraTransitions one row per allowed move, same first columns as the sheet Transitions;
'   JiraFields      the field ids, for the Config sheet.
' The parameter sheets are never written: the user reviews the proposals, decides what stays human (Order, IsStart, IsEnd,
' Active or Wait, the Type of a move) and copies them. Calls (claude-code-handoff.md section 8c): GET project/search,
' GET project/{key}/statuses, POST workflows/preview, GET field. A 401, 403 or 404 on the preview means the account may
' not read the workflows: the other sheets are still written. The answers stay in this workbook; the log gets counts only.

Private Const KEYS_PER_CALL As Long = 50             ' API limit of project/search
Private Const PROJECTS_PER_PAGE As Long = 50
Private Const TYPES_PER_PREVIEW As Long = 25         ' API limit of workflows/preview
Private Const MAX_PAGES As Long = 100

Public Function ReadJiraConfiguration(ByVal cfg As Object, ByVal keys As Collection, ByVal wb As Object) As String
    ' Writes the five proposal sheets of wb and returns the summary for the status cell.
    Cov "JCF01"
    Dim projects As Collection, project As Variant, n As Long, denied As Long, workflowOf As Object, seen As Object
    Dim typeRows As New Collection, statusRows As New Collection, flowRows As New Collection, moveRows As New Collection
    Dim fieldRows As New Collection, key As String, summary As String
    If keys.Count = 0 Then
        Cov "JCF02"
        Err.Raise vbObjectError + 620, "ReadJiraConfiguration", "Sheet Projects has no project: type the keys of the Jira projects first."
    End If
    LogBegin "projects"
    Set projects = ReadProjects(cfg, keys)
    LogEnd "projects: " & projects.Count & " of " & keys.Count
    Set seen = CreateObject("Scripting.Dictionary")
    For Each project In projects
        n = n + 1
        key = CStr(JGet(project, "key") & "")
        Application.StatusBar = "Reading the Jira configuration: project " & n & " of " & projects.Count
        LogBegin "configuration of project " & n & " of " & projects.Count
        Set workflowOf = CreateObject("Scripting.Dictionary")
        If Not ReadWorkflows(cfg, project, workflowOf, seen, flowRows, moveRows) Then denied = denied + 1
        AddIssueTypeRows project, workflowOf, typeRows
        AddStatusRows key, JiraJson(cfg, "GET", StatusesUrl(cfg("BaseUrl"), key), "", "project statuses"), statusRows
        LogEnd "configuration of project " & n
    Next project
    LogBegin "fields"
    AddFieldRows JiraJson(cfg, "GET", cfg("BaseUrl") & "/rest/api/3/field", "", "fields"), fieldRows
    LogEnd "fields: " & fieldRows.Count
    LogBegin "proposal sheets"
    WriteProposals wb, typeRows, statusRows, flowRows, moveRows, fieldRows
    LogEnd "proposal sheets"
    Application.StatusBar = False
    summary = "Jira configuration read: " & projects.Count & " project(s), " & seen.Count & " workflow(s), " & _
              fieldRows.Count & " field(s). Review the sheets JiraIssueTypes, JiraStatuses, JiraWorkflows, " & _
              "JiraTransitions and JiraFields, then copy what fits into the parameter sheets."
    If projects.Count < keys.Count Then
        Cov "JCF03"
        summary = summary & " " & (keys.Count - projects.Count) & " project key(s) of the Projects sheet not found in Jira."
    End If
    If denied > 0 Then
        Cov "JCF04"
        summary = summary & " The workflows could not be read for " & denied & " project(s): ask for the project " & _
                  "permission View (read-only) workflow. JiraStatuses still shows the statuses of each issue type."
    End If
    ReadJiraConfiguration = summary
End Function

Public Function ProjectSearchUrl(ByVal baseUrl As String, ByVal keys As Collection, ByVal first As Long, _
                                 ByVal startAt As Long) As String
    ' keys(first) to keys(first + 49), with their issue types, one page of 50 projects from startAt.
    Cov "JCF05"
    Dim url As String, i As Long, last As Long
    last = first + KEYS_PER_CALL - 1
    If last > keys.Count Then last = keys.Count
    url = baseUrl & "/rest/api/3/project/search?"
    For i = first To last
        url = url & "keys=" & UrlEncode(CStr(keys(i))) & "&"
    Next i
    ProjectSearchUrl = url & "expand=issueTypes&startAt=" & startAt & "&maxResults=" & PROJECTS_PER_PAGE
End Function

Public Function StatusesUrl(ByVal baseUrl As String, ByVal projectKey As String) As String
    StatusesUrl = baseUrl & "/rest/api/3/project/" & UrlEncode(projectKey) & "/statuses"
End Function

Public Function ReadProjects(ByVal cfg As Object, ByVal keys As Collection) As Collection
    ' Every project of keys that Jira knows, with its issue types (offset paging; 50 keys per call at most).
    Cov "JCF06"
    Dim out As New Collection, first As Long, startAt As Long, page As Object, value As Variant, pages As Long
    Dim values As Collection
    For first = 1 To keys.Count Step KEYS_PER_CALL
        startAt = 0
        Do
            pages = pages + 1
            If pages > MAX_PAGES Then
                Cov "JCF07"
                Err.Raise vbObjectError + 621, "ReadProjects", "Jira project/search did not end after " & MAX_PAGES & " pages."
            End If
            Set page = JiraJson(cfg, "GET", ProjectSearchUrl(cfg("BaseUrl"), keys, first, startAt), "", "projects")
            Set values = AsList(JGet(page, "values"))
            For Each value In values
                out.Add value
            Next value
            If IsTrue(JGet(page, "isLast")) Or values.Count = 0 Then Exit Do
            Cov "JCF08"
            startAt = startAt + values.Count
        Loop
    Next first
    Set ReadProjects = out
End Function

Public Function PreviewBody(ByVal projectId As String, ByVal typeIds As Collection, ByVal first As Long) As String
    ' {"projectId": "...", "issueTypeIds": [typeIds(first) .. typeIds(first + 24)]}
    Cov "JCF09"
    Dim body As Object, ids As New Collection, i As Long, last As Long
    last = first + TYPES_PER_PREVIEW - 1
    If last > typeIds.Count Then last = typeIds.Count
    For i = first To last
        ids.Add CStr(typeIds(i))
    Next i
    Set body = CreateObject("Scripting.Dictionary")
    body("projectId") = projectId
    Set body("issueTypeIds") = ids
    PreviewBody = JsonStringify(body)
End Function

Private Function ReadWorkflows(ByVal cfg As Object, ByVal project As Variant, ByVal workflowOf As Object, ByVal seen As Object, _
                               ByVal flowRows As Collection, ByVal moveRows As Collection) As Boolean
    ' The workflows of the issue types of one project (25 issue types per call). False when the account may not read them.
    Dim typeIds As New Collection, t As Variant, first As Long, resp As Object
    For Each t In AsList(JGet(project, "issueTypes"))
        typeIds.Add CStr(JGet(t, "id") & "")
    Next t
    ReadWorkflows = True
    For first = 1 To typeIds.Count Step TYPES_PER_PREVIEW
        Set resp = HttpRequest(cfg, "POST", cfg("BaseUrl") & "/rest/api/3/workflows/preview", _
                               PreviewBody(CStr(JGet(project, "id") & ""), typeIds, first))
        If IsNotAllowed(resp("Status")) Then Cov "JCF10": ReadWorkflows = False: Exit Function
        CheckStatus resp, "workflows"
        RequireJson resp, "workflows"
        AddWorkflowRows JsonParse(resp("Text")), workflowOf, seen, flowRows, moveRows
    Next first
End Function

Public Function IsNotAllowed(ByVal status As Long) As Boolean
    ' The preview answers 401 (not 403) when the account lacks the right; 404 when it sees no workflow of the project.
    ' The projects call worked with the same account just before, so this is a missing right, not a wrong token.
    IsNotAllowed = (status = 401 Or status = 403 Or status = 404)
End Function

Public Sub AddWorkflowRows(ByVal answer As Object, ByVal workflowOf As Object, ByVal seen As Object, _
                           ByVal flowRows As Collection, ByVal moveRows As Collection)
    ' One preview answer: issue type id -> workflow name (queryContext) and, for a workflow not met yet (by its id, else
    ' its name), its rows. Two project-scoped workflows with the same name are both written: the clash stays visible.
    ' A statusReference is valid only inside its answer: it is turned into a status name here.
    Cov "JCF11"
    Dim names As Object, cats As Object, s As Variant, wf As Variant, ctx As Variant, t As Variant, name As String, id As String
    Set names = CreateObject("Scripting.Dictionary")
    Set cats = CreateObject("Scripting.Dictionary")
    For Each s In AsList(JGet(answer, "statuses"))
        names(CStr(JGet(s, "statusReference") & "")) = CStr(JGet(s, "name") & "")
        cats(CStr(JGet(s, "statusReference") & "")) = CStr(JGet(s, "statusCategory") & "")
    Next s
    For Each wf In AsList(JGet(answer, "workflows"))
        name = CStr(JGet(wf, "name") & "")
        id = CStr(JGet(wf, "id") & "")
        If Len(id) = 0 Then id = name
        For Each ctx In AsList(JGet(wf, "queryContext"))
            For Each t In AsList(JGet(ctx, "issueTypes"))
                workflowOf(CStr(t)) = name
            Next t
        Next ctx
        If seen.Exists(id) Then
            Cov "JCF12"                                     ' a workflow shared by several projects is written once
        Else
            seen(id) = True
            AddOneWorkflow wf, name, names, cats, flowRows, moveRows
        End If
    Next wf
End Sub

Private Sub AddOneWorkflow(ByVal wf As Variant, ByVal name As String, ByVal names As Object, ByVal cats As Object, _
                           ByVal flowRows As Collection, ByVal moveRows As Collection)
    ' JiraWorkflows: Workflow | Order (1 for the status of creation) | Status | Category (proposed) | IsStart | IsEnd |
    ' IsFree (1 when a GLOBAL transition leads to it: reachable from any status) | JiraCategory.
    ' JiraTransitions: one row per link of a DIRECTED transition: Workflow | From | To | Type (human) | JiraTransition.
    Dim free As Object, initial As String, t As Variant, link As Variant, s As Variant, ref As String, order As Variant
    Set free = CreateObject("Scripting.Dictionary")
    For Each t In AsList(JGet(wf, "transitions"))
        ref = CStr(JGet(t, "toStatusReference") & "")
        Select Case UCase$(CStr(JGet(t, "type") & ""))
            Case "GLOBAL"
                Cov "JCF13": free(ref) = True
            Case "INITIAL"
                Cov "JCF14": initial = ref
            Case "DIRECTED"
                For Each link In AsList(JGet(t, "links"))
                    Cov "JCF15"
                    moveRows.Add Array(name, NameOf(names, CStr(JGet(link, "fromStatusReference") & "")), NameOf(names, ref), _
                                       "", CStr(JGet(t, "name") & ""))
                Next link
        End Select
    Next t
    For Each s In AsList(JGet(wf, "statuses"))
        ref = CStr(JGet(s, "statusReference") & "")
        order = ""
        If Len(initial) > 0 And ref = initial Then order = 1
        flowRows.Add Array(name, order, NameOf(names, ref), ProposedCategory(ValueOf(cats, ref)), "", "", _
                           IIf(free.Exists(ref), 1, 0), ValueOf(cats, ref))
    Next s
End Sub

Private Function NameOf(ByVal names As Object, ByVal ref As String) As String
    ' The status name of a reference; the reference itself when the answer does not list it (shown, not hidden).
    If names.Exists(ref) Then NameOf = names(ref) Else Cov "JCF16": NameOf = ref
End Function

Private Function ValueOf(ByVal dict As Object, ByVal key As String) As String
    If dict.Exists(key) Then ValueOf = dict(key)
End Function

Public Function ProposedCategory(ByVal jiraCategory As String) As String
    ' Jira has three status categories, the workbook five: to do -> Backlog; done -> Done (Removed is a human choice);
    ' in progress -> empty (Active or Wait is a human choice). Both enums: new / indeterminate / done (statuses call) and
    ' TODO / IN_PROGRESS / DONE (preview).
    Select Case UCase$(TrimWhite(jiraCategory))
        Case "TODO", "NEW": Cov "JCF17": ProposedCategory = "Backlog"
        Case "DONE": Cov "JCF18": ProposedCategory = "Done"
        Case Else: Cov "JCF19": ProposedCategory = ""
    End Select
End Function

Public Sub AddIssueTypeRows(ByVal project As Variant, ByVal workflowOf As Object, ByVal rows As Collection)
    ' JiraIssueTypes: Project | IssueType | Workflow (empty when it could not be read) | Subtask | HierarchyLevel | ProjectStyle
    Cov "JCF20"
    Dim t As Variant, id As String, workflow As String
    For Each t In AsList(JGet(project, "issueTypes"))
        id = CStr(JGet(t, "id") & "")
        workflow = ""
        If workflowOf.Exists(id) Then workflow = workflowOf(id)
        rows.Add Array(CStr(JGet(project, "key") & ""), CStr(JGet(t, "name") & ""), workflow, YesNo(JGet(t, "subtask")), _
                       JGet(t, "hierarchyLevel"), CStr(JGet(project, "style") & ""))
    Next t
End Sub

Public Sub AddStatusRows(ByVal projectKey As String, ByVal answer As Object, ByVal rows As Collection)
    ' JiraStatuses: Project | IssueType | Status | JiraCategory (new / indeterminate / done)
    Cov "JCF21"
    Dim t As Variant, s As Variant
    For Each t In AsList(answer)
        For Each s In AsList(JGet(t, "statuses"))
            rows.Add Array(projectKey, CStr(JGet(t, "name") & ""), CStr(JGet(s, "name") & ""), CStr(JPath(s, "statusCategory.key") & ""))
        Next s
    Next t
End Sub

Public Sub AddFieldRows(ByVal answer As Object, ByVal rows As Collection)
    ' JiraFields: Id | Name | Custom | SchemaType | SchemaDetail (the custom field type, or the system field)
    Cov "JCF22"
    Dim f As Variant, detail As String
    For Each f In AsList(answer)
        detail = CStr(JPath(f, "schema.custom") & "")
        If Len(detail) = 0 Then detail = CStr(JPath(f, "schema.system") & "")
        rows.Add Array(CStr(JGet(f, "id") & ""), CStr(JGet(f, "name") & ""), YesNo(JGet(f, "custom")), _
                       CStr(JPath(f, "schema.type") & ""), detail)
    Next f
End Sub

Private Function YesNo(ByVal value As Variant) As String
    If IsTrue(value) Then YesNo = "Yes" Else YesNo = "No"
End Function

Private Function JiraJson(ByVal cfg As Object, ByVal method As String, ByVal url As String, ByVal body As String, _
                          ByVal what As String) As Object
    Dim resp As Object
    Set resp = HttpRequest(cfg, method, url, body)
    CheckStatus resp, what
    RequireJson resp, what
    Set JiraJson = JsonParse(resp("Text"))
End Function

Private Sub WriteProposals(ByVal wb As Object, ByVal typeRows As Collection, ByVal statusRows As Collection, _
                           ByVal flowRows As Collection, ByVal moveRows As Collection, ByVal fieldRows As Collection)
    ' A missing sheet is added at the end; the sheet shown before stays shown (the status cell is on the Dashboard).
    Dim shown As Object
    On Error Resume Next
    Set shown = wb.ActiveSheet
    On Error GoTo 0
    WriteProposal wb, "JiraIssueTypes", Array("Project", "IssueType", "Workflow", "Subtask", "HierarchyLevel", "ProjectStyle"), typeRows
    WriteProposal wb, "JiraStatuses", Array("Project", "IssueType", "Status", "JiraCategory"), statusRows
    WriteProposal wb, "JiraWorkflows", Array("Workflow", "Order", "Status", "Category", "IsStart", "IsEnd", "IsFree", _
                                            "JiraCategory"), flowRows
    WriteProposal wb, "JiraTransitions", Array("Workflow", "From", "To", "Type", "JiraTransition"), moveRows
    WriteProposal wb, "JiraFields", Array("Id", "Name", "Custom", "SchemaType", "SchemaDetail"), fieldRows
    On Error Resume Next
    If Not shown Is Nothing Then shown.Activate
    Err.Clear
End Sub

Private Sub WriteProposal(ByVal wb As Object, ByVal name As String, ByVal headers As Variant, ByVal rows As Collection)
    Dim ws As Object
    If SheetExists(wb, name) Then
        Set ws = wb.Worksheets(name)
    Else
        Cov "JCF23"
        Set ws = wb.Worksheets.Add(, wb.Worksheets(wb.Worksheets.Count))
        ws.Name = name
    End If
    WriteTable ws, headers, rows
End Sub
