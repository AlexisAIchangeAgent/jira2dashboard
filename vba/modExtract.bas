Option Explicit
' Extraction of the flow data: Jira -> the raw sheets Issues, StatusIntervals and IssueLinks (docs section 2.1).
' VBA writes raw tables only; every figure (ItemFlow, DailyLoad, Metrics, DataChecks) is an Excel formula that reads them.

' The extraction is a job of small steps (one page of the search, one chunk of changelogs, a batch of issue rows, one
' workflow of intervals, then the writing of the raw sheets):
' ExtractFlowDataTo runs them one after the other; the button of the Dashboard (modActions) runs one per Application.OnTime
' slice, so that the progress bar moves, Excel stays usable and the user can cancel between two steps.
Private Const XL_CALC_MANUAL As Long = -4135                  ' xlCalculationManual
Private mCfg As Object, mJql As String, mWorkType As String, mAsOf As Date, mClear As Boolean, mSettings As Object
Private mWsProjects As Object, mWsWorkflows As Object, mWsIssues As Object, mWsIntervals As Object, mWsLinks As Object
Private mIssues As Collection, mNames As Object, mToken As String, mPage As Long
Private mKeys As Collection, mIdToKey As Object, mRaw As Collection, mNextChunk As Long
Private mPhase As Long                  ' 0 no job, 1 search, 2 changelogs, 3 issue rows, 4 intervals, 5 write
Private mRows As Collection, mLinks As Collection, mChanges As Collection, mMoved As Object, mIntervals As Collection
Private mIds As Variant, mBlocksType As String, mRowNext As Long, mWfNext As Long
Public gRowBatch As Long                ' issues per step of phase 3 (0: ROW_BATCH); a test sets a small one
Private Const ROW_BATCH As Long = 100   ' about 20 ms per batch on the development machine: a step stays short

Private Sub ExtractParameters(ByRef cfg As Object, ByRef jql As String, ByRef settings As Object)
    Dim wsConfig As Object
    Set wsConfig = ThisWorkbook.Worksheets("Config")
    RequireMandatory wsConfig                        ' first, as in the first steps: an empty mandatory value would fail a
                                                     ' request or give wrong figures
    ValidateBuildInputs ThisWorkbook                 ' before any call to Jira: nothing is read or cleared if the build cannot run
    Set cfg = JiraConfigFromSheet(wsConfig)
    jql = ExtractionJql(ThisWorkbook.Worksheets("Projects"), wsConfig)
    Set settings = FieldSettings(wsConfig)
End Sub

Public Sub ExtractFlowData()
    ' Everything comes from the workbook: Config (connection, proxy, Filter, field ids, AsOf), Projects, Workflows.
    Dim cfg As Object, jql As String, settings As Object
    ExtractParameters cfg, jql, settings
    ExtractFlowDataTo cfg, jql, ConfigValue(ThisWorkbook.Worksheets("Config"), "WorkTypeField"), _
                      AsOfFromConfig(ThisWorkbook.Worksheets("Config")), ThisWorkbook.Worksheets("Projects"), _
                      ThisWorkbook.Worksheets("Workflows"), SheetNamed("Issues"), SheetNamed("StatusIntervals"), True, settings, _
                      SheetNamed("IssueLinks")
End Sub

Public Sub ExtractFlowDataStart()
    ' Starts the same extraction as a job: call ExtractJobStep until it returns True.
    Dim cfg As Object, jql As String, settings As Object
    ExtractParameters cfg, jql, settings
    ExtractJobStart cfg, jql, ConfigValue(ThisWorkbook.Worksheets("Config"), "WorkTypeField"), _
                    AsOfFromConfig(ThisWorkbook.Worksheets("Config")), ThisWorkbook.Worksheets("Projects"), _
                    ThisWorkbook.Worksheets("Workflows"), SheetNamed("Issues"), SheetNamed("StatusIntervals"), True, settings, _
                    SheetNamed("IssueLinks")
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

Private Function WorkTypeOf(ByVal fields As Variant, ByVal fieldId As String) As Variant
    ' The work type (TypeCategorie). A select field comes as an object {"value": "..."}; empty if not set.
    Cov "MEX02"
    Dim v As Variant
    If Len(fieldId) = 0 Then Exit Function
    JSet v, JPath(fields, fieldId)
    If IsObject(v) Then Cov "MEX03": JSet v, JGet(v, "value")
    If Not IsNull(v) Then WorkTypeOf = v
End Function

Private Function FieldValue(ByVal fields As Variant, ByVal fieldId As String) As Variant
    ' The raw value of a field by its id; Empty when no id is configured or the field is missing or null.
    Dim v As Variant
    If Len(fieldId) = 0 Then Exit Function
    JSet v, JPath(fields, fieldId)
    If IsObject(v) Then Set FieldValue = v: Exit Function          ' a team or a sprint list is an object: needs Set
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
    Select Case VarType(v)
        Case vbDouble, vbSingle, vbCurrency, vbLong, vbInteger: TextOf = Trim$(Str$(v))     ' a dot, whatever the language of the PC
        Case vbBoolean: TextOf = IIf(v, "true", "false")
        Case Else: TextOf = CStr(v)
    End Select
End Function

Private Function NumberOf(ByVal v As Variant) As Variant
    ' A number field: Jira sends a JSON number. A text is accepted only as plain digits with a dot (never read with the
    ' decimal separator of the PC); a Boolean is not a number.
    If IsEmpty(v) Or IsObject(v) Or IsNull(v) Then Exit Function
    Select Case VarType(v)
        Case vbBoolean
        Case vbString
            If Len(Trim$(v)) > 0 Then
                If Not Trim$(v) Like "*[!0-9.+-]*" Then NumberOf = Val(v)
            End If
        Case Else
            If IsNumeric(v) Then NumberOf = CDbl(v)
    End Select
End Function

Private Function DateOf(ByVal v As Variant) As Variant
    If IsEmpty(v) Or IsObject(v) Then Exit Function
    DateOf = ParseJiraDate(CStr(v))
End Function

Private Function VersionsOf(ByVal fields As Variant) As Variant
    ' Names of the fix versions, joined; Empty when there is none.
    Dim v As Variant, out As String
    For Each v In AsList(JGet(fields, "fixVersions"))
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

Private Function IssueRowOf(ByVal issue As Variant, ByVal workTypeField As String, ByVal ids As Variant, _
                          ByVal moved As Object) As Variant
    ' One row of the sheet Issues (the columns are those of FlowFormulas, table Issues). `ids`: the field ids of
    ' RowFieldIds. `fields` is read once and every lookup starts from it: an issue has some thirty lookups.
    Dim sprint As Variant, key As Variant, fields As Variant
    JSet fields, JGet(issue, "fields")
    key = JGet(issue, "key")
    sprint = SprintInfo(FieldValue(fields, CStr(ids(0))))
    IssueRowOf = Array(key, JGet(issue, "id"), JPath(fields, "project.key"), _
                     JPath(fields, "issuetype.name"), WorkTypeOf(fields, workTypeField), _
                     ParseJiraDate(JPath(fields, "created")), JPath(fields, "status.name"), _
                     TextOf(JPath(fields, "parent.key")), TextOf(JPath(fields, "status.statusCategory.key")), _
                     ParseJiraDate(JPath(fields, "resolutiondate")), VersionsOf(fields), _
                     NumberOf(FieldValue(fields, CStr(ids(1)))), _
                     TextOf(FieldValue(fields, CStr(ids(2)))), _
                     NumberOf(FieldValue(fields, CStr(ids(3)))), _
                     TextOf(FieldValue(fields, CStr(ids(4)))), _
                     sprint(0), sprint(1), sprint(2), sprint(3), _
                     DateOf(FieldValue(fields, CStr(ids(5)))), _
                     DateOf(FieldValue(fields, CStr(ids(6)))), _
                     KeyChangesOf(moved, CStr(key)))
End Function

Private Function RowFieldIds(ByVal settings As Object) As Variant
    ' The field ids IssueRowOf reads, in its order, fetched once per job (not once per issue).
    RowFieldIds = Array(Setting(settings, "SprintField"), Setting(settings, "StoryPointsField"), Setting(settings, "PIField"), _
                        Setting(settings, "BusinessValueField"), Setting(settings, "TeamField"), _
                        Setting(settings, "TargetStartField"), Setting(settings, "TargetEndField"))
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
    ExtractJobStart cfg, jql, workTypeField, asOf, wsProjects, wsWorkflows, wsIssues, wsIntervals, clearDerived, settings, wsLinks
    Do While Not ExtractJobStep()
    Loop
End Sub

Public Sub ExtractJobStart(ByVal cfg As Object, ByVal jql As String, ByVal workTypeField As String, _
                           ByVal asOf As Date, ByVal wsProjects As Object, ByVal wsWorkflows As Object, _
                           ByVal wsIssues As Object, ByVal wsIntervals As Object, _
                           Optional ByVal clearDerived As Boolean = False, _
                           Optional ByVal settings As Object = Nothing, Optional ByVal wsLinks As Object = Nothing)
    Cov "MEX07"
    Set mCfg = cfg: mJql = jql: mWorkType = workTypeField: mAsOf = asOf: mClear = clearDerived: Set mSettings = settings
    Set mWsProjects = wsProjects: Set mWsWorkflows = wsWorkflows: Set mWsIssues = wsIssues
    Set mWsIntervals = wsIntervals: Set mWsLinks = wsLinks
    Set mIssues = New Collection: Set mNames = CreateObject("Scripting.Dictionary"): mToken = "": mPage = 0
    Set mKeys = Nothing: Set mIdToKey = Nothing: Set mRaw = New Collection: mNextChunk = 1
    Set mRows = Nothing: Set mLinks = Nothing: Set mChanges = Nothing: Set mMoved = Nothing: Set mIntervals = Nothing
    mRowNext = 1: mWfNext = 0
    mPhase = 1
End Sub

Public Function ExtractJobStep() As Boolean
    ' One step of the job; True when the raw sheets are written. A step that fails leaves the job unusable: start a new one.
    ' Nothing is written to the sheets before the last step (phase 5), so a job dropped before it leaves the old results.
    Dim isLast As Boolean, issue As Variant
    Select Case mPhase
        Case 1                                                       ' one page of the search
            LogBegin "search page " & (mPage + 1)
            mToken = JiraSearchPage(mCfg, mJql, mToken, mIssues, mNames, ExtractFields(mWorkType, mSettings), "", isLast)
            mPage = mPage + 1
            LogEnd "search page " & mPage & ": " & mIssues.Count & " issues"
            If Not isLast And mPage >= MAX_SEARCH_PAGES Then Err.Raise vbObjectError + 601, "ExtractJobStep", "Jira search did not end after " & MAX_SEARCH_PAGES & " pages: narrow the extraction with the Filter of the Config sheet."
            If isLast Then
                Set mKeys = New Collection: Set mIdToKey = CreateObject("Scripting.Dictionary")
                For Each issue In mIssues
                    mKeys.Add JGet(issue, "key")
                    mIdToKey(CStr(JGet(issue, "id"))) = JGet(issue, "key")
                Next issue
                mPhase = 2
            End If
        Case 2                                                       ' one chunk of changelogs
            Cov "MEX08"
            LogBegin "changelogs from issue " & mNextChunk & " of " & mKeys.Count
            If mNextChunk <= mKeys.Count Then mNextChunk = JiraChangelogChunk(mCfg, mKeys, mNextChunk, mRaw)
            LogEnd "changelogs: " & mRaw.Count & " items"
            If mNextChunk > mKeys.Count Then mPhase = 3
        Case 3                                                       ' the rows of the sheet Issues, one batch of issues
            Cov "MEX10"
            If mRows Is Nothing Then RowsInit
            RowsBatch
            If mRowNext > mIssues.Count Then mPhase = 4: ReleasePages    ' the rows are built: the parsed pages are not read again
        Case 4                                                       ' the intervals of one workflow (Feature, then Team)
            Cov "MEX11"
            IntervalsStep
            If mWfNext > 1 Then mPhase = 5: Set mChanges = Nothing: Set mMoved = Nothing
        Case 5                                                       ' the only step that writes: clear, then the three sheets
            Cov "MEX12"
            WriteRawSheets
            ExtractJobAbort                                          ' the pages and the changelogs are no longer needed
            ExtractJobStep = True
        Case Else
            ExtractJobStep = True
    End Select
End Function

Public Function ExtractJobFraction() As Double
    ' How far the job is, 0 to 1. The search has no known size: its share (30 %) fills as pages arrive.
    Cov "MEX09"
    Select Case mPhase
        Case 1: ExtractJobFraction = 0.3 * mPage / (mPage + 4)
        Case 2: ExtractJobFraction = 0.3 + 0.6 * (mNextChunk - 1) / IIf(mKeys.Count = 0, 1, mKeys.Count)
        Case 3: ExtractJobFraction = 0.9 + 0.006 * (mRowNext - 1) / IIf(mIssues.Count = 0, 1, mIssues.Count)
        Case 4: ExtractJobFraction = 0.906 + 0.006 * mWfNext
        Case 5: ExtractJobFraction = 0.918
        Case Else: ExtractJobFraction = 1
    End Select
End Function

Public Function ExtractJobLabel() As String
    Select Case mPhase
        Case 1: ExtractJobLabel = "Reading Jira: " & mIssues.Count & " issues (page " & mPage & ")"
        Case 2: ExtractJobLabel = "Reading the changelogs: " & (mNextChunk - 1) & " / " & mKeys.Count & " issues"
        Case 3: ExtractJobLabel = "Preparing the issues: " & (mRowNext - 1) & " / " & mIssues.Count
        Case 4: ExtractJobLabel = "Computing the status intervals..."
        Case 5: ExtractJobLabel = "Writing the raw sheets..."
        Case Else: ExtractJobLabel = "Extraction finished."
    End Select
End Function

Public Sub ExtractJobAbort()
    ' Forgets the job (finished, cancelled or failed): nothing was written to the sheets before the last step. Releases the
    ' parsed pages and changelogs, and the connection (it holds the Authorization header).
    mPhase = 0
    Set mIssues = Nothing: Set mKeys = Nothing: Set mRaw = Nothing: Set mIdToKey = Nothing: Set mNames = Nothing
    Set mRows = Nothing: Set mLinks = Nothing: Set mChanges = Nothing: Set mMoved = Nothing: Set mIntervals = Nothing
    Set mCfg = Nothing: Set mSettings = Nothing
    Set mWsProjects = Nothing: Set mWsWorkflows = Nothing: Set mWsIssues = Nothing: Set mWsIntervals = Nothing: Set mWsLinks = Nothing
End Sub

Private Sub RowsInit()
    ' Start of phase 3: the status changes and the key changes of all the issues (pure computing on the changelogs).
    Set mRows = New Collection: Set mLinks = New Collection: mRowNext = 1
    LogBegin "raw sheets: status changes"
    Set mChanges = StatusChangesFromChangelog(mRaw, mIdToKey)
    Set mMoved = KeyChangeCounts(mRaw, mIdToKey)
    LogEnd "raw sheets: status changes: " & mChanges.Count
    Set mRaw = New Collection                                          ' every changelog item of Jira: the status changes above are all we keep
    mIds = RowFieldIds(mSettings)
    mBlocksType = Setting(mSettings, "BlocksLinkType")
End Sub

Private Sub ReleasePages()
    ' The memory of a long run is what a 32-bit Excel runs out of ("Out of memory", error 7, at the write of the sheets): free what
    ' the next steps do not read. The count of issues is the only thing the label of phase 3 asked of them.
    Set mIssues = New Collection: Set mKeys = Nothing: Set mIdToKey = Nothing: Set mNames = Nothing
End Sub

Private Sub RowsBatch()
    ' The rows of the sheet Issues (and the "blocks" links) for the next batch of issues.
    Dim last As Long, i As Long, batch As Long, issue As Variant
    batch = gRowBatch
    If batch < 1 Then Cov "MEX13": batch = ROW_BATCH
    last = mRowNext + batch - 1
    If last > mIssues.Count Then last = mIssues.Count
    LogBegin "raw sheets: issues " & mRowNext & " to " & last
    For i = mRowNext To last
        JSet issue, mIssues(i)
        mRows.Add IssueRowOf(issue, mWorkType, mIds, mMoved)
        AddBlockLinks mLinks, issue, mBlocksType
    Next i
    LogEnd "raw sheets: issues " & mRowNext & " to " & last
    mRowNext = last + 1
End Sub

Private Sub IntervalsStep()
    ' The intervals of one workflow: the Program follows the Feature workflow, every other project the Team workflow. It
    ' reads the rows of phase 3 (key, project, creation date, status), not the JSON again.
    Dim wf As String, program As String, info As Object, row As Variant, mine As Collection, ch As Variant, iv As Variant
    If mIntervals Is Nothing Then Set mIntervals = New Collection
    wf = IIf(mWfNext = 0, "Feature", "Team")
    program = ProgramProject(mWsProjects)
    LogBegin "raw sheets: intervals " & wf
    Set info = CreateObject("Scripting.Dictionary")
    For Each row In mRows
        If (CStr(row(2) & "") = program) = (wf = "Feature") Then info(row(0)) = Array(row(5), row(6))
    Next row
    Set mine = New Collection
    For Each ch In mChanges
        If info.Exists(ch(0)) Then mine.Add ch
    Next ch
    For Each iv In BuildStatusIntervals(info, mine, mAsOf, StatusCategories(mWsWorkflows, wf))
        mIntervals.Add Array(iv(0), wf, iv(1), iv(2), iv(3), iv(4), iv(5), iv(6), iv(7), iv(8), iv(9))
    Next iv
    LogEnd "raw sheets: intervals " & wf & ": " & mIntervals.Count
    mWfNext = mWfNext + 1
End Sub

Private Sub WriteRawSheets()
    ' The clearing of the old sheets and the writing of the new ones, without flicker and without a recalculation per change.
    Dim calcMode As Long, screen As Boolean, num As Long, desc As String
    calcMode = Application.Calculation
    screen = Application.ScreenUpdating
    On Error GoTo Failed                                              ' from here Failed gives both back
    Application.ScreenUpdating = False
    Application.Calculation = XL_CALC_MANUAL
    WriteRawSheetsCore
    Application.Calculation = calcMode
    Application.ScreenUpdating = screen
    Exit Sub
Failed:
    num = Err.Number: desc = Err.Description
    Application.Calculation = calcMode
    Application.ScreenUpdating = screen
    Err.Raise num, "WriteRawSheets", desc
End Sub

Private Sub WriteRawSheetsCore()
    ' One step of the log per thing done (they never overlap): the clearing, then each table with its number of rows.
    If mClear Then
        LogBegin "raw sheets: clear the old results"
        ClearFlowSheets ThisWorkbook
        LogEnd "raw sheets: clear the old results"
    End If
    LogBegin "raw sheets: write Issues " & mRows.Count
    WriteTable mWsIssues, Array("Key", "Id", "Project", "Type", "WorkType", "CreatedUtc", "Status", "Parent", "StatusCategory", _
                                "ResolvedUtc", "FixVersions", "StoryPoints", "PI", "BV", "TeamValue", "SprintState", _
                                "SprintStart", "SprintEnd", "SprintCount", "TargetStart", "TargetEnd", "KeyChanges"), mRows
    Set mRows = Nothing                                               ' each table is freed once written: the next copy needs the room
    LogEnd "raw sheets: write Issues"
    LogBegin "raw sheets: write StatusIntervals " & mIntervals.Count
    WriteTable mWsIntervals, Array("Key", "Workflow", "Seq", "Status", "Category", "EnteredUtc", "LeftUtc", _
                                   "DurationDays", "IsOpen", "ChangeId", "Check"), mIntervals
    Set mIntervals = Nothing
    LogEnd "raw sheets: write StatusIntervals"
    If Not mWsLinks Is Nothing Then
        LogBegin "raw sheets: write IssueLinks " & mLinks.Count
        WriteTable mWsLinks, Array("Blocker", "Blocked"), mLinks
        LogEnd "raw sheets: write IssueLinks"
    End If
End Sub
