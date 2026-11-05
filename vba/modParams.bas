Attribute VB_Name = "modParams"
Option Explicit
' Parameters live in the workbook, never in code. The sheets Workflows, Transitions and FieldMapping are written by
' the Python generator from params\*.md; this module only reads them. Each reader takes the sheet as an argument
' (testable on a scratch sheet). Every call re-reads the sheet: keep the result (for example the Dictionary of
' StatusCategories) when you need it in a loop.

Private Const SHEET_WORKFLOWS As String = "Workflows"
Private Const SHEET_TRANSITIONS As String = "Transitions"

Public Function ReadTable(ByVal ws As Object) As Collection
    ' Row 1 holds the headers. Returns one Dictionary (header -> cell value) per row whose first cell is not empty.
    Cov "PRM01"
    Dim data As Variant, out As New Collection, r As Long, c As Long, d As Object, nRows As Long, nCols As Long
    Set ReadTable = out
    nRows = ws.Cells(ws.Rows.Count, 1).End(-4162).Row            ' -4162 = xlUp
    nCols = ws.Cells(1, ws.Columns.Count).End(-4159).Column      ' -4159 = xlToLeft
    If nRows < 2 Then Cov "PRM02": Exit Function
    data = ws.Range(ws.Cells(1, 1), ws.Cells(nRows, nCols)).Value
    For r = 2 To nRows
        If Len(CStr(data(r, 1) & "")) > 0 Then
            Set d = CreateObject("Scripting.Dictionary")
            For c = 1 To nCols
                d(CStr(data(1, c))) = data(r, c)
            Next c
            out.Add d
        End If
    Next r
End Function

Public Function StatusCategories(ByVal wsWorkflows As Object, ByVal workflow As String) As Object
    ' status -> category ("Active", "Wait", ...) for one workflow. This is the input of BuildStatusIntervals.
    Cov "PRM03"
    Dim out As Object, row As Variant
    Set out = CreateObject("Scripting.Dictionary")
    For Each row In ReadTable(wsWorkflows)
        If CStr(row("Workflow")) = workflow Then out(CStr(row("Status"))) = CStr(row("Category"))
    Next row
    Set StatusCategories = out
End Function

Public Function FlaggedStatus(ByVal wsWorkflows As Object, ByVal workflow As String, ByVal flagColumn As String) As String
    ' First status of the workflow whose flag column (IsStart, IsEnd or IsFree) is 1. Empty text if none.
    Cov "PRM04"
    Dim row As Variant
    For Each row In ReadTable(wsWorkflows)
        If CStr(row("Workflow")) = workflow Then
            If Val(CStr(row(flagColumn) & "")) = 1 Then Cov "PRM05": FlaggedStatus = CStr(row("Status")): Exit Function
        End If
    Next row
End Function

Public Function StatusHasFlag(ByVal wsWorkflows As Object, ByVal workflow As String, ByVal status As String, _
                              ByVal flagColumn As String) As Boolean
    Cov "PRM06"
    Dim row As Variant
    For Each row In ReadTable(wsWorkflows)
        If CStr(row("Workflow")) = workflow Then
            If CStr(row("Status")) = status Then StatusHasFlag = (Val(CStr(row(flagColumn) & "")) = 1): Exit Function
        End If
    Next row
End Function

Public Function TransitionAllowed(ByVal wsWorkflows As Object, ByVal wsTransitions As Object, ByVal workflow As String, _
                                  ByVal fromStatus As String, ByVal toStatus As String) As Boolean
    ' Allowed if a row From -> To exists, or if the target is a free status (IsFree = 1: reachable from any status).
    Cov "PRM07"
    Dim row As Variant
    For Each row In ReadTable(wsTransitions)
        If CStr(row("Workflow")) = workflow And CStr(row("From")) = fromStatus And CStr(row("To")) = toStatus Then
            Cov "PRM08": TransitionAllowed = True: Exit Function
        End If
    Next row
    If StatusHasFlag(wsWorkflows, workflow, toStatus, "IsFree") Then Cov "PRM09": TransitionAllowed = True
End Function

Public Function LogicalNames(ByVal wsMapping As Object) As Object
    ' extract column -> logical name (sheet FieldMapping). The macros use the logical name, never the extract column.
    Cov "PRM10"
    Dim out As Object, row As Variant
    Set out = CreateObject("Scripting.Dictionary")
    For Each row In ReadTable(wsMapping)
        out(CStr(row("Extract"))) = CStr(row("Logical"))
    Next row
    Set LogicalNames = out
End Function

Public Function ProjectKeys(ByVal wsProjects As Object) As Collection
    ' Keys of the Projects sheet (columns Project | Role | Team), the Program first.
    Cov "PRM18"
    Dim out As New Collection, rows As Collection, row As Variant
    Set rows = ReadTable(wsProjects)
    For Each row In rows
        If StrComp(TrimWhite(CStr(row("Role") & "")), "Program", vbTextCompare) = 0 Then out.Add TrimWhite(CStr(row("Project")))
    Next row
    For Each row In rows
        If StrComp(TrimWhite(CStr(row("Role") & "")), "Program", vbTextCompare) <> 0 Then out.Add TrimWhite(CStr(row("Project")))
    Next row
    Set ProjectKeys = out
End Function

Public Function ProgramProject(ByVal wsProjects As Object) As String
    ' The project whose Role is Program (the train, it holds the Features). Empty text if the sheet has none.
    Cov "PRM19"
    Dim row As Variant
    For Each row In ReadTable(wsProjects)
        If StrComp(TrimWhite(CStr(row("Role") & "")), "Program", vbTextCompare) = 0 Then
            ProgramProject = TrimWhite(CStr(row("Project"))): Exit Function
        End If
    Next row
End Function

Public Function TeamOfProject(ByVal wsProjects As Object, ByVal projectKey As String) As String
    ' Team name of a project. Empty text for the Program and for an unknown project.
    Cov "PRM20"
    Dim row As Variant
    For Each row In ReadTable(wsProjects)
        If TrimWhite(CStr(row("Project"))) = projectKey Then TeamOfProject = TrimWhite(CStr(row("Team") & "")): Exit Function
    Next row
End Function

Public Function ParseMoment(ByVal text As String) As Variant
    ' "2026-11-03", "2026-11-03 18:00" or "2026-11-03T18:00:00" (UTC) -> Date. Empty text if blank or invalid.
    ' Parsed by hand: CDate depends on the Excel language.
    Cov "PRM21"
    Dim s As String, d As Date, sec As Long
    ParseMoment = Empty
    s = TrimWhite(text)
    If Len(s) < 10 Then Exit Function
    If Mid$(s, 5, 1) <> "-" Or Mid$(s, 8, 1) <> "-" Then Exit Function
    On Error GoTo Invalid
    d = DateSerial(CInt(Left$(s, 4)), CInt(Mid$(s, 6, 2)), CInt(Mid$(s, 9, 2)))
    If Month(d) <> CInt(Mid$(s, 6, 2)) Or Day(d) <> CInt(Mid$(s, 9, 2)) Then GoTo Invalid     ' 2026-13-45 does not exist
    If Len(s) >= 16 Then
        If Len(s) >= 19 Then sec = CInt(Mid$(s, 18, 2))
        d = d + TimeSerial(CInt(Mid$(s, 12, 2)), CInt(Mid$(s, 15, 2)), sec)
    End If
    ParseMoment = d
    Exit Function
Invalid:
    Err.Clear
    ParseMoment = Empty
End Function

Public Function AsOfFromConfig(ByVal wsConfig As Object) As Date
    ' The AsOf parameter of the Config sheet (UTC). The metrics need it: open statuses run up to it.
    Cov "PRM22"
    Dim v As Variant, row As Variant
    For Each row In ReadTable(wsConfig)
        If CStr(row("Parameter")) = "AsOf" Then v = row("Value"): Exit For
    Next row
    If VarType(v) = vbDate Or VarType(v) = vbDouble Then
        Cov "PRM24"                                  ' typed as a date: Excel keeps the date, not the text
        If CDbl(v) < 36526 Then                      ' before 2000-01-01: a plain number such as 5 was typed, not a date
            Cov "PRM25"
            Err.Raise vbObjectError + 604, "AsOfFromConfig", "AsOf in the Config sheet is not a date (" & CStr(v) & "): type a UTC date and time such as 2026-11-03 18:00"
        End If
        AsOfFromConfig = CDate(v): Exit Function
    End If
    v = ParseMoment(CStr(v & ""))
    If IsEmpty(v) Then
        Cov "PRM23"
        Err.Raise vbObjectError + 604, "AsOfFromConfig", "Type AsOf in the Config sheet: a UTC date and time such as 2026-11-03 18:00"
    End If
    AsOfFromConfig = v
End Function

Public Function ConfigValue(ByVal wsConfig As Object, ByVal name As String, Optional ByVal defaultValue As String = "") As String
    ' One value of the Config sheet (columns read by header name: Parameter, Value), trimmed like a pasted value.
    ' An empty cell or an unknown parameter gives defaultValue.
    Cov "PRM14"
    Dim row As Variant, v As String
    For Each row In ReadTable(wsConfig)
        If CStr(row("Parameter")) = name Then
            v = TrimWhite(CStr(row("Value") & ""))
            If Len(v) = 0 Then Cov "PRM15": v = defaultValue
            ConfigValue = v
            Exit Function
        End If
    Next row
    Cov "PRM16"
    ConfigValue = defaultValue
End Function

Public Function MissingMandatory(ByVal wsConfig As Object, Optional ByVal onlyNames As String = "") As String
    ' The parameters of the Config sheet whose Mandatory is True and whose Value is empty, as "BaseUrl, AsOf" ("" = none),
    ' in the order of the sheet. onlyNames ("BaseUrl,Email") limits the check to those parameters; "" = all of them.
    ' A sheet without the column Mandatory (a workbook built before it) has nothing to check.
    Cov "PRM26"
    Dim row As Variant, name As String, out As String
    For Each row In ReadTable(wsConfig)
        If Not row.Exists("Mandatory") Then Cov "PRM27": Exit Function
        name = CStr(row("Parameter") & "")
        If IsYes(CStr(row("Mandatory") & "")) And Len(TrimWhite(CStr(row("Value") & ""))) = 0 Then
            If Len(onlyNames) = 0 Or InStr(1, "," & Replace(onlyNames, " ", "") & ",", "," & name & ",", vbTextCompare) > 0 Then
                Cov "PRM28"
                If Len(out) > 0 Then out = out & ", "
                out = out & name
            End If
        End If
    Next row
    MissingMandatory = out
End Function

Public Sub RequireMandatory(ByVal wsConfig As Object, Optional ByVal onlyNames As String = "")
    ' Stops a button before any call to Jira while a mandatory value of the Config sheet is empty (error 608).
    Dim missing As String
    missing = MissingMandatory(wsConfig, onlyNames)
    If Len(missing) > 0 Then
        Cov "PRM29"
        Err.Raise vbObjectError + 608, "RequireMandatory", "Fill these values of the Config sheet first (Mandatory = True): " & missing
    End If
End Sub

Public Function IsYes(ByVal text As String) As Boolean
    ' The sheet is read in a French Excel: Yes / Oui / True / Vrai / 1 mean yes.
    Cov "PRM17"
    Select Case LCase$(TrimWhite(text))
        Case "yes", "oui", "true", "vrai", "1": IsYes = True
    End Select
End Function

Public Function CategoryOf(ByVal workflow As String, ByVal status As String) As String
    ' Category of a status in this workbook's Workflows sheet. A status that is not listed is "Unclassified".
    Cov "PRM11"
    Dim cats As Object
    Set cats = StatusCategories(ThisWorkbook.Worksheets(SHEET_WORKFLOWS), workflow)
    If cats.Exists(status) Then CategoryOf = cats(status) Else Cov "PRM12": CategoryOf = "Unclassified"
End Function

Public Function IsAllowed(ByVal workflow As String, ByVal fromStatus As String, ByVal toStatus As String) As Boolean
    ' Transition check on this workbook's sheets Workflows and Transitions.
    Cov "PRM13"
    IsAllowed = TransitionAllowed(ThisWorkbook.Worksheets(SHEET_WORKFLOWS), ThisWorkbook.Worksheets(SHEET_TRANSITIONS), _
                                  workflow, fromStatus, toStatus)
End Function
