Attribute VB_Name = "modRovoDebug"
Option Explicit
' The button Debug prompt for Rovo (AL, 2026-11-26): Markdown files that let Rovo compare the figures of the workbook with Jira.
' The texts, the JQL, the limits and the columns of MetricItems are the sheet RovoDebugSpec (src/jira2dashboard/rovo_debug.py, written at build).
' This module reads the tables of 2. Compute the metrics (Reconciliation, Metrics, VelocityCheck, the lists of MetricItems), StatusRules, Projects
' and three values of Config (AsOf through AsOfUtc, Filter, FeatureTypes: never a value of the connection), replaces the tokens {name} of the texts,
' packs the iterations (the last MaxIterations that started) into files of at most MaxCharacters, and writes 01.md, 02.md... as UTF-8 without BOM
' (WriteUtf8File: plain VBA, no ADODB.Stream) into the folder rovo-debug next to the workbook. File 1 holds the instructions, the context and the
' controls; every other file starts with a head that sends Rovo to file 1. The log gets counts and sizes only: no key, no path.

Public Const SHEET_ROVO_DEBUG_SPEC As String = "RovoDebugSpec"
Public Const ROVO_DEBUG_FOLDER As String = "rovo-debug"
Private Const COMPUTE_CAPTION As String = "2. Compute the metrics"      ' flow/dashboard_spec.COMPUTE_CAPTION
Private Const PROGRAM_TEAM As String = "ART"                            ' flow/model.ART: the team of the Features of the Program
Private Const FEATURE_KIND As String = "Feature"                        ' the workflow kinds of StatusRules: the Program project, the team projects
Private Const TEAM_KIND As String = "Team"
Private Const SEP As String = ", "                                      ' between keys, names of a JQL list
Private Const ITEM_SEP As String = "; "                                 ' between items that hold ", " ("key: team, reason"): Rovo counts them
Private Const NO_AS_OF As Double = 2958465                              ' 9999-12-31: no AsOf, every iteration counts as started

Public Function RovoDebugFolder(ByVal workbookPath As String, ByVal localAppData As String, ByVal temp As String) As String
    ' rovo-debug next to the workbook; under the folder of the user when the workbook has no folder of its own (never saved, or opened from
    ' SharePoint: an http path), as modLog.LogFolderFrom.
    Cov "RVD01"
    If Len(workbookPath) > 0 And LCase$(Left$(workbookPath, 4)) <> "http" Then RovoDebugFolder = workbookPath & "\" & ROVO_DEBUG_FOLDER: Exit Function
    Cov "RVD02"
    If Len(localAppData) = 0 Then localAppData = temp
    RovoDebugFolder = localAppData & "\jira2dashboard\" & ROVO_DEBUG_FOLDER
End Function

Public Sub RequireRovoDebugTables(ByVal wb As Object)
    ' The tables of 2. Compute the metrics (1. Extract from Jira deletes them): error 633 (J2D-ROVO-004) without them.
    Cov "RVD03"
    Dim name As Variant, ready As Boolean
    For Each name In Array("Reconciliation", "Metrics", "VelocityCheck", "MetricItems")
        If name = "MetricItems" Then ready = SheetExists(wb, CStr(name)) Else ready = TableExists(wb, CStr(name))   ' MetricItems: spilled lists, no table
        If Not ready Then
            Cov "RVD04"
            Err.Raise vbObjectError + 633, "WriteRovoDebug", "The sheet " & SheetName(CStr(name)) & " is not computed: click " & COMPUTE_CAPTION & _
                      ", then Debug prompt for Rovo."
        End If
    Next name
End Sub

Public Function WriteRovoDebug(ByVal wb As Object, ByVal folder As String, ByVal health As String) As String
    ' Writes the files into `folder` (its older NN.md deleted first) and returns "files|largest|listed|total|limit": the number of files, the
    ' characters of the largest, the iterations listed, those of the table Metrics and MaxCharacters (an iteration larger than it has a file of its
    ' own: the status says so). `health`: the fields of modHealth.HealthFields (" | Name: value...").
    Cov "RVD05"
    Dim spec As Object, wsConfig As Object, asOf As Double, filter As String, featureTypes As String, filterJql As String
    Dim programKeys As Collection, allKeys As Collection, teamProjects As Object, projectLines As Collection
    Dim ends As Object, statusLines As Collection, labels As Collection, starts As Collection, finishes As Collection, rowsOf As Object
    Dim metrics As Variant, mh As Object, reasons As Object, done As Object, leftOut As Object, cutoff As Object
    Dim chosen As Collection, sections As New Collection, sizes() As Long, files() As Long, i As Long, k As Variant, n As Long
    Dim summary As String, headUpper As String, nextLine As String, texts() As String, largest As Long, nFiles As Long, maxChars As Long
    RequireRovoDebugTables wb
    Set spec = ReadSpec(wb)
    maxChars = CLng(SpecText(spec, "limit", "MaxCharacters"))
    Set wsConfig = wb.Worksheets(SheetName("Config"))
    asOf = AsOfOf(wb)
    filter = TrimWhite(ConfigValue(wsConfig, "Filter"))
    featureTypes = TrimWhite(ConfigValue(wsConfig, "FeatureTypes"))
    If Len(filter) > 0 Then filterJql = RovoFill(SpecText(spec, "jql", "filter"), Array("{filter}", filter))
    ReadProjects wb, programKeys, allKeys, teamProjects, projectLines
    ReadRules wb, ends, statusLines
    metrics = TableData(wb, "Metrics", mh)
    ReadIterations metrics, mh, labels, starts, finishes, rowsOf
    Set reasons = KeyedColumn(wb, "VelocityCheck", "Reason")
    ReadLists wb, spec, done, leftOut, cutoff
    Set chosen = RovoLastStarted(starts, asOf, CLng(SpecText(spec, "limit", "MaxIterations")))
    For Each k In chosen
        sections.Add IterationText(spec, CStr(labels(k)), CDbl(starts(k)), CDbl(finishes(k)), rowsOf(CStr(labels(k))), metrics, mh, reasons, done, _
                                   leftOut, cutoff, programKeys, featureTypes, teamProjects, ends, filterJql)
    Next k
    ' file 1: the head and the summary, sized with "file 99" for each iteration of the index; then the sections in files of at most maxChars
    headUpper = RovoFill(SpecText(spec, "head", ""), Array("{file}", "99", "{files}", "99"))
    nextLine = SpecText(spec, "next", "")
    ReDim files(0 To chosen.Count)
    For i = 1 To chosen.Count: files(i) = 99: Next i
    summary = SummaryText(spec, wb, asOf, filter, featureTypes, filterJql, programKeys, allKeys, projectLines, statusLines, labels, starts, finishes, _
                          chosen, files, health)
    ReDim sizes(0 To sections.Count)
    For i = 1 To sections.Count: sizes(i) = Len(sections(i)) + 1: Next i
    files = RovoPack(sizes, sections.Count, Len(headUpper) + 1 + Len(summary), Len(headUpper) + 1 + Len(nextLine), maxChars)
    nFiles = 1
    If sections.Count > 0 Then nFiles = files(sections.Count)
    summary = SummaryText(spec, wb, asOf, filter, featureTypes, filterJql, programKeys, allKeys, projectLines, statusLines, labels, starts, finishes, _
                          chosen, files, health)
    ReDim texts(1 To nFiles)
    For n = 1 To nFiles
        texts(n) = RovoFill(SpecText(spec, "head", ""), Array("{file}", CStr(n), "{files}", CStr(nFiles)))
        If n = 1 Then texts(n) = texts(n) & vbLf & summary Else texts(n) = texts(n) & vbLf & nextLine
        For i = 1 To sections.Count
            If files(i) = n Then texts(n) = texts(n) & vbLf & sections(i)
        Next i
        If Len(texts(n)) > largest Then largest = Len(texts(n))
    Next n
    MakeFolder folder
    On Error Resume Next
    For n = 1 To 99                                                     ' the files of an earlier click only (a wildcard ??.md also takes x.md, ok.md)
        Kill folder & "\" & Format$(n, "00") & ".md"
    Next n
    Err.Clear
    On Error GoTo 0
    For n = 1 To nFiles
        WriteUtf8File folder & "\" & Format$(n, "00") & ".md", texts(n)
    Next n
    WriteRovoDebug = nFiles & "|" & largest & "|" & chosen.Count & "|" & labels.Count & "|" & maxChars
End Function

' ---- the text ---------------------------------------------------------------------------------------------------------------------------

Public Function RovoFill(ByVal text As String, ByVal pairs As Variant) As String
    ' Each token of `pairs` (token, value, token, value...) replaced in `text`, in that order.
    Dim i As Long
    For i = LBound(pairs) To UBound(pairs) - 1 Step 2
        text = Replace(text, CStr(pairs(i)), CStr(pairs(i + 1)))
    Next i
    RovoFill = text
End Function

Public Function RovoPack(ByRef sizes() As Long, ByVal count As Long, ByVal firstSize As Long, ByVal headSize As Long, ByVal limit As Long) As Long()
    ' The file of each section 1..count (characters in sizes(i)): file 1 starts with firstSize characters, every other file with headSize; a
    ' section goes into the current file while it stays within `limit`, else it starts the next file. A section larger than `limit` alone gets a
    ' file of its own (the status gives the size of the largest file).
    Cov "RVD06"
    Dim out() As Long, i As Long, file As Long, used As Long, inFile As Long
    ReDim out(0 To count)
    file = 1: used = firstSize
    For i = 1 To count
        If used + sizes(i) > limit And (inFile > 0 Or file = 1) Then
            Cov "RVD07"
            file = file + 1: used = headSize: inFile = 0
        End If
        used = used + sizes(i): inFile = inFile + 1
        out(i) = file
    Next i
    RovoPack = out
End Function

Public Function JqlQuoted(ByVal items As Collection) As String
    ' "a", "b": the names of a JQL list, a backslash and a quote escaped.
    Dim parts As New Collection, item As Variant
    For Each item In items
        parts.Add """" & Replace(Replace(CStr(item), "\", "\\"), """", "\""") & """"
    Next item
    JqlQuoted = JoinWith(parts, SEP)
End Function

Public Function RovoNumber(ByVal v As Variant) As String
    ' A cell as text for the file: a number with a point and no more than 2 decimals (never "1,5" on a French PC), an error "not computable".
    Dim s As String
    If IsError(v) Then Cov "RVD08": RovoNumber = "not computable": Exit Function
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then RovoNumber = v: Exit Function
    If IsNumeric(v) Then
        s = Trim$(Str$(Round(CDbl(v), 2)))
        If Left$(s, 1) = "." Then s = "0" & s
        If Left$(s, 2) = "-." Then s = "-0" & Mid$(s, 2)
        RovoNumber = s
        Exit Function
    End If
    RovoNumber = CStr(v)
End Function

Private Function SummaryText(ByVal spec As Object, ByVal wb As Object, ByVal asOf As Double, ByVal filter As String, ByVal featureTypes As String, _
                             ByVal filterJql As String, ByVal programKeys As Collection, ByVal allKeys As Collection, ByVal projectLines As Collection, _
                             ByVal statusLines As Collection, ByVal labels As Collection, ByVal starts As Collection, ByVal finishes As Collection, _
                             ByVal chosen As Collection, ByRef files() As Long, ByVal health As String) As String
    ' File 1 after its head: how to answer, the causes, the context and the controls (the part "summary" of the spec).
    Dim index As New Collection, k As Variant, i As Long, none As String, first As String, last As String, features As String, asOfText As String
    none = SpecText(spec, "word", "none")
    For Each k In chosen
        i = i + 1
        index.Add "  - " & labels(k) & ": from " & Moment(starts(k)) & " to " & Moment(finishes(k)) & ", file " & files(i)
    Next k
    first = none: last = none
    If chosen.Count > 0 Then first = Moment(starts(chosen(1))): last = Moment(finishes(chosen(chosen.Count)))
    features = none
    If programKeys.Count > 0 Then features = ScopeArt(spec, programKeys, featureTypes)
    If Len(features) > 0 And features <> none Then features = "`" & features & filterJql & "`" Else features = none
    asOfText = none
    If asOf < NO_AS_OF Then asOfText = Moment(asOf)
    SummaryText = FillLines(SpecLines(spec, "summary"), Array( _
        "{version}", WorkbookVersion(wb), "{as_of}", asOfText, "{filter}", IIf(Len(filter) > 0, filter, none), _
        "{feature_types}", IIf(Len(featureTypes) > 0, featureTypes, none), "{projects}", Listed(projectLines, "  - " & none), _
        "{statuses}", Listed(statusLines, "  - " & none), "{iterations_total}", CStr(labels.Count), "{iterations_listed}", CStr(chosen.Count), _
        "{window_start}", first, "{window_end}", last, "{iteration_index}", Listed(index, "  - " & none), _
        "{jql_extraction}", "`" & RovoFill(SpecText(spec, "jql", "scope_team"), Array("{projects}", JqlQuoted(allKeys))) & filterJql & "`", _
        "{jql_features}", features, "{health}", HealthLines(health, none), "{parameters}", ParameterLines(wb, none), _
        "{controls}", ControlLines(wb, spec)))
End Function

Private Function IterationText(ByVal spec As Object, ByVal label As String, ByVal start As Double, ByVal finish As Double, ByVal teamRows As Collection, _
                               ByVal metrics As Variant, ByVal mh As Object, ByVal reasons As Object, ByVal done As Object, ByVal leftOut As Object, _
                               ByVal cutoff As Object, ByVal programKeys As Collection, ByVal featureTypes As String, ByVal teamProjects As Object, _
                               ByVal ends As Object, ByVal filterJql As String) As String
    ' One iteration: its head, the tickets left out of the metrics and those near a boundary, then one block per team of the table Metrics.
    Dim blocks As New Collection, r As Variant, team As String, kind As String, keys As Collection, why As Collection, item As Variant
    Dim scope As String, endList As String, jqlA As String, jqlB As String, reason As String, none As String, startText As String, lastText As String
    none = SpecText(spec, "word", "none")
    startText = Moment(start): lastText = Moment(finish - 1 / 1440)            ' DURING includes both ends: the last minute of the iteration
    For Each r In teamRows
        team = CStr(metrics(r, mh("Team")) & "")
        Set keys = New Collection: Set why = New Collection
        If done.Exists(label & "|" & team) Then
            For Each item In done(label & "|" & team)
                keys.Add item(0)
                If CStr(item(1)) <> "1" Then why.Add item(0) & ": " & item(2)       ' joined with "; " (ITEM_SEP)
            Next item
        End If
        If team = PROGRAM_TEAM Then
            kind = FEATURE_KIND
            scope = ""
            If programKeys.Count > 0 Then scope = ScopeArt(spec, programKeys, featureTypes)
        Else
            kind = TEAM_KIND
            scope = ""
            If teamProjects.Exists(team) Then scope = RovoFill(SpecText(spec, "jql", "scope_team"), Array("{projects}", JqlQuoted(teamProjects(team))))
        End If
        endList = ""
        If ends.Exists(kind) Then endList = JqlQuoted(ends(kind))
        If Len(endList) = 0 Or Len(scope) = 0 Then
            Cov "RVD09"                                                     ' no JQL rather than a wrong one: "in ()" is no JQL
            jqlA = SpecText(spec, "word", IIf(Len(scope) = 0, "no_scope", "no_end")): jqlB = jqlA
        Else
            jqlA = "`" & RovoFill(SpecText(spec, "jql", "missing"), Array("{scope}", scope, "{filter_jql}", filterJql, "{end_statuses}", endList, _
                   "{start}", startText, "{last}", lastText, "{not_keys}", _
                   IIf(keys.Count > 0, RovoFill(SpecText(spec, "jql", "not_keys"), Array("{keys}", JoinWith(keys, SEP))), ""))) & "`"
            If keys.Count = 0 Then
                Cov "RVD10"
                jqlB = SpecText(spec, "word", "no_keys")
            Else
                Cov "RVD11"
                jqlB = "`" & RovoFill(SpecText(spec, "jql", "extra"), Array("{keys}", JoinWith(keys, SEP), "{end_statuses}", endList, _
                       "{start}", startText, "{last}", lastText)) & "`"
            End If
        End If
        reason = ""
        If reasons.Exists(label & "|" & team) Then reason = reasons(label & "|" & team)
        If Len(reason) > 0 Then reason = " (" & reason & ")"
        blocks.Add FillLines(SpecLines(spec, "team"), Array("{iteration}", label, "{team}", team, "{velocity}", RovoNumber(metrics(r, mh("Velocity"))), _
            "{velocity_reason}", reason, "{flow_time_count}", RovoNumber(metrics(r, mh("FlowTimeCount"))), "{done_count}", CStr(keys.Count), _
            "{done_keys}", IIf(keys.Count > 0, JoinWith(keys, SEP), none), "{not_in_flow_time}", IIf(why.Count > 0, JoinWith(why, ITEM_SEP), none), _
            "{jql_missing}", jqlA, "{jql_extra}", jqlB))
    Next r
    IterationText = FillLines(SpecLines(spec, "iteration"), Array("{iteration}", label, "{start}", startText, "{end}", Moment(finish), _
        "{left_out}", KeyedList(leftOut, label, none), "{cutoff}", KeyedList(cutoff, label, none), "{teams}", JoinWith(blocks, vbLf)))
End Function

Private Function ScopeArt(ByVal spec As Object, ByVal programKeys As Collection, ByVal featureTypes As String) As String
    ' The Program projects and the types of FeatureTypes (a comma list); "" without a type: "issuetype in ()" is no JQL.
    Dim types As New Collection, t As Variant
    For Each t In Split(featureTypes, ",")
        If Len(TrimWhite(CStr(t))) > 0 Then types.Add TrimWhite(CStr(t))
    Next t
    If types.Count = 0 Then Exit Function
    ScopeArt = RovoFill(SpecText(spec, "jql", "scope_art"), Array("{program}", JqlQuoted(programKeys), "{feature_types}", JqlQuoted(types)))
End Function

Private Function ControlLines(ByVal wb As Object, ByVal spec As Object) As String
    ' One line per control of the table Reconciliation: its result, the title of the spec (Rovo knows no DCnn), Expected and Found, and what to do
    ' when it is not OK.
    Dim texts As Object, part As Variant, item As Variant, data As Variant, heads As Object, r As Long, lines As New Collection
    Dim name As String, result As String, title As String, hint As String
    Set texts = CreateObject("Scripting.Dictionary")
    For Each part In Array("jira", "partial", "workbook")
        For Each item In SpecLines(spec, CStr(part))
            texts(item(0)) = Array(item(1), item(2))
        Next item
    Next part
    data = TableData(wb, "Reconciliation", heads)
    If IsEmpty(data) Then Exit Function
    For r = 1 To UBound(data, 1)
        name = CStr(data(r, heads("Control")) & "")
        If Len(name) > 0 Then
            result = "?"
            If Not IsError(data(r, heads("Result"))) Then result = CStr(data(r, heads("Result")) & "")
            title = name: hint = ""
            If texts.Exists(name) Then title = texts(name)(0): hint = texts(name)(1)
            lines.Add "- " & result & ": " & title & ". Expected " & RovoNumber(data(r, heads("Expected"))) & ", found " & _
                      RovoNumber(data(r, heads("Found"))) & "." & IIf(result <> "OK" And Len(hint) > 0, " " & hint, "")
        End If
    Next r
    ControlLines = JoinWith(lines, vbLf)
End Function

Private Function ParameterLines(ByVal wb As Object, ByVal none As String) As String
    ' The parameters in force under the table Reconciliation (modFlowBuild.BuildReconciliation: a blank row, a title, then one row each), as shown.
    Dim ws As Object, lo As Object, r As Long, lines As New Collection
    Set ws = wb.Worksheets(SheetName("Reconciliation"))
    Set lo = ws.ListObjects("Reconciliation")
    r = lo.Range.Row + lo.Range.Rows.Count + 2
    Do While Len(CStr(ws.Cells(r, 1).Value & "")) > 0
        lines.Add "  - " & ws.Cells(r, 1).Value & ": " & ws.Cells(r, 2).Text
        r = r + 1
    Loop
    ParameterLines = Listed(lines, "  - " & none)
End Function

Private Function HealthLines(ByVal health As String, ByVal none As String) As String
    ' " | Name: value | Name: value" -> one line per field.
    Dim lines As New Collection, part As Variant
    For Each part In Split(health, " | ")
        If Len(TrimWhite(CStr(part))) > 0 Then lines.Add "  - " & TrimWhite(CStr(part))
    Next part
    HealthLines = Listed(lines, "  - " & none)
End Function

' ---- the data of the workbook -------------------------------------------------------------------------------------------------------------

Private Sub ReadProjects(ByVal wb As Object, ByRef programKeys As Collection, ByRef allKeys As Collection, ByRef teamProjects As Object, _
                         ByRef lines As Collection)
    ' The sheet Projects: the Program keys (the team ART), every key, the keys of each team, one line per project. A team is named as
    ' modFlowBuild.TeamNames names it (the names of Metrics): the Team cell, else the key AS TYPED, for the role Team only; the keys of the JQL
    ' are in capitals (Jira's keys).
    Dim row As Variant, key As String, role As String, team As String
    Set programKeys = New Collection: Set allKeys = New Collection: Set lines = New Collection
    Set teamProjects = CreateObject("Scripting.Dictionary")
    For Each row In ReadTable(wb.Worksheets(SheetName("Projects")))
        key = TrimWhite(CStr(row("Project") & "")): role = TrimWhite(CStr(row("Role") & "")): team = TrimWhite(CStr(row("Team") & ""))
        If Len(key) > 0 Then
            allKeys.Add UCase$(key)
            If StrComp(role, "Program", vbTextCompare) = 0 Then
                programKeys.Add UCase$(key)
                lines.Add "  - " & UCase$(key) & ": " & role & ", team " & PROGRAM_TEAM
            ElseIf StrComp(role, "Team", vbTextCompare) = 0 Then
                If Len(team) = 0 Then team = key
                If Not teamProjects.Exists(team) Then teamProjects.Add team, New Collection
                teamProjects(team).Add UCase$(key)
                lines.Add "  - " & UCase$(key) & ": " & role & ", team " & team
            Else
                lines.Add "  - " & UCase$(key) & ": " & role
            End If
        End If
    Next row
End Sub

Private Sub ReadRules(ByVal wb As Object, ByRef ends As Object, ByRef lines As Collection)
    ' The sheet StatusRules: the end statuses of each workflow kind (the JQL), and one line per kind: its start and end statuses, then the statuses
    ' of each category in the order met.
    Dim row As Variant, kind As String, status As String, category As String, kinds As Object, k As Variant, c As Variant, text As String
    Set ends = CreateObject("Scripting.Dictionary"): Set kinds = CreateObject("Scripting.Dictionary"): Set lines = New Collection
    If Not SheetExists(wb, "StatusRules") Then Exit Sub
    For Each row In ReadTable(wb.Worksheets(SheetName("StatusRules")))
        kind = TrimWhite(CStr(row("Workflow") & "")): status = CStr(row("Status") & ""): category = TrimWhite(CStr(row("Category") & ""))
        If Len(status) > 0 Then
            If Not kinds.Exists(kind) Then
                kinds.Add kind, CreateObject("Scripting.Dictionary")
                kinds(kind).Add "start", New Collection
                kinds(kind).Add "end", New Collection
            End If
            If IsOne(row("IsStart")) Then kinds(kind)("start").Add status
            If IsOne(row("IsEnd")) Then
                kinds(kind)("end").Add status
                If Not ends.Exists(kind) Then ends.Add kind, New Collection
                ends(kind).Add status
            End If
            If Len(category) > 0 Then
                If Not kinds(kind).Exists(category) Then kinds(kind).Add category, New Collection
                kinds(kind)(category).Add status
            End If
        End If
    Next row
    For Each k In kinds.Keys
        text = "  - " & k & ":"
        For Each c In kinds(k).Keys
            text = text & " " & c & ": " & IIf(kinds(k)(c).Count > 0, JoinWith(kinds(k)(c), SEP), "none") & ";"
        Next c
        lines.Add Left$(text, Len(text) - 1)
    Next k
End Sub

Private Sub ReadIterations(ByVal metrics As Variant, ByVal mh As Object, ByRef labels As Collection, ByRef starts As Collection, _
                           ByRef finishes As Collection, ByRef rowsOf As Object)
    ' The iterations of the table Metrics, each once in the order of its rows, and the rows of each (one per team).
    Dim r As Long, label As String, s As Double, e As Double
    Set labels = New Collection: Set starts = New Collection: Set finishes = New Collection
    Set rowsOf = CreateObject("Scripting.Dictionary")
    If IsEmpty(metrics) Then Exit Sub
    For r = 1 To UBound(metrics, 1)
        label = CStr(metrics(r, mh("Iteration-label")) & "")
        If Len(label) > 0 And IsSerial(metrics(r, mh("IterationStart")), s) And IsSerial(metrics(r, mh("IterationEnd")), e) Then
            If Not rowsOf.Exists(label) Then
                rowsOf.Add label, New Collection
                labels.Add label: starts.Add s: finishes.Add e
            End If
            rowsOf(label).Add r
        End If
    Next r
End Sub

Public Function RovoLastStarted(ByVal starts As Collection, ByVal asOf As Double, ByVal most As Long) As Collection
    ' The iterations that started before AsOf, by Start, the last `most` of them: their numbers in `starts`.
    Dim order() As Long, n As Long, i As Long, j As Long, t As Long, out As New Collection
    ReDim order(0 To starts.Count)
    For i = 1 To starts.Count
        If starts(i) < asOf Then n = n + 1: order(n) = i
    Next i
    For i = 2 To n                                                      ' insertion sort by Start: stable, a few dozen iterations
        t = order(i): j = i - 1
        Do While j >= 1
            If starts(order(j)) <= starts(t) Then Exit Do
            order(j + 1) = order(j): j = j - 1
        Loop
        order(j + 1) = t
    Next i
    If n > most Then Cov "RVD12"
    For i = IIf(n > most, n - most + 1, 1) To n
        out.Add order(i)
    Next i
    Set RovoLastStarted = out
End Function

Private Sub ReadLists(ByVal wb As Object, ByVal spec As Object, ByRef done As Object, ByRef leftOut As Object, ByRef cutoff As Object)
    ' The lists of MetricItems at the columns of the spec (formulas.metricitems_column): done "iteration|team" -> Array(key, counted, why); left
    ' out and cut-off "iteration" -> lines "key: team, ...".
    Dim ws As Object, keyCol As Long, last As Long, keys As Variant, its As Variant, teams As Variant, a As Variant, b As Variant, r As Long
    Dim tag As String
    Set done = CreateObject("Scripting.Dictionary"): Set leftOut = CreateObject("Scripting.Dictionary"): Set cutoff = CreateObject("Scripting.Dictionary")
    Set ws = wb.Worksheets(SheetName("MetricItems"))
    keyCol = Col(spec, "Done.Key"): last = ws.Cells(ws.Rows.Count, keyCol).End(-4162).Row         ' -4162 = xlUp
    keys = ListColumn(ws, keyCol, last): its = ListColumn(ws, Col(spec, "Done.Iteration"), last): teams = ListColumn(ws, Col(spec, "Done.Team"), last)
    a = ListColumn(ws, Col(spec, "Done.FlowTimeCounted"), last): b = ListColumn(ws, Col(spec, "Done.WhyNotInFlowTime"), last)
    If Not IsEmpty(keys) Then
        For r = 1 To UBound(keys, 1)
            If Len(CellText(keys(r, 1))) > 0 Then
                tag = CellText(its(r, 1)) & "|" & CellText(teams(r, 1))
                If Not done.Exists(tag) Then done.Add tag, New Collection
                done(tag).Add Array(CellText(keys(r, 1)), RovoNumber(a(r, 1)), CellText(b(r, 1)))
            End If
        Next r
    End If
    keyCol = Col(spec, "LeftOut.Key"): last = ws.Cells(ws.Rows.Count, keyCol).End(-4162).Row
    keys = ListColumn(ws, keyCol, last): its = ListColumn(ws, Col(spec, "LeftOut.Iteration"), last): teams = ListColumn(ws, Col(spec, "LeftOut.Team"), last)
    a = ListColumn(ws, Col(spec, "LeftOut.Account"), last)
    AddKeyed leftOut, keys, its, teams, a, ""
    keyCol = Col(spec, "Cutoff.Key"): last = ws.Cells(ws.Rows.Count, keyCol).End(-4162).Row
    keys = ListColumn(ws, keyCol, last): its = ListColumn(ws, Col(spec, "Cutoff.Iteration"), last): teams = ListColumn(ws, Col(spec, "Cutoff.Team"), last)
    a = ListColumn(ws, Col(spec, "Cutoff.Hours"), last)
    AddKeyed cutoff, keys, its, teams, a, " h"
End Sub

Private Sub AddKeyed(ByVal into As Object, ByVal keys As Variant, ByVal its As Variant, ByVal teams As Variant, ByVal extra As Variant, ByVal unit As String)
    ' "iteration" -> "key: team, <extra><unit>" for each row of a list whose key and iteration are not empty.
    Dim r As Long, it As String
    If IsEmpty(keys) Then Exit Sub
    For r = 1 To UBound(keys, 1)
        it = CellText(its(r, 1))
        If Len(CellText(keys(r, 1))) > 0 And Len(it) > 0 Then
            If Not into.Exists(it) Then into.Add it, New Collection
            into(it).Add CellText(keys(r, 1)) & ": " & CellText(teams(r, 1)) & ", " & RovoNumber(extra(r, 1)) & unit
        End If
    Next r
End Sub

Private Function KeyedColumn(ByVal wb As Object, ByVal table As String, ByVal column As String) As Object
    ' "iteration|team" -> the text of `column` of a table keyed like Metrics.
    Dim data As Variant, heads As Object, r As Long, out As Object, v As Variant
    Set out = CreateObject("Scripting.Dictionary")
    data = TableData(wb, table, heads)
    If Not IsEmpty(data) Then
        For r = 1 To UBound(data, 1)
            v = data(r, heads(column))
            If Not IsError(v) Then out(CStr(data(r, heads("Iteration-label")) & "") & "|" & CStr(data(r, heads("Team")) & "")) = CStr(v & "")
        Next r
    End If
    Set KeyedColumn = out
End Function

Private Function TableData(ByVal wb As Object, ByVal name As String, ByRef heads As Object) As Variant
    ' The body of the table `name` (on the sheet of the same logical name) as a 2-D array, and its column numbers by header; Empty without a row.
    Dim lo As Object, c As Long
    Set heads = CreateObject("Scripting.Dictionary")
    Set lo = wb.Worksheets(SheetName(name)).ListObjects(name)
    For c = 1 To lo.ListColumns.Count
        heads(lo.ListColumns(c).Name) = c
    Next c
    If lo.DataBodyRange Is Nothing Then Exit Function
    TableData = lo.DataBodyRange.Value
End Function

Private Function TableExists(ByVal wb As Object, ByVal name As String) As Boolean
    Dim lo As Object
    If Not SheetExists(wb, name) Then Exit Function
    On Error Resume Next
    Set lo = wb.Worksheets(SheetName(name)).ListObjects(name)
    On Error GoTo 0
    TableExists = Not lo Is Nothing
End Function

Private Function ListColumn(ByVal ws As Object, ByVal column As Long, ByVal last As Long) As Variant
    ' The cells of a column of MetricItems from row 3 (the first row of a list) to `last`, as a 2-D array (rows, 1), also for one row; Empty when
    ' the list is empty.
    Dim one(1 To 1, 1 To 1) As Variant
    If last < 3 Then Exit Function
    If last = 3 Then one(1, 1) = ws.Cells(3, column).Value: ListColumn = one: Exit Function      ' one cell: Value is no array
    ListColumn = ws.Range(ws.Cells(3, column), ws.Cells(last, column)).Value
End Function

Private Function AsOfOf(ByVal wb As Object) As Double
    ' The AsOf the figures read (the control cell AsOfUtc); NO_AS_OF when it holds no date.
    Dim v As Variant
    AsOfOf = NO_AS_OF
    On Error Resume Next
    v = wb.Names("AsOfUtc").RefersToRange.Value
    If Err.Number <> 0 Then Err.Clear: Exit Function
    On Error GoTo 0
    Dim serial As Double
    If IsSerial(v, serial) Then
        If serial > 0 Then AsOfOf = serial
    End If
End Function

Private Function IsSerial(ByVal v As Variant, ByRef serial As Double) As Boolean
    ' True (and the serial) for a date or a number; IsNumeric alone says False for a Date.
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then Exit Function
    If VarType(v) = vbDate Or IsNumeric(v) Then serial = CDbl(v): IsSerial = True
End Function

' ---- small helpers ----------------------------------------------------------------------------------------------------------------------------

Private Function ReadSpec(ByVal wb As Object) As Object
    ' part -> Collection of Array(name, text, hint), in the order of the sheet RovoDebugSpec.
    Dim out As Object, row As Variant, part As String
    Set out = CreateObject("Scripting.Dictionary")
    If SheetExists(wb, SHEET_ROVO_DEBUG_SPEC) Then
        For Each row In ReadTable(wb.Worksheets(SheetName(SHEET_ROVO_DEBUG_SPEC)))
            part = CStr(row("Part") & "")
            If Not out.Exists(part) Then out.Add part, New Collection
            out(part).Add Array(CStr(row("Name") & ""), CStr(row("Text") & ""), CStr(row("Hint") & ""))
        Next row
    End If
    If Not out.Exists("summary") Then
        Cov "RVD14"
        Err.Raise vbObjectError + 634, "WriteRovoDebug", "The sheet " & SheetName(SHEET_ROVO_DEBUG_SPEC) & " is missing or empty: use a fresh copy of " & _
                  "the delivered workbook."
    End If
    Set ReadSpec = out
End Function

Private Function SpecLines(ByVal spec As Object, ByVal part As String) As Collection
    If spec.Exists(part) Then Set SpecLines = spec(part) Else Set SpecLines = New Collection
End Function

Private Function SpecText(ByVal spec As Object, ByVal part As String, ByVal name As String) As String
    Dim item As Variant
    For Each item In SpecLines(spec, part)
        If item(0) = name Then SpecText = item(1): Exit Function
    Next item
End Function

Private Function Col(ByVal spec As Object, ByVal name As String) As Long
    Col = CLng(SpecText(spec, "column", name))
End Function

Private Function FillLines(ByVal lines As Collection, ByVal pairs As Variant) As String
    Dim out As New Collection, item As Variant
    For Each item In lines
        out.Add RovoFill(CStr(item(1)), pairs)
    Next item
    FillLines = JoinWith(out, vbLf)
End Function

Private Function JoinWith(ByVal items As Collection, ByVal separator As String) As String
    Dim a() As String, i As Long
    If items.Count = 0 Then Exit Function
    ReDim a(1 To items.Count)
    For i = 1 To items.Count
        a(i) = CStr(items(i))
    Next i
    JoinWith = Join(a, separator)
End Function

Private Function Listed(ByVal lines As Collection, ByVal noneText As String) As String
    If lines.Count = 0 Then Listed = noneText Else Listed = JoinWith(lines, vbLf)
End Function

Private Function KeyedList(ByVal lists As Object, ByVal key As String, ByVal none As String) As String
    If lists.Exists(key) Then KeyedList = JoinWith(lists(key), ITEM_SEP) Else KeyedList = none
End Function

Private Function CellText(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    CellText = TrimWhite(CStr(v & ""))
End Function

Private Function IsOne(ByVal v As Variant) As Boolean
    ' 1, "1" or True (never CStr of a Boolean: "Vrai" on a French PC).
    If IsError(v) Then Exit Function
    If VarType(v) = vbBoolean Then IsOne = v: Exit Function
    IsOne = (TrimWhite(CStr(v & "")) = "1")
End Function

Private Function Moment(ByVal serial As Variant) As String
    Moment = Format$(CDate(CDbl(serial) + 1 / 172800), "yyyy-mm-dd hh:nn")    ' + half a second: a serial a hair under the minute stays on it
End Function

Public Sub WriteUtf8File(ByVal path As String, ByVal text As String)
    ' The text as UTF-8 without BOM (modText.Utf8Encode: plain VBA, no ADODB.Stream). An older file is deleted first: Put into a Binary file never
    ' makes it shorter, the end of a longer older file would stay.
    Cov "RVD13"
    Dim f As Integer, bytes() As Byte, num As Long, desc As String
    On Error Resume Next
    Kill path
    Err.Clear
    On Error GoTo Failed
    f = FreeFile
    Open path For Binary Access Write As #f
    If Len(text) > 0 Then bytes = Utf8Encode(text): Put #f, , bytes
    Close #f
    Exit Sub
Failed:
    Cov "RVD15"
    num = Err.Number: desc = Err.Description                            ' first: Close may clear Err
    On Error Resume Next
    Close #f                                                            ' a file left open would refuse the next click
    On Error GoTo 0
    Err.Raise num, "WriteUtf8File", desc
End Sub

Private Sub MakeFolder(ByVal folder As String)
    ' Each missing folder of the path (MkDir fails harmlessly on one that exists).
    Dim parts() As String, i As Long, path As String
    On Error Resume Next
    parts = Split(folder, "\")
    path = parts(0)
    For i = 1 To UBound(parts)
        path = path & "\" & parts(i)
        MkDir path
    Next i
    Err.Clear
End Sub
