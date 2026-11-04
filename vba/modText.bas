Option Explicit
' Text helpers: UTF-8, Base64, URL encoding, Jira dates, field values to cells.

Public Function Utf8Encode(ByVal text As String) As Byte()
    Cov "TXT01"
    Dim st As Object, empty_() As Byte
    If Len(text) = 0 Then Utf8Encode = empty_: Exit Function
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8": st.Open
    st.WriteText text
    st.Position = 0: st.Type = 1: st.Position = 3   ' skip the UTF-8 BOM
    Utf8Encode = st.Read
    st.Close
End Function

Public Function Utf8Decode(ByVal bytes As Variant) As String
    Cov "TXT02"
    Dim st As Object, upper As Long
    On Error Resume Next
    upper = UBound(bytes)                              ' fails on an uninitialised (empty) array
    If Err.Number <> 0 Then Err.Clear: Exit Function   ' expected: leave no stale error
    On Error GoTo 0
    If upper < LBound(bytes) Then Exit Function
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1: st.Open
    st.Write bytes
    st.Position = 0: st.Type = 2: st.Charset = "utf-8"
    Utf8Decode = st.ReadText
    st.Close
End Function

Public Function Base64Encode(ByVal bytes As Variant) As String
    Cov "TXT03"
    Dim node As Object
    Set node = CreateObject("MSXML2.DOMDocument.6.0").createElement("b64")
    node.DataType = "bin.base64"
    node.nodeTypedValue = bytes
    Base64Encode = Replace(Replace(node.Text, vbLf, ""), vbCr, "")
End Function

Public Function BasicAuthHeader(ByVal email As String, ByVal apiToken As String) As String
    Cov "TXT04"
    BasicAuthHeader = "Basic " & Base64Encode(Utf8Encode(email & ":" & apiToken))
End Function

Public Function TrimWhite(ByVal text As String) As String
    ' Like PowerShell .Trim(): removes spaces, tabs, line breaks and non-breaking spaces at both ends.
    ' VBA Trim$ removes spaces only, and a pasted API token often ends with a line break.
    Cov "TXT16"
    Dim first As Long, last As Long
    first = 1: last = Len(text)
    Do While first <= last
        If Not IsWhiteChar(Mid$(text, first, 1)) Then Exit Do
        first = first + 1
    Loop
    Do While last >= first
        If Not IsWhiteChar(Mid$(text, last, 1)) Then Exit Do
        last = last - 1
    Loop
    TrimWhite = Mid$(text, first, last - first + 1)
End Function

Private Function IsWhiteChar(ByVal ch As String) As Boolean
    IsWhiteChar = (ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Or ch = ChrW(160))
End Function

Public Function UrlEncode(ByVal text As String) As String
    Cov "TXT05"
    Dim b() As Byte, i As Long, out As String
    If Len(text) = 0 Then Exit Function
    b = Utf8Encode(text)
    For i = LBound(b) To UBound(b)
        Select Case b(i)
            Case 48 To 57, 65 To 90, 97 To 122, 45, 46, 95, 126   ' 0-9 A-Z a-z - . _ ~
                out = out & Chr$(b(i))
            Case Else
                out = out & "%" & Right$("0" & Hex$(b(i)), 2)
        End Select
    Next i
    UrlEncode = out
End Function

Public Function ParseJiraDate(ByVal value As Variant) As Variant
    ' "2026-10-01T09:12:33.123+0200", "...+02:00", "...Z" -> Date in UTC.
    ' "2026-10-01" -> that date. Empty or invalid -> Empty (never raises).
    Dim s As String, d As Date, offsetMin As Long, tail As String, sign As Long
    ParseJiraDate = Empty
    If IsNull(value) Or IsEmpty(value) Then Cov "TXT06": Exit Function
    s = CStr(value)
    If Len(s) < 10 Then Cov "TXT07": Exit Function
    If Mid$(s, 5, 1) <> "-" Or Mid$(s, 8, 1) <> "-" Then Cov "TXT08": Exit Function
    d = DateSerial(CInt(Left$(s, 4)), CInt(Mid$(s, 6, 2)), CInt(Mid$(s, 9, 2)))
    If Len(s) = 10 Then Cov "TXT09": ParseJiraDate = d: Exit Function
    Cov "TXT10"
    d = d + TimeSerial(CInt(Mid$(s, 12, 2)), CInt(Mid$(s, 15, 2)), CInt(Mid$(s, 18, 2)))
    tail = Mid$(s, 20)
    Do While Len(tail) > 0 And InStr("+-Zz", Left$(tail, 1)) = 0
        tail = Mid$(tail, 2)                              ' skip ".123"
    Loop
    If Len(tail) > 0 And UCase$(Left$(tail, 1)) <> "Z" Then
        sign = IIf(Left$(tail, 1) = "-", -1, 1)
        tail = Replace(Mid$(tail, 2), ":", "")
        offsetMin = sign * (CLng(Left$(tail, 2)) * 60 + CLng(Mid$(tail, 3, 2)))
    End If
    ParseJiraDate = d - offsetMin / 1440#
End Function

Public Function FieldToCell(ByVal value As Variant) As Variant
    ' Turn any Jira field value into something a cell can hold.
    If IsObject(value) Then
        If TypeName(value) = "Collection" Then
            Cov "TXT11": FieldToCell = JoinCollection(value)
        Else
            Cov "TXT12": FieldToCell = ObjectToText(value)
        End If
    ElseIf IsNull(value) Or IsEmpty(value) Then
        Cov "TXT13": FieldToCell = Empty
    ElseIf VarType(value) = vbString Then
        FieldToCell = Left$(value, 32767)
    Else
        FieldToCell = value
    End If
End Function

Private Function JoinCollection(ByVal c As Collection) As String
    Dim parts() As String, i As Long
    If c.Count = 0 Then Exit Function
    ReDim parts(1 To c.Count)
    For i = 1 To c.Count
        parts(i) = CStr(FieldToCell(c(i)))
    Next i
    JoinCollection = Left$(Join(parts, "; "), 32767)
End Function

Private Function ObjectToText(ByVal d As Object) As String
    Dim k As Variant
    If d.Exists("type") Then
        If Not IsObject(d("type")) Then
            If d("type") = "doc" Then Cov "TXT14": ObjectToText = Left$(AdfToText(d), 32767): Exit Function
        End If
    End If
    For Each k In Array("displayName", "value", "name", "key", "id")
        If d.Exists(k) Then
            If Not IsObject(d(k)) And Not IsNull(d(k)) Then ObjectToText = CStr(d(k)): Exit Function
        End If
    Next k
    Cov "TXT15"
    ObjectToText = Left$(JsonStringify(d), 32767)
End Function

Public Function AdfToText(ByVal node As Variant) As String
    ' Atlassian Document Format (rich text, e.g. description) -> plain text.
    Dim out As String, child As Variant, nodeType As String
    If Not IsObject(node) Then Exit Function
    If TypeName(node) <> "Dictionary" Then Exit Function
    nodeType = CStr(JGet(node, "type") & "")
    If node.Exists("text") Then out = CStr(node("text") & "")
    If node.Exists("content") Then
        For Each child In node("content")
            out = out & AdfToText(child)
        Next child
        If nodeType = "paragraph" Or nodeType = "heading" Or nodeType = "listItem" Then out = out & vbLf
    End If
    AdfToText = out
End Function
