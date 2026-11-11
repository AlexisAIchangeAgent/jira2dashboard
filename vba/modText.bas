Option Explicit
' Text helpers: UTF-8, Base64, URL encoding, Jira dates, field values to cells. Plain VBA: no ADODB.Stream, no MSXML.

Public Function Utf8Encode(ByVal text As String) As Byte()
    ' Plain VBA, no ADODB.Stream (an endpoint protection reads a stream next to an HTTP client as a downloader). No BOM.
    ' A lone surrogate becomes U+FFFD. Empty text gives an empty (uninitialised) array.
    Cov "TXT01"
    Dim out() As Byte, u() As Byte, k As Long, last As Long, j As Long, c As Long, d As Long, pair As Boolean
    If Len(text) = 0 Then Utf8Encode = out: Exit Function
    u = text                                                   ' the UTF-16LE bytes of the text
    last = UBound(u) - 1                                       ' index of the low byte of the last unit
    ReDim out(0 To 3 * Len(text) - 1)                          ' at most 3 bytes per UTF-16 unit
    Do While k <= last
        c = u(k) Or (CLng(u(k + 1)) * 256&)
        k = k + 2
        If c < &H80& Then
            out(j) = c: j = j + 1
        ElseIf c < &H800& Then
            Cov "TXT18"
            out(j) = &HC0& Or (c \ 64&): out(j + 1) = &H80& Or (c And &H3F&): j = j + 2
        Else
            pair = False
            If c >= &HD800& And c <= &HDFFF& Then              ' a surrogate: valid only as a high one followed by a low one
                d = 0
                If c <= &HDBFF& And k <= last Then d = u(k) Or (CLng(u(k + 1)) * 256&)
                If d >= &HDC00& And d <= &HDFFF& Then pair = True Else Cov "TXT21": c = &HFFFD&
            End If
            If pair Then
                Cov "TXT20"
                c = &H10000 + (c - &HD800&) * &H400& + (d - &HDC00&)
                k = k + 2
                out(j) = &HF0& Or (c \ &H40000): out(j + 1) = &H80& Or ((c \ &H1000&) And &H3F&)
                out(j + 2) = &H80& Or ((c \ 64&) And &H3F&): out(j + 3) = &H80& Or (c And &H3F&): j = j + 4
            Else
                Cov "TXT19"
                out(j) = &HE0& Or (c \ &H1000&): out(j + 1) = &H80& Or ((c \ 64&) And &H3F&): out(j + 2) = &H80& Or (c And &H3F&): j = j + 3
            End If
        End If
    Loop
    ReDim Preserve out(0 To j - 1)
    Utf8Encode = out
End Function

Public Function Utf8Decode(ByVal bytes As Variant) As String
    ' Plain VBA, no ADODB.Stream. One pass over the bytes writing UTF-16LE into a byte buffer, then the buffer becomes the text
    ' (VBA copies a byte array into a string as it is). Every answer of Jira goes through here: no call inside the ASCII path.
    ' Bad UTF-8 never raises: a byte that starts no valid sequence (stray continuation, overlong form, encoded surrogate, value
    ' above U+10FFFF, truncated sequence) becomes U+FFFD, one per byte, and the text around it is kept. A leading BOM is dropped.
    Cov "TXT02"
    Dim src() As Byte, out() As Byte, lower As Long, upper As Long, i As Long, j As Long, k As Long, b As Long, t As Long
    Dim c As Long, need As Long, ok As Boolean
    On Error Resume Next
    lower = LBound(bytes): upper = UBound(bytes)               ' fails on an uninitialised (empty) array
    If Err.Number <> 0 Then Err.Clear: Exit Function           ' expected: leave no stale error
    On Error GoTo 0
    If upper < lower Then Exit Function
    src = bytes
    ReDim out(0 To 2 * (upper - lower + 1) - 1)                ' at most one UTF-16 unit per byte
    i = lower
    If upper - lower >= 2 Then
        If src(lower) = &HEF& And src(lower + 1) = &HBB& And src(lower + 2) = &HBF& Then Cov "TXT22": i = lower + 3
    End If
    Do While i <= upper
        b = src(i)
        If b < &H80& Then
            out(j) = b: j = j + 2: i = i + 1                   ' ASCII: the high byte of the unit stays 0
        Else
            need = 0
            If b >= &HC2& And b <= &HDF& Then
                need = 1: c = b And &H1F&
            ElseIf b >= &HE0& And b <= &HEF& Then
                need = 2: c = b And &HF&
            ElseIf b >= &HF0& And b <= &HF4& Then
                need = 3: c = b And &H7&
            End If
            ok = False
            If need > 0 And i + need <= upper Then
                ok = True
                For k = 1 To need
                    t = src(i + k)
                    If (t And &HC0&) <> &H80& Then ok = False: Exit For
                    c = c * 64& Or (t And &H3F&)
                Next k
                If ok Then
                    If need = 2 Then ok = (c >= &H800& And (c < &HD800& Or c > &HDFFF&))        ' no overlong form, no surrogate
                    If need = 3 Then ok = (c >= &H10000 And c <= &H10FFFF)                    ' no overlong form, nothing above U+10FFFF
                End If
            End If
            If ok Then
                If need = 3 Then
                    Cov "TXT24"                                ' a character above U+FFFF: a surrogate pair
                    c = c - &H10000
                    t = &HD800& + (c \ &H400&)
                    out(j) = t And &HFF&: out(j + 1) = t \ 256&
                    t = &HDC00& + (c And &H3FF&)
                    out(j + 2) = t And &HFF&: out(j + 3) = t \ 256&: j = j + 4
                Else
                    Cov "TXT23"
                    out(j) = c And &HFF&: out(j + 1) = c \ 256&: j = j + 2
                End If
                i = i + need + 1
            Else
                Cov "TXT25"
                out(j) = &HFD&: out(j + 1) = &HFF&: j = j + 2     ' U+FFFD for this byte
                i = i + 1
            End If
        End If
    Loop
    If j = 0 Then Cov "TXT26": Exit Function                   ' nothing but a BOM
    ReDim Preserve out(0 To j - 1)
    Utf8Decode = out
End Function

Public Function Base64Encode(ByVal bytes As Variant) As String
    ' Plain VBA, no MSXML. RFC 4648 with padding, no line breaks. Empty input gives "".
    Cov "TXT03"
    Const ALPHABET As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    Dim src() As Byte, lower As Long, upper As Long, i As Long, j As Long, out As String, b0 As Long, b1 As Long, b2 As Long
    On Error Resume Next
    lower = LBound(bytes): upper = UBound(bytes)
    If Err.Number <> 0 Then Err.Clear: Cov "TXT27": Exit Function    ' an uninitialised (empty) array
    On Error GoTo 0
    If upper < lower Then Cov "TXT27": Exit Function
    src = bytes
    out = String$(((upper - lower + 3) \ 3) * 4, "=")          ' the padding is already in place
    i = lower: j = 1
    Do While i <= upper
        b0 = src(i): b1 = 0: b2 = 0
        If i + 1 <= upper Then b1 = src(i + 1)
        If i + 2 <= upper Then b2 = src(i + 2)
        Mid$(out, j, 1) = Mid$(ALPHABET, (b0 \ 4&) + 1, 1)
        Mid$(out, j + 1, 1) = Mid$(ALPHABET, (((b0 And 3&) * 16&) Or (b1 \ 16&)) + 1, 1)
        If i + 1 <= upper Then Mid$(out, j + 2, 1) = Mid$(ALPHABET, (((b1 And 15&) * 4&) Or (b2 \ 64&)) + 1, 1)
        If i + 2 <= upper Then Mid$(out, j + 3, 1) = Mid$(ALPHABET, (b2 And 63&) + 1, 1) Else Cov "TXT28"
        i = i + 3: j = j + 4
    Loop
    Base64Encode = out
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
    On Error GoTo Invalid
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
    Exit Function
Invalid:
    Cov "TXT17"
    Err.Clear
    ParseJiraDate = Empty
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

Private Function AdfToText(ByVal node As Variant) As String
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

Public Function StatusParts(ByVal text As String) As Collection
    ' "Title | Name: value | Name: value" -> Array("", "Title"), Array("Name", "value"), ...: the status of the Dashboard as
    ' its headline (cell A3) and one row per field under it. A field is cut at its first ": " (a value may hold more); a part
    ' without ": " is a row with no name. A text without " | " is the headline alone (the progress of a run).
    Dim out As New Collection, parts() As String, i As Long, p As Long
    Cov "TXT29"
    Set StatusParts = out
    If Len(text) = 0 Then out.Add Array("", ""): Exit Function
    parts = Split(text, " | ")
    out.Add Array("", parts(0))
    For i = 1 To UBound(parts)
        p = InStr(parts(i), ": ")
        If p > 0 Then
            out.Add Array(Left$(parts(i), p - 1), Mid$(parts(i), p + 2))
        Else
            Cov "TXT30"
            out.Add Array("", parts(i))
        End If
    Next i
End Function
