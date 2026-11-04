Option Explicit
' JSON parser and writer. No references needed (late binding).
' Objects -> Scripting.Dictionary (keys case-sensitive), arrays -> Collection, null -> Null.
' Numbers use Val(), which ignores the Windows locale (French uses a comma).

Private mText As String
Private mPos As Long
Private mLen As Long
Private mNextBackslash As Long

Public Function JsonParse(ByVal text As String) As Variant
    Cov "JSN01"
    mText = text: mPos = 1: mLen = Len(text): mNextBackslash = 0
    Dim result As Variant
    ParseValue result
    SkipWhitespace
    If mPos <= mLen Then Cov "JSN02": Fail "Unexpected text after JSON"
    If IsObject(result) Then Set JsonParse = result Else JsonParse = result
End Function

Private Sub ParseValue(ByRef target As Variant)
    SkipWhitespace
    If mPos > mLen Then Cov "JSN03": Fail "Unexpected end of JSON"
    Select Case Mid$(mText, mPos, 1)
        Case "{": Set target = ParseObject()
        Case "[": Set target = ParseArray()
        Case """": target = ParseString()
        Case "t": ExpectWord "true": target = True
        Case "f": ExpectWord "false": target = False
        Case "n": ExpectWord "null": target = Null
        Case Else: target = ParseNumber()
    End Select
End Sub

Private Function ParseObject() As Object
    Cov "JSN04"
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 0
    mPos = mPos + 1
    SkipWhitespace
    If Mid$(mText, mPos, 1) = "}" Then mPos = mPos + 1: Set ParseObject = d: Exit Function
    Do
        SkipWhitespace
        If Mid$(mText, mPos, 1) <> """" Then Cov "JSN05": Fail "Expected a key"
        Dim key As String
        key = ParseString()
        SkipWhitespace
        If Mid$(mText, mPos, 1) <> ":" Then Cov "JSN06": Fail "Expected ':'"
        mPos = mPos + 1
        ParseMember d, key
        SkipWhitespace
        Select Case Mid$(mText, mPos, 1)
            Case ",": mPos = mPos + 1
            Case "}": mPos = mPos + 1: Exit Do
            Case Else: Cov "JSN07": Fail "Expected ',' or '}'"
        End Select
    Loop
    Set ParseObject = d
End Function

Private Sub ParseMember(ByVal d As Object, ByVal key As String)
    Dim v As Variant
    ParseValue v
    If IsObject(v) Then Set d(key) = v Else d(key) = v
End Sub

Private Function ParseArray() As Collection
    Cov "JSN08"
    Dim c As New Collection
    mPos = mPos + 1
    SkipWhitespace
    If Mid$(mText, mPos, 1) = "]" Then mPos = mPos + 1: Set ParseArray = c: Exit Function
    Do
        ParseItem c
        SkipWhitespace
        Select Case Mid$(mText, mPos, 1)
            Case ",": mPos = mPos + 1
            Case "]": mPos = mPos + 1: Exit Do
            Case Else: Cov "JSN09": Fail "Expected ',' or ']'"
        End Select
    Loop
    Set ParseArray = c
End Function

Private Sub ParseItem(ByVal c As Collection)
    Dim v As Variant
    ParseValue v
    c.Add v
End Sub

Private Function ParseString() As String
    Dim out As String, q As Long, esc As String
    mPos = mPos + 1
    Do
        q = InStr(mPos, mText, """")
        If q = 0 Then Cov "JSN10": Fail "Unterminated string"
        If mNextBackslash < mPos And mNextBackslash <> -1 Then
            mNextBackslash = InStr(mPos, mText, "\")
            If mNextBackslash = 0 Then mNextBackslash = -1
        End If
        If mNextBackslash = -1 Or mNextBackslash > q Then
            ParseString = out & Mid$(mText, mPos, q - mPos)
            mPos = q + 1
            Exit Function
        End If
        Cov "JSN11"
        out = out & Mid$(mText, mPos, mNextBackslash - mPos)
        esc = Mid$(mText, mNextBackslash + 1, 1)
        mPos = mNextBackslash + 2
        Select Case esc
            Case """", "\", "/": out = out & esc
            Case "b": out = out & Chr$(8)
            Case "f": out = out & Chr$(12)
            Case "n": out = out & vbLf
            Case "r": out = out & vbCr
            Case "t": out = out & vbTab
            Case "u": out = out & ChrW$(CLng("&H" & Mid$(mText, mPos, 4))): mPos = mPos + 4
            Case Else: Cov "JSN12": Fail "Bad escape \" & esc
        End Select
    Loop
End Function

Private Function ParseNumber() As Double
    Dim startPos As Long
    startPos = mPos
    Do While mPos <= mLen
        If InStr("+-0123456789.eE", Mid$(mText, mPos, 1)) = 0 Then Exit Do
        mPos = mPos + 1
    Loop
    If mPos = startPos Then Cov "JSN13": Fail "Unexpected character '" & Mid$(mText, mPos, 1) & "'"
    ParseNumber = Val(Mid$(mText, startPos, mPos - startPos))
End Function

Private Sub ExpectWord(ByVal word As String)
    If Mid$(mText, mPos, Len(word)) <> word Then Cov "JSN14": Fail "Expected " & word
    mPos = mPos + Len(word)
End Sub

Private Sub SkipWhitespace()
    Do While mPos <= mLen
        Select Case AscW(Mid$(mText, mPos, 1))
            Case 32, 9, 10, 13: mPos = mPos + 1
            Case Else: Exit Do
        End Select
    Loop
End Sub

Private Sub Fail(ByVal message As String)
    Err.Raise vbObjectError + 513, "modJson", message & " at position " & mPos
End Sub

' ---------- Writer ----------

Public Function JsonStringify(ByVal value As Variant) As String
    If IsObject(value) Then
        If TypeName(value) = "Dictionary" Then
            Cov "JSN15": JsonStringify = StringifyObject(value)
        ElseIf TypeName(value) = "Collection" Then
            Cov "JSN16": JsonStringify = StringifyArray(value)
        Else
            Cov "JSN17": Err.Raise vbObjectError + 514, "modJson", "Cannot write " & TypeName(value)
        End If
    ElseIf IsNull(value) Or IsEmpty(value) Then
        JsonStringify = "null"
    ElseIf VarType(value) = vbBoolean Then
        JsonStringify = IIf(value, "true", "false")
    ElseIf IsNumeric(value) And VarType(value) <> vbString Then
        JsonStringify = Trim$(Str$(value))
    ElseIf IsArray(value) Then
        Cov "JSN18": JsonStringify = StringifyVbArray(value)
    Else
        JsonStringify = """" & JsonEscape(CStr(value)) & """"
    End If
End Function

Private Function StringifyObject(ByVal d As Object) As String
    Dim parts() As String, i As Long, k As Variant
    If d.Count = 0 Then StringifyObject = "{}": Exit Function
    ReDim parts(0 To d.Count - 1)
    For Each k In d.Keys
        parts(i) = """" & JsonEscape(CStr(k)) & """:" & JsonStringify(d(k))
        i = i + 1
    Next k
    StringifyObject = "{" & Join(parts, ",") & "}"
End Function

Private Function StringifyArray(ByVal c As Collection) As String
    Dim parts() As String, i As Long
    If c.Count = 0 Then StringifyArray = "[]": Exit Function
    ReDim parts(1 To c.Count)
    For i = 1 To c.Count
        parts(i) = JsonStringify(c(i))
    Next i
    StringifyArray = "[" & Join(parts, ",") & "]"
End Function

Private Function StringifyVbArray(ByVal a As Variant) As String
    Dim parts() As String, i As Long
    If UBound(a) < LBound(a) Then StringifyVbArray = "[]": Exit Function
    ReDim parts(LBound(a) To UBound(a))
    For i = LBound(a) To UBound(a)
        parts(i) = JsonStringify(a(i))
    Next i
    StringifyVbArray = "[" & Join(parts, ",") & "]"
End Function

Public Function JsonEscape(ByVal s As String) As String
    Dim i As Long, ch As String, code As Long, out As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCr, "\r")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)
        If code >= 0 And code < 32 Then out = out & "\u" & Right$("000" & Hex$(code), 4) Else out = out & ch
    Next i
    JsonEscape = out
End Function

' ---------- Safe access helpers ----------

Public Function JGet(ByVal node As Variant, ByVal key As String) As Variant
    ' Returns node(key), or Empty if node is not a Dictionary or the key is missing.
    JGet = Empty
    If Not IsObject(node) Then Exit Function
    If TypeName(node) <> "Dictionary" Then Exit Function
    If Not node.Exists(key) Then Exit Function
    If IsObject(node(key)) Then Set JGet = node(key) Else JGet = node(key)
End Function

Public Function JPath(ByVal node As Variant, ByVal path As String) As Variant
    ' JPath(issue, "fields.status.name"). Returns Empty when any step is missing.
    Dim parts() As String, i As Long, cur As Variant
    parts = Split(path, ".")
    If IsObject(node) Then Set cur = node Else cur = node
    For i = LBound(parts) To UBound(parts)
        If IsObject(JGet(cur, parts(i))) Then
            Set cur = JGet(cur, parts(i))
        Else
            cur = JGet(cur, parts(i))
        End If
    Next i
    If IsObject(cur) Then Set JPath = cur Else JPath = cur
End Function

Public Function AsList(ByVal value As Variant) As Collection
    ' Returns value if it is a Collection (JSON array), else an empty Collection. Safe for For Each.
    If IsObject(value) Then
        If TypeName(value) = "Collection" Then Set AsList = value: Exit Function
    End If
    Set AsList = New Collection
End Function

Public Function IsTrue(ByVal value As Variant) As Boolean
    ' True only for a real Boolean True. Safe with Null and Empty.
    If VarType(value) = vbBoolean Then IsTrue = value
End Function
