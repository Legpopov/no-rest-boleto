Attribute VB_Name = "modJson"
' Минимальный JSON-парсер без внешних библиотек и ссылок (References).
'   объект -> Scripting.Dictionary (позднее связывание)
'   массив -> Variant() с индексами от 0 (пустой массив: UBound = -1)
'   число  -> Double, true/false -> Boolean, null -> Null
' Источник: github.com/Legpopov/no-rest-boleto (excel/modJson.bas), лицензия MIT.
Option Explicit

Private s As String      ' разбираемый текст
Private p As Long        ' текущая позиция (с 1)
Private n As Long        ' длина текста
Private nextBs As Long   ' позиция ближайшего "\" (кэш, чтобы не искать заново для каждой строки)

Public Function JsonParse(ByVal text As String) As Variant
    Dim v As Variant
    s = text: p = 1: n = Len(s): nextBs = 0
    If n > 0 Then
        If AscW(Left$(s, 1)) = &HFEFF Then p = 2   ' BOM
    End If
    ParseValue v
    SkipWs
    If p <= n Then Fail "лишние символы после конца JSON"
    s = vbNullString
    If IsObject(v) Then Set JsonParse = v Else JsonParse = v
End Function

' Значение по пути "air.all.median"; Empty, если какого-то ключа нет.
Public Function JGet(ByVal node As Variant, ByVal path As String) As Variant
    Dim parts() As String, i As Long
    parts = Split(path, ".")
    For i = 0 To UBound(parts)
        If Not IsObject(node) Then JGet = Empty: Exit Function
        If node Is Nothing Then JGet = Empty: Exit Function
        If TypeName(node) <> "Dictionary" Then JGet = Empty: Exit Function
        If Not node.Exists(parts(i)) Then JGet = Empty: Exit Function
        If IsObject(node(parts(i))) Then Set node = node(parts(i)) Else node = node(parts(i))
    Next
    If IsObject(node) Then Set JGet = node Else JGet = node
End Function

Public Function JIsArray(ByVal v As Variant) As Boolean
    JIsArray = IsArray(v)
End Function

' ---------- разбор ----------

Private Sub ParseValue(ByRef v As Variant)
    SkipWs
    If p > n Then Fail "неожиданный конец текста"
    Select Case AscW(Mid$(s, p, 1))
    Case 123: Set v = ParseObject()          ' {
    Case 91: v = ParseArray()                ' [
    Case 34: v = ParseString()               ' "
    Case 116: Expect "true": v = True
    Case 102: Expect "false": v = False
    Case 110: Expect "null": v = Null
    Case Else: v = ParseNumber()
    End Select
End Sub

Private Function ParseObject() As Object
    Dim d As Object, k As String, v As Variant
    Set d = CreateObject("Scripting.Dictionary")
    p = p + 1
    SkipWs
    If Mid$(s, p, 1) = "}" Then
        p = p + 1
        Set ParseObject = d
        Exit Function
    End If
    Do
        SkipWs
        If Mid$(s, p, 1) <> """" Then Fail "ожидался ключ в кавычках"
        k = ParseString()
        SkipWs
        If Mid$(s, p, 1) <> ":" Then Fail "ожидалось ':'"
        p = p + 1
        v = Empty
        ParseValue v
        If IsObject(v) Then Set d(k) = v Else d(k) = v
        SkipWs
        Select Case Mid$(s, p, 1)
        Case ",": p = p + 1
        Case "}": p = p + 1: Exit Do
        Case Else: Fail "ожидалось ',' или '}'"
        End Select
    Loop
    Set ParseObject = d
End Function

Private Function ParseArray() As Variant
    Dim a() As Variant, cnt As Long, v As Variant
    p = p + 1
    SkipWs
    If Mid$(s, p, 1) = "]" Then
        p = p + 1
        ParseArray = Array()
        Exit Function
    End If
    ReDim a(0 To 7)
    Do
        v = Empty
        ParseValue v
        If cnt > UBound(a) Then ReDim Preserve a(0 To 2 * cnt - 1)
        If IsObject(v) Then Set a(cnt) = v Else a(cnt) = v
        cnt = cnt + 1
        SkipWs
        Select Case Mid$(s, p, 1)
        Case ",": p = p + 1
        Case "]": p = p + 1: Exit Do
        Case Else: Fail "ожидалось ',' или ']'"
        End Select
    Loop
    ReDim Preserve a(0 To cnt - 1)
    ParseArray = a
End Function

Private Function ParseString() As String
    Dim q As Long, out As String, c As String
    p = p + 1
    Do
        q = InStr(p, s, """")
        If q = 0 Then Fail "незакрытая строка"
        If nextBs <> -1 And nextBs < p Then
            nextBs = InStr(p, s, "\")
            If nextBs = 0 Then nextBs = -1
        End If
        If nextBs = -1 Or nextBs > q Then
            ParseString = out & Mid$(s, p, q - p)
            p = q + 1
            Exit Function
        End If
        out = out & Mid$(s, p, nextBs - p)
        c = Mid$(s, nextBs + 1, 1)
        Select Case c
        Case """", "\", "/": out = out & c
        Case "b": out = out & ChrW(8)
        Case "f": out = out & ChrW(12)
        Case "n": out = out & vbLf
        Case "r": out = out & vbCr
        Case "t": out = out & vbTab
        Case "u"
            out = out & ChrW(Val("&H" & Mid$(s, nextBs + 2, 4) & "&"))
            nextBs = nextBs + 4
        Case Else: Fail "неверная escape-последовательность"
        End Select
        p = nextBs + 2
    Loop
End Function

Private Function ParseNumber() As Double
    Dim q As Long
    q = p
    Do While q <= n
        Select Case AscW(Mid$(s, q, 1))
        Case 48 To 57, 43, 45, 46, 69, 101: q = q + 1   ' 0-9 + - . E e
        Case Else: Exit Do
        End Select
    Loop
    If q = p Then Fail "неожиданный символ '" & Mid$(s, p, 1) & "'"
    ParseNumber = Val(Mid$(s, p, q - p))   ' Val не зависит от региональных настроек
    p = q
End Function

Private Sub SkipWs()
    Do While p <= n
        Select Case AscW(Mid$(s, p, 1))
        Case 32, 9, 10, 13: p = p + 1
        Case Else: Exit Sub
        End Select
    Loop
End Sub

Private Sub Expect(ByVal word As String)
    If Mid$(s, p, Len(word)) <> word Then Fail "ожидалось '" & word & "'"
    p = p + Len(word)
End Sub

Private Sub Fail(ByVal msg As String)
    Err.Raise vbObjectError + 700, "JsonParse", "Ошибка JSON на позиции " & p & ": " & msg
End Sub
