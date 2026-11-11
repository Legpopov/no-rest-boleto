Attribute VB_Name = "modPlanning"
Option Explicit

' Offline semester forecast. No network access from worksheet functions.
' Reference fares are one-way RUB; annual reference changes repeat as a scenario.
Private Const TRIPS As String = "Командировки и выезды"
Private Const PARAMS As String = "Параметры цен"
Private Const DATA As String = "Данные плана"

Public Sub RefreshPlanning(Optional ByVal source As String = "")
    On Error GoTo Failed
    Dim req As Object, root As Object, records As Variant, rec As Object
    Dim text As String, stream As Object, matrix() As Variant, i As Long, j As Long
    Dim ws As Worksheet, ps As Worksheet, fields As Variant, stamp As String
    Dim coverage As Variant, coverageMatrix() As Variant, coverageFields As Variant
    Set ps = ThisWorkbook.Worksheets(PARAMS)
    If source = "" Then source = CStr(ps.Range("O2").Value2)
    If LCase$(Left$(source, 8)) = "https://" Then
        Set req = CreateObject("WinHttp.WinHttpRequest.5.1")
        req.SetTimeouts 10000, 10000, 30000, 30000
        req.Open "GET", source, False
        req.Send
        If req.Status <> 200 Then Err.Raise 5, , "HTTP " & req.Status
        text = req.ResponseText
    Else
        Set stream = CreateObject("ADODB.Stream")
        stream.Type = 2: stream.Charset = "utf-8": stream.Open
        stream.LoadFromFile source
        text = stream.ReadText: stream.Close
    End If
    Set root = JsonParse(text)
    If root("schema_version") <> 2 Or root("currency") <> "RUB" Or root("basis") <> "one_way" Then
        Err.Raise 5, , "Неподдерживаемая схема плановых цен"
    End If
    records = root("records")
    If UBound(records) < 0 Then Err.Raise 5, , "Пустой набор плановых цен"
    fields = Array("origin", "destination", "mode", "month", "base_year", "price", _
                   "annual_rate", "baggage_included", "source", "previous_source", _
                   "reference_date", "previous_date", "rate_method", "quote_id", _
                   "rail_class", "firm", "sample_count", "price_method", "uncertainty")
    ReDim matrix(1 To UBound(records) + 1, 1 To 19)
    For i = 0 To UBound(records)
        Set rec = records(i)
        If rec("price") <= 0 Or rec("annual_rate") <= -1 Then Err.Raise 5, , "Некорректная цена или индекс"
        For j = 0 To 18
            If rec.Exists(fields(j)) Then matrix(i + 1, j + 1) = rec(fields(j))
        Next j
        If matrix(i + 1, 17) = "" Then matrix(i + 1, 17) = 1
    Next i
    coverage = root("coverage")
    coverageFields = Array("origin", "destination", "mode", "status", "checked_at", "source")
    ReDim coverageMatrix(1 To UBound(coverage) + 1, 1 To 6)
    For i = 0 To UBound(coverage)
        Set rec = coverage(i)
        For j = 0 To 5
            coverageMatrix(i + 1, j + 1) = rec(coverageFields(j))
        Next j
    Next i
    BackupPlan
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(DATA)
    On Error GoTo Failed
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = DATA
    End If
    ws.Cells.ClearContents
    For j = 0 To 18
        ws.Cells(1, j + 1).Value2 = fields(j)
    Next j
    ws.Range("A2").Resize(UBound(matrix, 1), 19).Value2 = matrix
    For j = 0 To 5
        ws.Cells(1, j + 21).Value2 = coverageFields(j)
    Next j
    ws.Range("U2").Resize(UBound(coverageMatrix, 1), 6).Value2 = coverageMatrix
    ws.Range("F2:F" & UBound(matrix, 1) + 1).NumberFormat = "#,##0.00"
    ws.Range("G2:G" & UBound(matrix, 1) + 1).NumberFormat = "0.00%"
    ws.Rows(1).Font.Bold = True
    ws.Columns("A:N").ColumnWidth = 20
    ws.Columns("I:J").ColumnWidth = 45
    ThisWorkbook.Names.Add Name:="PlanningData", RefersTo:="='" & DATA & "'!$A$2:$S$" & UBound(matrix, 1) + 1
    ThisWorkbook.Names.Add Name:="PlanningCoverage", RefersTo:="='" & DATA & "'!$U$2:$Z$" & UBound(coverageMatrix, 1) + 1
    ps.Range("O8").Value2 = root("retrieved_at")
    ps.Range("O9").Value2 = "Авиа: АТОР; ЖД/автобус: резервные ориентиры, без сезонности, макросценарий ЦБ. Точность не оценена."
    Application.CalculateFull
    Exit Sub
Failed:
    Err.Raise Err.Number, "RefreshPlanning", Err.Description
End Sub

Private Sub BackupPlan()
    Dim folder As String, file As String, files As Object, item As Object
    Dim oldest As Object, count As Long
    If ThisWorkbook.Path = "" Then Err.Raise 5, , "Сначала сохраните книгу"
    folder = ThisWorkbook.Path & "\_backup_план"
    Set files = CreateObject("Scripting.FileSystemObject")
    If Not files.FolderExists(folder) Then files.CreateFolder folder
    file = folder & "\" & files.GetBaseName(ThisWorkbook.Name) & "_" & Format(Now, "yyyymmdd_hhnnss") & ".xlsm"
    ThisWorkbook.SaveCopyAs file
    Do
        count = 0: Set oldest = Nothing
        For Each item In files.GetFolder(folder).Files
            If Left$(item.Name, Len(files.GetBaseName(ThisWorkbook.Name)) + 1) = files.GetBaseName(ThisWorkbook.Name) & "_" Then
                count = count + 1
                If oldest Is Nothing Then
                    Set oldest = item
                ElseIf item.DateLastModified < oldest.DateLastModified Then
                    Set oldest = item
                End If
            End If
        Next item
        If count <= 3 Then Exit Do
        oldest.Delete
    Loop
End Sub

Public Sub InstallPlanning()
    Dim ws As Worksheet, ps As Worksheet, tbl As ListObject, headers As Variant
    Dim r As Long, c As Long, args As String, emptyRow As String, guard As String
    Dim railCell As Range, extra As String
    Set ws = ThisWorkbook.Worksheets(TRIPS)
    Set ps = ThisWorkbook.Worksheets(PARAMS)
    Set tbl = ws.ListObjects("Таблица3")
    Set railCell = ps.Columns(1).Find(What:="ЖД: допустимые классы", LookIn:=xlValues, LookAt:=xlWhole)
    If railCell Is Nothing Then Err.Raise 5, , "Нет параметра ЖД: допустимые классы"
    extra = ",'Параметры цен'!$B$" & railCell.Row & ",PlanningCoverage"
    headers = Array("Откуда", "Поездка ID", "Участок", "Остановка", "Ночей", _
                    "Порядок транспорта", "Багаж авиа", "Авиа, руб./чел.", "Поезд, руб./чел.", _
                    "Автобус, руб./чел.", "Выбранный транспорт", "Метод / причина")
    BackupPlan
    tbl.Resize ws.Range("A2:W" & tbl.Range.Row + tbl.Range.Rows.Count - 1)
    For c = 0 To UBound(headers)
        ws.Cells(2, c + 12).Value2 = headers(c)
    Next c
    ps.Range("N2").Value2 = "Источник плановых цен"
    If ps.Range("O2").Value2 = "" Then ps.Range("O2").Value2 = "https://raw.githubusercontent.com/Legpopov/no-rest-boleto/fix/sem-route-planning/data/planning.json"
    ps.Range("N3").Value2 = "Осень: первый месяц"
    ps.Range("N4").Value2 = "Осень: число месяцев"
    ps.Range("N5").Value2 = "Весна: первый месяц"
    ps.Range("N6").Value2 = "Весна: число месяцев"
    If ps.Range("O3").Value2 = "" Then ps.Range("O3:O6").Value = Application.Transpose(Array(9, 5, 2, 5))
    ps.Range("N8").Value2 = "Снимок источников (UTC)"
    ps.Range("N9").Value2 = "Допущение"
    ps.Columns("N").ColumnWidth = 30
    ps.Columns("O").ColumnWidth = 65
    ps.Range("O9").WrapText = True
    ps.Rows(9).RowHeight = 32
    For r = 3 To tbl.Range.Row + tbl.Range.Rows.Count - 1
        If InStr(1, ws.Cells(r, 8).Value2, "транспортные услуги", vbTextCompare) > 0 Then
            If ws.Cells(r, 17).Value2 = "" Then ws.Cells(r, 17).Value2 = "авиа, поезд, автобус"
            If ws.Cells(r, 18).Value2 = "" Then ws.Cells(r, 18).Value2 = "ДА"
            emptyRow = "AND(C" & r & "="""",F" & r & "="""",L" & r & "="""")"
            args = "L" & r & ",F" & r & ",'Параметры'!$C$2,B" & r & ",A" & r & ",Q" & r & ",R" & r & ",'Параметры цен'!$O$3:$O$6,PlanningData,"
            guard = "PlanRouteIssue(ROW(),$L$3:$N$" & tbl.Range.Rows.Count + 1 & ",$F$3:$F$" & tbl.Range.Rows.Count + 1 & ",$A$3:$C$" & tbl.Range.Rows.Count + 1 & ",$H$3:$H$" & tbl.Range.Rows.Count + 1 & ")"
            For c = 19 To 21
                ws.Cells(r, c).Formula = "=IF(" & emptyRow & ","""",PlanFare(" & args & """" & Choose(c - 18, "air", "rail", "bus") & """" & extra & "))"
            Next c
            ws.Cells(r, 22).Formula = "=IF(" & emptyRow & ","""",PlanFare(" & args & """mode""" & extra & "))"
            ws.Cells(r, 23).Formula = "=IF(" & emptyRow & ","""",IF(" & guard & "<>""""," & guard & ",PlanFare(" & args & """status""" & extra & ")))"
            ws.Cells(r, 10).Formula = "=IF(" & emptyRow & ","""",IF(OR(C" & r & "<=0,NOT(ISNUMBER(C" & r & "))," & guard & "<>""""),NA(),C" & r & "*PlanFare(" & args & """price""" & extra & ")))"
            If ws.Rows(r).RowHeight < 65 Then ws.Rows(r).RowHeight = 65
        ElseIf InStr(1, ws.Cells(r, 8).Value2, "оплата проживания", vbTextCompare) > 0 Then
            emptyRow = "AND(C" & r & "="""",F" & r & "="""",O" & r & "="""",P" & r & "="""")"
            args = "C" & r & ",F" & r & ",O" & r & ",P" & r & ",'Справочник'!$BK$4:$BR$248,"
            ws.Cells(r, 10).Formula = "=IF(" & emptyRow & ","""",PlanStay(" & args & """price""))"
            ws.Cells(r, 23).Formula = "=IF(" & emptyRow & ","""",PlanStay(" & args & """status""))"
        End If
    Next r
    ws.Range("C3:C" & r - 1).Copy
    ws.Range("L3:R" & r - 1).PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    ws.Range("L2:W2").WrapText = True
    ws.Columns("L:M").ColumnWidth = 23
    ws.Columns("N:P").ColumnWidth = 13
    ws.Columns("Q").ColumnWidth = 28
    ws.Columns("R").ColumnWidth = 10
    ws.Columns("S:V").ColumnWidth = 20
    ws.Columns("W").ColumnWidth = 70
    ws.Range("S3:U" & r - 1).NumberFormat = "#,##0.00"
    ws.Range("L3:R" & r - 1).Font.Color = RGB(0, 0, 255)
    ws.Range("W3:W" & r - 1).WrapText = True
    ws.Range("V3:V" & r - 1).HorizontalAlignment = xlCenter
    With ws.Range("O3:O" & r - 1).Validation
        .Delete
        .Add xlValidateList, xlValidAlertStop, xlBetween, "пребывание,пересадка"
    End With
    With ws.Range("R3:R" & r - 1).Validation
        .Delete
        .Add xlValidateList, xlValidAlertStop, xlBetween, "ДА,НЕТ"
    End With
    With ps.Range("O3:O6").Validation
        .Delete
        .Add xlValidateWholeNumber, xlValidAlertStop, xlBetween, 1, 12
    End With
    Dim button As Shape
    On Error Resume Next
    ps.Shapes("RefreshPlanButton").Delete
    On Error GoTo 0
    Set button = ps.Shapes.AddFormControl(xlButtonControl, ps.Range("N11").Left, ps.Range("N11").Top, 200, 25)
    button.Name = "RefreshPlanButton"
    button.TextFrame.Characters.Text = "Обновить плановые цены"
    button.OnAction = "'" & ThisWorkbook.Name & "'!RefreshPlanning"
End Sub

Private Function ModeCode(ByVal value As String) As String
    Select Case LCase$(Trim$(value))
    Case "авиа", "air": ModeCode = "air"
    Case "поезд", "rail": ModeCode = "rail"
    Case "автобус", "bus": ModeCode = "bus"
    End Select
End Function

Public Function PlanStay(ByVal people As Variant, ByVal city As String, _
    ByVal stopType As String, ByVal nights As Variant, ByVal source As Range, _
    ByVal field As String) As Variant
    On Error GoTo Invalid
    Dim values As Variant, i As Long, price As Double, matches As Long, count As Double
    stopType = LCase$(Trim$(stopType))
    If stopType <> "пересадка" And stopType <> "пребывание" Then Err.Raise 5, , "Проживание: укажите пребывание или пересадка"
    If Not IsNumeric(people) Then Err.Raise 5, , "Проживание: укажите число человек"
    If people <= 0 Then Err.Raise 5, , "Проживание: число человек должно быть положительным"
    If Trim$(city) = "" Then Err.Raise 5, , "Проживание: укажите город"
    If CStr(nights) = "" Then
        If stopType = "пересадка" Then count = 0 Else Err.Raise 5, , "Проживание: укажите число ночей"
    Else
        If Not IsNumeric(nights) Then Err.Raise 5, , "Ночи должны быть целым числом от 0"
        count = CDbl(nights)
        If count < 0 Or count <> Fix(count) Then Err.Raise 5, , "Ночи должны быть целым числом от 0"
    End If
    If count = 0 Then
        If field = "status" Then PlanStay = "Ночей нет; проживание не начисляется" Else PlanStay = 0
        Exit Function
    End If
    values = source.Value2
    For i = 1 To UBound(values, 1)
        If StrComp(Trim$(CStr(values(i, 1))), Trim$(city), vbTextCompare) = 0 Then
            matches = matches + 1
            If IsNumeric(values(i, 8)) Then price = CDbl(values(i, 8))
        End If
    Next i
    If matches <> 1 Or price <= 0 Then Err.Raise 5, , "Проживание: нет однозначной справочной цены"
    If field = "status" Then
        PlanStay = "Текущая справочная цена проживания без индексации; " & count & " ночей/чел."
    Else
        PlanStay = Application.Round(people * count * price, 2)
    End If
    Exit Function
Invalid:
    If field = "status" Then PlanStay = Err.Description Else PlanStay = CVErr(xlErrNA)
End Function

Private Function CityKey(ByVal value As String) As String
    CityKey = LCase$(Trim$(Replace(Replace(value, "_", " "), "ё", "е")))
End Function

Private Function ModeTitle(ByVal code As String) As String
    ModeTitle = ""
    Select Case code
    Case "air": ModeTitle = "Авиа"
    Case "rail": ModeTitle = "Поезд"
    Case "bus": ModeTitle = "Автобус"
    End Select
End Function

Private Function AllowedRail(ByVal cls As String, ByVal firm As Boolean, ByVal settings As String) As Boolean
    Dim rules As Variant, rule As Variant, parts As Variant
    rules = Split(Replace(LCase$(settings), ";", ","), ",")
    For Each rule In rules
        parts = Split(Trim$(CStr(rule)), ":")
        If Trim$(parts(0)) = LCase$(cls) Then
            If UBound(parts) = 0 Then
                AllowedRail = True: Exit Function
            ElseIf UBound(parts) = 1 Then
                If Trim$(parts(1)) = "firm" And firm Then AllowedRail = True: Exit Function
            End If
        End If
    Next rule
End Function

Private Function CoverageReason(ByVal origin As String, ByVal destination As String, _
    ByVal mode As String, ByVal coverage As Range) As String
    If coverage Is Nothing Then Exit Function
    Dim values As Variant, i As Long
    values = coverage.Value2
    For i = 1 To UBound(values, 1)
        If CityKey(CStr(values(i, 1))) = origin And CityKey(CStr(values(i, 2))) = destination And values(i, 3) = mode Then
            Select Case CStr(values(i, 4))
            Case "no_direct_service": CoverageReason = "источник сообщает: прямого рейса нет"
            Case "no_priced_services": CoverageReason = "источник не опубликовал цену"
            Case "source_page_missing": CoverageReason = "страница направления у источника отсутствует (404)"
            Case "source_error": CoverageReason = "ошибка получения данных источника"
            Case "no_historical_reference": CoverageReason = "нет исторического ориентира по направлению"
            End Select
            Exit Function
        End If
    Next i
End Function

Public Function PlanFare(ByVal origin As String, ByVal destination As String, _
    ByVal launch As Variant, ByVal cohort As String, ByVal semester As String, _
    ByVal preferences As String, ByVal baggage As String, ByVal calendar As Range, _
    ByVal source As Range, ByVal field As String, Optional ByVal railClasses As String = "", _
    Optional ByVal coverage As Range = Nothing) As Variant
    On Error GoTo Invalid
    Dim values As Variant, cal As Variant, order As Variant, mode As Variant, allModes As Variant
    Dim seen As Object, prices As Object, methods As Object, routeRows As Collection, rowIndex As Variant
    Dim reasons As String, code As String, missing As String, label As String, reason As String
    Dim year As Long, start As Long, duration As Long, k As Long, m As Long, y As Long
    Dim row As Long, matches As Long, baseYear As Long, price As Double, total As Double, sample As Double
    Dim rate As Double, bag As Boolean, chosen As String, when As Date
    Dim eligible As Boolean, bad As Boolean, weighted As Double, weight As Double
    Dim knownOrigin As Boolean, knownDestination As Boolean
    origin = CityKey(origin): destination = CityKey(destination)
    If origin = "" Or destination = "" Then Err.Raise 5, , "Укажите откуда и куда"
    If origin = destination Then Err.Raise 5, , "Города участка совпадают"
    values = source.Value2
    Set routeRows = New Collection
    For row = 1 To UBound(values, 1)
        If CityKey(CStr(values(row, 1))) = origin Or CityKey(CStr(values(row, 2))) = origin Then knownOrigin = True
        If CityKey(CStr(values(row, 1))) = destination Or CityKey(CStr(values(row, 2))) = destination Then knownDestination = True
        If CityKey(CStr(values(row, 1))) = origin And CityKey(CStr(values(row, 2))) = destination Then routeRows.Add row
    Next row
    If Not knownOrigin Then Err.Raise 5, , "Город отправления вне покрытия: " & origin
    If Not knownDestination Then Err.Raise 5, , "Город прибытия вне покрытия: " & destination
    If Not IsNumeric(launch) Then Err.Raise 5, , "Укажите год запуска программы на листе Паспорт, G23"
    If CDbl(launch) = 0 Then Err.Raise 5, , "Укажите год запуска программы на листе Паспорт, G23"
    If CDbl(launch) <> Fix(CDbl(launch)) Or launch < 2026 Or launch > 9900 Then Err.Raise 5, , "Год запуска вне 2026..9900"
    If Val(cohort) < 1 Or Val(semester) < 1 Then Err.Raise 5, , "Нет семестра или года набора"
    cal = calendar.Value2
    For k = 1 To 4
        If Not IsNumeric(cal(k, 1)) Then Err.Raise 5, , "Некорректный календарь"
        If cal(k, 1) < 1 Or cal(k, 1) > 12 Or cal(k, 1) <> Fix(cal(k, 1)) Then Err.Raise 5, , "Календарь: нужны целые числа 1..12"
    Next k
    year = CLng(launch) + CLng(Val(cohort)) - 1 + (CLng(Val(semester)) - 1) \ 2
    If CLng(Val(semester)) Mod 2 = 1 Then
        start = cal(1, 1): duration = cal(2, 1)
    Else
        start = cal(3, 1): duration = cal(4, 1)
        If start < cal(1, 1) Then year = year + 1
    End If
    baggage = UCase$(Trim$(baggage))
    If baggage <> "ДА" And baggage <> "НЕТ" Then Err.Raise 5, , "Багаж авиа: укажите ДА или НЕТ"
    bag = (baggage = "ДА")
    Set seen = CreateObject("Scripting.Dictionary")
    Set prices = CreateObject("Scripting.Dictionary")
    Set methods = CreateObject("Scripting.Dictionary")
    order = Split(preferences, ",")
    For Each mode In order
        code = ModeCode(CStr(mode))
        If code = "" Then Err.Raise 5, , "Порядок: авиа, поезд, автобус через запятую"
        If seen.Exists(code) Then Err.Raise 5, , "В порядке транспорта есть повтор"
        seen.Add code, True
    Next mode
    If field = "air" Or field = "rail" Or field = "bus" Then
        allModes = Array(field)
    Else
        allModes = order
    End If
    For Each mode In allModes
        code = ModeCode(CStr(mode)): total = 0: missing = "": label = ""
        For k = 0 To duration - 1
            when = DateSerial(year, start + k, 1): m = Month(when): y = VBA.Year(when)
            matches = 0: weighted = 0: weight = 0: price = 0: bad = False
            For Each rowIndex In routeRows
                row = CLng(rowIndex)
                If values(row, 3) = code And values(row, 4) = m Then
                    eligible = True
                    If code = "air" Then eligible = (CBool(values(row, 8)) = bag)
                    If code = "rail" Then
                        eligible = False
                        If UBound(values, 2) >= 19 Then eligible = AllowedRail(CStr(values(row, 15)), CBool(values(row, 16)), railClasses)
                    End If
                    If eligible Then
                        matches = matches + 1
                        baseYear = CLng(values(row, 5)): rate = CDbl(values(row, 7))
                        sample = 1
                        If UBound(values, 2) >= 17 Then
                            If IsNumeric(values(row, 17)) Then sample = CDbl(values(row, 17))
                        End If
                        If y < baseYear Or rate <= -1 Or values(row, 6) <= 0 Or values(row, 9) = "" Or sample <= 0 Then
                            bad = True
                        Else
                            price = CDbl(values(row, 6)) * (1 + rate) ^ (y - baseYear)
                            weighted = weighted + price * sample: weight = weight + sample
                            If code = "air" Then
                                label = "АТОР: среднее контрольных минимумов; индекс маршрута 2025/26 повторяется; точность не оценена"
                            Else
                                label = "Резерв (" & ModeTitle(code) & "): среднее справочных цен; без сезонности; " & Format$(rate, "0.0%") & "/год по макросценарию ЦБ (не индекс транспорта); точность не оценена"
                                If y > 2029 Then label = label & "; после 2029 экстраполяция"
                            End If
                        End If
                    End If
                End If
            Next rowIndex
            If code <> "rail" And matches > 1 Then bad = True
            If matches = 0 Or bad Or weight <= 0 Then
                If missing <> "" Then missing = missing & ","
                missing = missing & Format$(when, "yyyy-mm")
            Else
                total = total + weighted / weight
            End If
        Next k
        If missing = "" Then
            prices.Add code, Application.Round(total / duration, 2)
            methods.Add code, label
        Else
            reason = CoverageReason(origin, destination, code, coverage)
            If reason = "" Then
                reason = "нет данных за " & missing
                If code = "rail" Then reason = reason & " для выбранных классов"
            End If
            reasons = reasons & ModeTitle(code) & ": " & reason & "; "
        End If
    Next mode
    For Each mode In order
        code = ModeCode(CStr(mode))
        If prices.Exists(code) Then chosen = code: Exit For
    Next mode
    If field = "status" Then
        If chosen <> "" Then PlanFare = methods(chosen) Else PlanFare = reasons
    ElseIf field = "mode" Then
        PlanFare = ModeTitle(chosen)
    ElseIf prices.Exists(field) Then
        PlanFare = prices(field)
    ElseIf field = "price" And chosen <> "" Then
        PlanFare = prices(chosen)
    Else
        PlanFare = CVErr(xlErrNA)
    End If
    Exit Function
Invalid:
    If field = "status" Then PlanFare = Err.Description Else PlanFare = CVErr(xlErrNA)
End Function


Public Function PlanRouteIssue(ByVal currentRow As Long, ByVal route As Range, _
    ByVal destinations As Range, ByVal groups As Range, ByVal categories As Range) As String
    PlanRouteIssue = ""
    On Error GoTo Invalid
    Dim legs As Object, data As Variant, dest As Variant, grp As Variant, cat As Variant
    Dim i As Long, index As Long, n As Double, item As Variant, previous As String, id As String
    data = route.Value2: dest = destinations.Value2: grp = groups.Value2: cat = categories.Value2
    index = currentRow - route.Row + 1: id = Trim$(CStr(data(index, 2)))
    If id = "" Then
        If data(index, 3) <> "" Then Err.Raise 5, , "Для номера участка укажите ID поездки"
        Exit Function
    End If
    Set legs = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, 2))) = id And grp(i, 1) = grp(index, 1) And grp(i, 2) = grp(index, 2) And InStr(cat(i, 1), "транспортные услуги") > 0 Then
            If Not IsNumeric(data(i, 3)) Then Err.Raise 5, , "Нужны номера участков от 1 без пропусков"
            n = CDbl(data(i, 3))
            If n < 1 Or n <> Fix(n) Then Err.Raise 5, , "Нужны целые номера участков от 1"
            If legs.Exists(CStr(n)) Then Err.Raise 5, , "Повтор номера участка"
            If grp(i, 3) <> grp(index, 3) Then Err.Raise 5, , "Количество поездок на участках различается: разделите ID"
            legs.Add CStr(n), Array(Trim$(CStr(data(i, 1))), Trim$(CStr(dest(i, 1))))
        End If
    Next i
    For i = 1 To legs.Count
        If Not legs.Exists(CStr(i)) Then Err.Raise 5, , "Номера участков должны идти от 1 без пропусков"
        item = legs(CStr(i))
        If item(0) = "" Or item(1) = "" Then Err.Raise 5, , "Укажите оба города каждого участка"
        If i > 1 Then
            If StrComp(previous, item(0), vbTextCompare) <> 0 Then Err.Raise 5, , "Разрыв маршрута между участками"
        End If
        previous = item(1)
    Next i
    Exit Function
Invalid:
    PlanRouteIssue = Err.Description
End Function
