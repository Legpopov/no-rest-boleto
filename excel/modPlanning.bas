Attribute VB_Name = "modPlanning"
Option Explicit

' Offline semester forecast. No network access from worksheet functions.
' Reference fares are one-way RUB; annual reference changes repeat as a scenario.
Private Const TRIPS As String = "Командировки и выезды"
Private Const PARAMS As String = "Параметры цен"
Private Const DATA As String = "Данные плана"

Public Sub RefreshPlanning(Optional ByVal source As String = "")
    On Error GoTo Failed
    Dim req As Object, root As Object
    Dim text As String, stream As Object, matrix As Variant, j As Long
    Dim ws As Worksheet, ps As Worksheet, fields As Variant, stamp As String
    Dim coverageMatrix As Variant, coverageFields As Variant, hotelMatrix As Variant, cityMatrix() As Variant, cities As Variant, n As Long
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
    If root("schema_version") <> 3 Or root("currency") <> "RUB" Or root("basis") <> "one_way" Then
        Err.Raise 5, , "Неподдерживаемая схема плановых цен"
    End If
    fields = PlanningCatalogFields()
    coverageFields = PlanningCoverageFields()
    matrix = PlanningCatalogMatrix(root)
    coverageMatrix = PlanningCoverageMatrix(root)
    hotelMatrix = PlanningHotelMatrix(root)
    cities = root("cities")
    ReDim cityMatrix(1 To UBound(cities) + 1, 1 To 2)
    For n = 0 To UBound(cities)
        cityMatrix(n + 1, 1) = CStr(cities(n))
        cityMatrix(n + 1, 2) = ""
    Next n
    If root.Exists("city_countries") Then
        Dim countries As Object
        Set countries = root("city_countries")
        For n = 1 To UBound(cityMatrix, 1)
            If countries.Exists(CStr(cityMatrix(n, 1))) Then cityMatrix(n, 2) = CStr(countries(CStr(cityMatrix(n, 1))))
        Next n
    End If
    stamp = PlanningText(root, "retrieved_at", True)
    CreatePlanningBackup
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(DATA)
    On Error GoTo Failed
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = DATA
    End If
    ws.Cells.ClearContents
    For j = 0 To UBound(fields)
        ws.Cells(1, j + 1).Value2 = fields(j)
    Next j
    WritePlanningValues ws.Range("A2").Resize(UBound(matrix, 1), pcFieldCount), matrix
    ws.Range("D2:G" & UBound(matrix, 1) + 1).NumberFormat = "General"
    ws.Range("P2:Q" & UBound(matrix, 1) + 1).NumberFormat = "General"
    For j = 0 To UBound(coverageFields)
        ws.Cells(1, j + 21).Value2 = coverageFields(j)
    Next j
    WritePlanningValues ws.Range("U2").Resize(UBound(coverageMatrix, 1), 6), coverageMatrix
    ws.Range("F2:F" & UBound(matrix, 1) + 1).NumberFormat = "#,##0.00"
    ws.Range("G2:G" & UBound(matrix, 1) + 1).NumberFormat = "0.00%"
    ws.Rows(1).Font.Bold = True
    ws.Columns("A:N").ColumnWidth = 20
    ws.Columns("I:J").ColumnWidth = 45
    ThisWorkbook.Names.Add Name:="PlanningData", RefersTo:="='" & DATA & "'!$A$2:$T$" & UBound(matrix, 1) + 1
    ThisWorkbook.Names.Add Name:="PlanningCoverage", RefersTo:="='" & DATA & "'!$U$2:$Z$" & UBound(coverageMatrix, 1) + 1
    WritePlanningValues ws.Range("AL2").Resize(UBound(cityMatrix, 1), 2), cityMatrix
    ws.Range("AL1").Value2 = "city": ws.Range("AM1").Value2 = "country_code"
    ThisWorkbook.Names.Add Name:="PlanningCities", RefersTo:="='" & DATA & "'!$AL$2:$AM$" & UBound(cityMatrix, 1) + 1
    ThisWorkbook.Names.Add Name:="PlanningDestinations", RefersTo:="='" & DATA & "'!$AL$2:$AL$" & UBound(cityMatrix, 1) + 1
    If Not IsEmpty(hotelMatrix) Then
        Dim hotelHeaders As Variant
        hotelHeaders = Array("city", "filtered_price", "base_year", "base_month", "annual_rate", "price_source", "rate_source", "retrieved_at", "rate_method", "status")
        For j = 0 To UBound(hotelHeaders): ws.Cells(1, j + 27).Value2 = hotelHeaders(j): Next j
        WritePlanningValues ws.Range("AA2").Resize(UBound(hotelMatrix, 1), 10), hotelMatrix
        ws.Range("AB2:AE" & UBound(hotelMatrix, 1) + 1).NumberFormat = "General"
        ThisWorkbook.Names.Add Name:="PlanningHotels", RefersTo:="='" & DATA & "'!$AA$2:$AJ$" & UBound(hotelMatrix, 1) + 1
    End If
    WritePlanningValues ps.Range("O8"), stamp
    ps.Range("O9").Value2 = "Авиа: текущие односторонние минимумы / АТОР; индексы категорий РФ как сценарии, за рубежом макросценарий RUB. ЖД/автобус/отели без сезонности. Точность не оценена."
    Application.CalculateFull
    PrunePlanningBackups
    Exit Sub
Failed:
    Err.Raise Err.Number, "RefreshPlanning", Err.Description
End Sub

Public Sub InstallPlanning()
    Dim ws As Worksheet, ps As Worksheet, tbl As ListObject, headers As Variant
    Dim r As Long, c As Long, args As String, emptyRow As String, guard As String
    Dim railCell As Range, extra As String, bagRU As Range, bagAB As Range
    Set ws = ThisWorkbook.Worksheets(TRIPS)
    Set ps = ThisWorkbook.Worksheets(PARAMS)
    Set tbl = ws.ListObjects("Таблица3")
    Set railCell = ps.Columns(1).Find(What:="ЖД: допустимые классы", LookIn:=xlValues, LookAt:=xlWhole)
    If railCell Is Nothing Then Err.Raise 5, , "Нет параметра ЖД: допустимые классы"
    Set bagRU = ps.Columns(1).Find(What:="Авиа: багаж РФ, руб. за перелёт в одну сторону", LookIn:=xlValues, LookAt:=xlWhole)
    Set bagAB = ps.Columns(1).Find(What:="Авиа: багаж зарубеж, руб. за перелёт в одну сторону", LookIn:=xlValues, LookAt:=xlWhole)
    If bagRU Is Nothing Or bagAB Is Nothing Then Err.Raise 5, , "Нет параметров надбавки за багаж"
    extra = ",'Параметры цен'!$B$" & railCell.Row & ",PlanningCoverage,PlanningCities," & _
        "IF('Параметры цен'!$C$" & bagRU.Row & "=""ДА"",'Параметры цен'!$B$" & bagRU.Row & ",-1)," & _
        "IF('Параметры цен'!$C$" & bagAB.Row & "=""ДА"",'Параметры цен'!$B$" & bagAB.Row & ",-1)"
    headers = Array("Откуда", "Поездка ID", "Участок", "Остановка", "Ночей", _
                    "Порядок транспорта", "Багаж авиа", "Авиа, руб./чел.", "Поезд, руб./чел.", _
                    "Автобус, руб./чел.", "Выбранный транспорт", "Метод / причина")
    CreatePlanningBackup
    tbl.Resize ws.Range("A2:W" & tbl.Range.Row + tbl.Range.Rows.Count - 1)
    For c = 0 To UBound(headers)
        ws.Cells(2, c + 12).Value2 = headers(c)
    Next c
    ps.Range("N2").Value2 = "Источник плановых цен"
    If ps.Range("O2").Value2 = "" Then ps.Range("O2").Value2 = "https://raw.githubusercontent.com/Legpopov/no-rest-boleto/main/data/planning-excel.json"
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
            If ws.Rows(r).RowHeight < 90 Then ws.Rows(r).RowHeight = 90
        ElseIf InStr(1, ws.Cells(r, 8).Value2, "оплата проживания", vbTextCompare) > 0 Then
            emptyRow = "AND(C" & r & "="""",F" & r & "="""",O" & r & "="""",P" & r & "="""")"
            args = "C" & r & ",F" & r & ",O" & r & ",P" & r & ",'Параметры'!$C$2,B" & r & ",A" & r & ",'Параметры цен'!$O$3:$O$6,PlanningHotels,"
            ws.Cells(r, 10).Formula = "=IF(" & emptyRow & ","""",PlanStayForecast(" & args & """price""))"
            ws.Cells(r, 23).Formula = "=IF(" & emptyRow & ","""",PlanStayForecast(" & args & """status""))"
            If ws.Rows(r).RowHeight < 75 Then ws.Rows(r).RowHeight = 75
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
    For r = 3 To tbl.Range.Row + tbl.Range.Rows.Count - 1
        If InStr(1, ws.Cells(r, 8).Value2, "транспортные услуги", vbTextCompare) > 0 Or InStr(1, ws.Cells(r, 8).Value2, "оплата проживания", vbTextCompare) > 0 Then
            With ws.Cells(r, 6).Validation
                .Delete
                .Add xlValidateList, xlValidAlertStop, xlBetween, "=PlanningDestinations"
                .ShowError = False
            End With
        End If
    Next r
    Dim button As Shape
    On Error Resume Next
    ps.Shapes("RefreshPlanButton").Delete
    On Error GoTo 0
    Set button = ps.Shapes.AddFormControl(xlButtonControl, ps.Range("N11").Left, ps.Range("N11").Top, 200, 25)
    button.Name = "RefreshPlanButton"
    button.TextFrame.Characters.Text = "Обновить плановые цены"
    button.OnAction = "'" & ThisWorkbook.Name & "'!RefreshPlanning"
    PrunePlanningBackups
End Sub

Public Function PlanFare(ByVal origin As String, ByVal destination As String, _
    ByVal launch As Variant, ByVal cohort As String, ByVal semester As String, _
    ByVal preferences As String, ByVal baggage As String, ByVal calendar As Range, _
    ByVal source As Range, ByVal field As String, Optional ByVal railClasses As String = "", _
    Optional ByVal coverage As Range = Nothing, Optional ByVal cityRegistry As Range = Nothing, _
    Optional ByVal allowanceRU As Double = -1, Optional ByVal allowanceAB As Double = -1) As Variant
    On Error GoTo Invalid
    Dim result As Object, prices As Object, methods As Object, chosen As String, coverageValues As Variant, requestedMode As String
    If Not coverage Is Nothing Then coverageValues = coverage.Value2
    If field = "air" Or field = "rail" Or field = "bus" Then requestedMode = field
    Dim allowance As Double, countries As Variant, i As Long, originCountry As String, destinationCountry As String
    allowance = -1
    If Not cityRegistry Is Nothing Then
        countries = cityRegistry.Value2
        For i = 1 To UBound(countries, 1)
            If LCase$(Replace(CStr(countries(i, 1)), "ё", "е")) = LCase$(Replace(Trim$(origin), "ё", "е")) Then originCountry = CStr(countries(i, 2))
            If LCase$(Replace(CStr(countries(i, 1)), "ё", "е")) = LCase$(Replace(Trim$(destination), "ё", "е")) Then destinationCountry = CStr(countries(i, 2))
        Next i
        If originCountry <> "" And destinationCountry <> "" Then
            If originCountry = "RU" And destinationCountry = "RU" Then allowance = allowanceRU Else allowance = allowanceAB
        End If
    End If
    Set result = PlanningEstimate(origin, destination, launch, cohort, semester, _
                                  preferences, baggage, calendar.Value2, source.Value2, railClasses, coverageValues, requestedMode, allowance, countries)
    Set prices = result("prices"): Set methods = result("methods"): chosen = result("chosen")
    If field = "status" Then
        If chosen <> "" Then PlanFare = methods(chosen) Else PlanFare = result("reasons")
    ElseIf field = "mode" Then
        PlanFare = PlanningModeTitle(chosen)
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

Public Function PlanStay(ByVal people As Variant, ByVal city As String, _
    ByVal stopType As String, ByVal nights As Variant, ByVal source As Range, _
    ByVal field As String) As Variant
    On Error GoTo Invalid
    PlanStay = PlanningStay(people, city, stopType, nights, source.Value2, field)
    Exit Function
Invalid:
    If field = "status" Then PlanStay = Err.Description Else PlanStay = CVErr(xlErrNA)
End Function

Public Function PlanRouteIssue(ByVal currentRow As Long, ByVal route As Range, _
    ByVal destinations As Range, ByVal groups As Range, ByVal categories As Range) As String
    On Error GoTo Invalid
    PlanRouteIssue = PlanningRouteIssue(currentRow - route.Row + 1, route.Value2, _
                                       destinations.Value2, groups.Value2, categories.Value2)
    Exit Function
Invalid:
    PlanRouteIssue = Err.Description
End Function

Public Function PlanStayForecast(ByVal people As Variant, ByVal city As String, _
    ByVal stopType As String, ByVal nights As Variant, ByVal launch As Variant, _
    ByVal cohort As String, ByVal semester As String, ByVal calendar As Range, _
    ByVal source As Range, ByVal field As String) As Variant
    On Error GoTo Invalid
    PlanStayForecast = PlanningIndexedStay(people, city, stopType, nights, launch, cohort, semester, calendar.Value2, source.Value2, field)
    Exit Function
Invalid:
    If field = "status" Then PlanStayForecast = Err.Description Else PlanStayForecast = CVErr(xlErrNA)
End Function
