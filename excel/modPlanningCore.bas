Attribute VB_Name = "modPlanningCore"
Option Explicit

' Domain calculations: arrays/scalars only; no worksheet, network or filesystem access.
Public Enum PlanningColumn
    pcOrigin = 1
    pcDestination = 2
    pcMode = 3
    pcMonth = 4
    pcBaseYear = 5
    pcPrice = 6
    pcAnnualRate = 7
    pcBaggage = 8
    pcSource = 9
    pcPreviousSource = 10
    pcReferenceDate = 11
    pcPreviousDate = 12
    pcRateMethod = 13
    pcQuoteId = 14
    pcRailClass = 15
    pcFirm = 16
    pcSampleCount = 17
    pcPriceMethod = 18
    pcUncertainty = 19
End Enum

Public Function PlanningSemesterMonths(ByVal launch As Variant, ByVal cohort As String, _
    ByVal semester As String, ByVal cal As Variant) As Variant
    Dim year As Long, start As Long, duration As Long, k As Long, months() As Date
    If Not IsNumeric(launch) Then Err.Raise 5, , "Укажите год запуска программы на листе Паспорт, G23"
    If CDbl(launch) = 0 Then Err.Raise 5, , "Укажите год запуска программы на листе Паспорт, G23"
    If CDbl(launch) <> Fix(CDbl(launch)) Or launch < 2026 Or launch > 9900 Then Err.Raise 5, , "Год запуска вне 2026..9900"
    If Val(cohort) < 1 Or Val(semester) < 1 Or Val(cohort) <> Fix(Val(cohort)) Or Val(semester) <> Fix(Val(semester)) Then Err.Raise 5, , "Нет семестра или года набора"
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
    ReDim months(0 To duration - 1)
    For k = 0 To duration - 1
        months(k) = DateSerial(year, start + k, 1)
    Next k
    PlanningSemesterMonths = months
End Function

Private Function PlanningRound(ByVal value As Double, ByVal digits As Long) As Double
    ' Excel-style rounding, without Excel objects or VBA's banker's rounding.
    Dim decimalFactor As Double
    decimalFactor = 10 ^ digits
    PlanningRound = Sgn(value) * CDbl(Fix(CDec(Abs(value)) * CDec(decimalFactor) + CDec(0.5))) / decimalFactor
End Function

Private Function PlanningModeCode(ByVal value As String) As String
    Select Case LCase$(Trim$(value))
    Case "авиа", "air": PlanningModeCode = "air"
    Case "поезд", "rail": PlanningModeCode = "rail"
    Case "автобус", "bus": PlanningModeCode = "bus"
    End Select
End Function

Public Function PlanningStay(ByVal people As Variant, ByVal city As String, _
    ByVal stopType As String, ByVal nights As Variant, ByVal values As Variant, _
    ByVal field As String) As Variant
    Dim i As Long, price As Double, matches As Long, count As Double
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
        If field = "status" Then PlanningStay = "Ночей нет; проживание не начисляется" Else PlanningStay = 0
        Exit Function
    End If
    For i = 1 To UBound(values, 1)
        If StrComp(Trim$(CStr(values(i, 1))), Trim$(city), vbTextCompare) = 0 Then
            matches = matches + 1
            If IsNumeric(values(i, 8)) Then price = CDbl(values(i, 8))
        End If
    Next i
    If matches <> 1 Or price <= 0 Then Err.Raise 5, , "Проживание: нет однозначной справочной цены"
    If field = "status" Then
        PlanningStay = "Текущая справочная цена проживания без индексации; " & count & " ночей/чел."
    Else
        PlanningStay = PlanningRound(people * count * price, 2)
    End If
    Exit Function
End Function

Private Function PlanningCityKey(ByVal value As String) As String
    PlanningCityKey = LCase$(Trim$(Replace(Replace(value, "_", " "), "ё", "е")))
End Function

Public Function PlanningModeTitle(ByVal code As String) As String
    PlanningModeTitle = ""
    Select Case code
    Case "air": PlanningModeTitle = "Авиа"
    Case "rail": PlanningModeTitle = "Поезд"
    Case "bus": PlanningModeTitle = "Автобус"
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
    ByVal mode As String, ByVal values As Variant) As String
    If IsEmpty(values) Then Exit Function
    Dim i As Long
    For i = 1 To UBound(values, 1)
        If PlanningCityKey(CStr(values(i, 1))) = origin And PlanningCityKey(CStr(values(i, 2))) = destination And values(i, 3) = mode Then
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

Public Function PlanningEstimate(ByVal origin As String, ByVal destination As String, _
    ByVal launch As Variant, ByVal cohort As String, ByVal semester As String, _
    ByVal preferences As String, ByVal baggage As String, ByVal cal As Variant, _
    ByVal values As Variant, Optional ByVal railClasses As String = "", _
    Optional ByVal coverageValues As Variant = Empty, Optional ByVal requestedMode As String = "") As Object
    Dim order As Variant, mode As Variant, allModes As Variant, months As Variant, period As Variant
    Dim seen As Object, prices As Object, methods As Object, routeRows As Collection, rowIndex As Variant
    Dim reasons As String, code As String, missing As String, label As String, reason As String
    Dim duration As Long, m As Long, y As Long
    Dim row As Long, matches As Long, baseYear As Long, price As Double, total As Double, sample As Double
    Dim rate As Double, bag As Boolean, chosen As String, when As Date
    Dim eligible As Boolean, bad As Boolean, weighted As Double, weight As Double
    Dim knownOrigin As Boolean, knownDestination As Boolean
    origin = PlanningCityKey(origin): destination = PlanningCityKey(destination)
    If origin = "" Or destination = "" Then Err.Raise 5, , "Укажите откуда и куда"
    If origin = destination Then Err.Raise 5, , "Города участка совпадают"
    Set routeRows = New Collection
    For row = 1 To UBound(values, 1)
        If PlanningCityKey(CStr(values(row, pcOrigin))) = origin Or PlanningCityKey(CStr(values(row, pcDestination))) = origin Then knownOrigin = True
        If PlanningCityKey(CStr(values(row, pcOrigin))) = destination Or PlanningCityKey(CStr(values(row, pcDestination))) = destination Then knownDestination = True
        If PlanningCityKey(CStr(values(row, pcOrigin))) = origin And PlanningCityKey(CStr(values(row, pcDestination))) = destination Then routeRows.Add row
    Next row
    If Not knownOrigin Then Err.Raise 5, , "Город отправления вне покрытия: " & origin
    If Not knownDestination Then Err.Raise 5, , "Город прибытия вне покрытия: " & destination
    months = PlanningSemesterMonths(launch, cohort, semester, cal)
    duration = UBound(months) + 1
    baggage = UCase$(Trim$(baggage))
    If baggage <> "ДА" And baggage <> "НЕТ" Then Err.Raise 5, , "Багаж авиа: укажите ДА или НЕТ"
    bag = (baggage = "ДА")
    Set seen = CreateObject("Scripting.Dictionary")
    Set prices = CreateObject("Scripting.Dictionary")
    Set methods = CreateObject("Scripting.Dictionary")
    order = Split(preferences, ",")
    For Each mode In order
        code = PlanningModeCode(CStr(mode))
        If code = "" Then Err.Raise 5, , "Порядок: авиа, поезд, автобус через запятую"
        If seen.Exists(code) Then Err.Raise 5, , "В порядке транспорта есть повтор"
        seen.Add code, True
    Next mode
    If requestedMode <> "" Then
        allModes = Array(requestedMode)
    Else
        allModes = order
    End If
    For Each mode In allModes
        code = PlanningModeCode(CStr(mode)): total = 0: missing = "": label = ""
        For Each period In months
            when = CDate(period): m = Month(when): y = VBA.Year(when)
            matches = 0: weighted = 0: weight = 0: price = 0: bad = False
            For Each rowIndex In routeRows
                row = CLng(rowIndex)
                If values(row, pcMode) = code And values(row, pcMonth) = m Then
                    eligible = True
                    If code = "air" Then eligible = (CBool(values(row, pcBaggage)) = bag)
                    If code = "rail" Then
                        eligible = False
                        If UBound(values, 2) >= 19 Then eligible = AllowedRail(CStr(values(row, pcRailClass)), CBool(values(row, pcFirm)), railClasses)
                    End If
                    If eligible Then
                        matches = matches + 1
                        baseYear = CLng(values(row, pcBaseYear)): rate = CDbl(values(row, pcAnnualRate))
                        sample = 1
                        If UBound(values, 2) >= 17 Then
                            If IsNumeric(values(row, pcSampleCount)) Then sample = CDbl(values(row, pcSampleCount))
                        End If
                        If y < baseYear Or rate <= -1 Or values(row, pcPrice) <= 0 Or values(row, pcSource) = "" Or sample <= 0 Then
                            bad = True
                        Else
                            price = CDbl(values(row, pcPrice)) * (1 + rate) ^ (y - baseYear)
                            weighted = weighted + price * sample: weight = weight + sample
                            If code = "air" Then
                                label = "АТОР: среднее контрольных минимумов; индекс маршрута 2025/26 повторяется; точность не оценена"
                            Else
                                label = "Резерв (" & PlanningModeTitle(code) & "): среднее справочных цен; без сезонности; " & Format$(rate, "0.0%") & "/год по макросценарию ЦБ (не индекс транспорта); точность не оценена"
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
        Next period
        If missing = "" Then
            prices.Add code, PlanningRound(total / duration, 2)
            methods.Add code, label
        Else
            reason = CoverageReason(origin, destination, code, coverageValues)
            If reason = "" Then
                reason = "нет данных за " & missing
                If code = "rail" Then reason = reason & " для выбранных классов"
            End If
            reasons = reasons & PlanningModeTitle(code) & ": " & reason & "; "
        End If
    Next mode
    For Each mode In order
        code = PlanningModeCode(CStr(mode))
        If prices.Exists(code) Then chosen = code: Exit For
    Next mode
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.Add "prices", prices: result.Add "methods", methods
    result.Add "chosen", chosen: result.Add "reasons", reasons
    Set PlanningEstimate = result
End Function

Public Function PlanningRouteIssue(ByVal index As Long, ByVal data As Variant, _
    ByVal dest As Variant, ByVal grp As Variant, ByVal cat As Variant) As String
    PlanningRouteIssue = ""
    On Error GoTo Invalid
    Dim legs As Object
    Dim i As Long, n As Double, item As Variant, previous As String, id As String
    id = Trim$(CStr(data(index, 2)))
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
    PlanningRouteIssue = Err.Description
End Function
