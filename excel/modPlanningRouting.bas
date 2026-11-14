Attribute VB_Name = "modPlanningRouting"
Option Explicit

Private mRoutingGraphs As Object
Private mRoutingEvents As clsPlanningRoutingEvents

Public Sub RoutingClearCache()
    Set mRoutingGraphs = Nothing
End Sub

Private Function RoutingCachedGraph(ByVal key As String) As Object
    If Not Application.EnableEvents Or Application.CalculationState <> xlCalculating Then RoutingClearCache: Exit Function
    If key = "" Then Exit Function
    If mRoutingEvents Is Nothing Then
        Set mRoutingEvents = New clsPlanningRoutingEvents
        Set mRoutingEvents.App = Application
    End If
    If Not mRoutingGraphs Is Nothing Then
        If mRoutingGraphs.Exists(key) Then Set RoutingCachedGraph = mRoutingGraphs(key)
    End If
End Function

Private Sub RoutingStoreGraph(ByVal key As String, ByVal value As Object)
    If key = "" Or Not Application.EnableEvents Or Application.CalculationState <> xlCalculating Then Exit Sub
    If mRoutingGraphs Is Nothing Then Set mRoutingGraphs = CreateObject("Scripting.Dictionary")
    If mRoutingGraphs.Exists(key) Then mRoutingGraphs.Remove key
    mRoutingGraphs.Add key, value
End Sub


' Offline graph search. Each accepted edge prices every semester month.
' No timetable connection or unknown transfer expense is declared confirmed.

Public Function RoutingTransfersMatrix(ByVal root As Object, ByVal values As Variant) As Variant
    Dim counts As Variant, hasCounts As Boolean, entries As Object, row As Long
    Dim id As String, count As Variant, number As Double, out() As Variant, key As Variant, n As Long
    hasCounts = root.Exists("transfer_counts")
    If hasCounts Then
        counts = root("transfer_counts")
        If UBound(counts) + 1 <> UBound(values, 1) Then Err.Raise 5, , "Transfer metadata has a different row count"
    End If
    Set entries = CreateObject("Scripting.Dictionary")
    For row = 1 To UBound(values, 1)
        id = CStr(values(row, pcQuoteId))
        number = 0
        If values(row, pcMode) = "air" Then number = -1
        If hasCounts Then
            count = counts(row - 1)
            If Not IsNull(count) Then
                Select Case VarType(count)
                Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
                    number = CDbl(count)
                    If number < 0 Or number <> Fix(number) Then Err.Raise 5, , "Invalid transfer count"
                Case Else
                    Err.Raise 5, , "Transfer count must be numeric or null"
                End Select
            Else
                number = -1
            End If
        End If
        If id <> "" Then
            If entries.Exists(id) Then
                If entries(id) <> number Then Err.Raise 5, , "Conflicting transfer counts for one quote"
            Else
                entries.Add id, number
            End If
        End If
    Next row
    If entries.Count = 0 Then
        ReDim out(1 To 1, 1 To 2)
        out(1, 1) = "": out(1, 2) = -1
        RoutingTransfersMatrix = out
        Exit Function
    End If
    ReDim out(1 To entries.Count, 1 To 2)
    For Each key In entries.Keys
        n = n + 1: out(n, 1) = CStr(key): out(n, 2) = entries(key)
    Next key
    RoutingTransfersMatrix = out
End Function

Private Function RoutingRanks(ByVal preferences As String) As Object
    Dim result As Object, item As Variant, code As String
    Set result = CreateObject("Scripting.Dictionary")
    For Each item In Split(preferences, ",")
        code = PlanningModeCode(CStr(item))
        If code = "" Then Err.Raise 5, , "Неизвестный вид транспорта в предпочтениях"
        If result.Exists(code) Then Err.Raise 5, , "В предпочтениях повторяется вид транспорта"
        result.Add code, result.Count
    Next item
    Set RoutingRanks = result
End Function

Public Function RoutingFind(ByVal edges As Collection, ByVal origin As String, ByVal destination As String, _
    ByVal preferences As String, Optional ByVal limit As Long = 2) As Object
    Dim graph As Object, ranks As Object, visited As Object, best As Object
    Dim path As New Collection, edge As Object, outgoing As Collection, count As Variant
    If origin = "" Or destination = "" Or origin = destination Then Err.Raise 5, , "Укажите два разных города"
    If limit < 0 Or limit > 2 Then Err.Raise 5, , "Допустимо не больше двух пересадок"
    Set graph = CreateObject("Scripting.Dictionary")
    Set ranks = RoutingRanks(preferences)
    For Each edge In edges
        If ranks.Exists(CStr(edge("mode"))) And CDbl(edge("price")) > 0 Then
            If Not edge.Exists("transfers") Then edge.Add "transfers", Null
            count = edge("transfers")
            If IsNull(count) Then
                edge("transfers") = 0
                If edge.Exists("unknown") Then edge("unknown") = True Else edge.Add "unknown", True
            Else
                Select Case VarType(count)
                Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
                    If CDbl(count) < 0 Or CDbl(count) <> Fix(CDbl(count)) Then Err.Raise 5, , "Некорректное число пересадок"
                Case Else
                    Err.Raise 5, , "Некорректный тип числа пересадок"
                End Select
                If Not edge.Exists("unknown") Then edge.Add "unknown", False
            End If
            If Not graph.Exists(CStr(edge("origin"))) Then
                Set outgoing = New Collection
                graph.Add CStr(edge("origin")), outgoing
            End If
            Set outgoing = graph(CStr(edge("origin")))
            outgoing.Add edge
        End If
    Next edge
    Set best = CreateObject("Scripting.Dictionary")
    best.Add "status", "insufficient_data": best.Add "price", Null
    Set visited = CreateObject("Scripting.Dictionary"): visited.Add origin, True
    RoutingVisit graph, origin, destination, ranks, limit, path, visited, 0, False, -1, 0, best
    Set RoutingFind = best
End Function

Private Function RoutingPreferenceKey(ByVal path As Collection, ByVal ranks As Object) As String
    Dim used As Object, edge As Object, code As Variant, rank As Long
    Set used = CreateObject("Scripting.Dictionary")
    For Each edge In path: used(CStr(edge("mode"))) = True: Next edge
    For rank = ranks.Count - 1 To 1 Step -1
        For Each code In ranks.Keys
            If ranks(code) = rank Then
                If used.Exists(CStr(code)) Then
                    RoutingPreferenceKey = RoutingPreferenceKey & "1"
                Else
                    RoutingPreferenceKey = RoutingPreferenceKey & "0"
                End If
            End If
        Next code
    Next rank
End Function

Private Sub RoutingVisit(ByVal graph As Object, ByVal city As String, ByVal destination As String, _
    ByVal ranks As Object, ByVal limit As Long, ByVal path As Collection, ByVal visited As Object, _
    ByVal known As Long, ByVal unknown As Boolean, ByVal worst As Long, ByVal amount As Double, ByVal best As Object)
    Dim edge As Object, target As String, count As Long, rank As Long, price As Double
    Dim uncertain As Boolean, better As Boolean, copied As Collection, item As Object, signature As String, penalty As String
    If Not graph.Exists(city) Then Exit Sub
    For Each edge In graph(city)
        target = CStr(edge("destination"))
        If Not visited.Exists(target) Then
            count = known + CLng(edge("transfers"))
            If path.Count > 0 Then count = count + 1
            If count <= limit Then
                rank = ranks(CStr(edge("mode")))
                If rank < worst Then rank = worst
                If best("status") <> "ok" Then
                    better = True
                Else
                    better = rank <= best("rank")
                End If
                If better Then
                    price = PlanningRound(amount + CDbl(edge("price")), 2)
                    uncertain = unknown Or CBool(edge("unknown"))
                    path.Add edge
                    penalty = RoutingPreferenceKey(path, ranks)
                    better = True
                    If best("status") = "ok" Then
                        If StrComp(penalty, CStr(best("penalty")), vbBinaryCompare) > 0 Then better = False
                        If penalty = best("penalty") Then
                            If count > best("known_transfers") Then better = False
                            If count = best("known_transfers") And price > best("price") Then better = False
                        End If
                    End If
                    If Not better Then
                        path.Remove path.Count
                        GoTo NextRoutingEdge
                    End If
                    If target = destination Then
                        signature = ""
                        For Each item In path
                            signature = signature & CStr(item("origin")) & vbTab & CStr(item("destination")) & vbTab & CStr(item("mode")) & vbLf
                        Next item
                        better = True
                        If best("status") = "ok" Then
                            If penalty <> best("penalty") Then
                                better = StrComp(penalty, CStr(best("penalty")), vbBinaryCompare) < 0
                            ElseIf count <> best("known_transfers") Then
                                better = count < best("known_transfers")
                            ElseIf price <> best("price") Then
                                better = price < best("price")
                            Else
                                better = StrComp(signature, CStr(best("signature")), vbBinaryCompare) < 0
                            End If
                        End If
                        If better Then
                            Set copied = New Collection
                            For Each item In path: copied.Add item: Next item
                            best("status") = "ok": best("price") = price
                            best("rank") = rank: best("penalty") = penalty: best("known_transfers") = count
                            best("transfer_limit_verified") = Not uncertain
                            best("signature") = signature
                            If best.Exists("legs") Then best.Remove "legs"
                            best.Add "legs", copied
                        End If
                    ElseIf path.Count <= limit Then
                        visited.Add target, True
                        RoutingVisit graph, target, destination, ranks, limit, path, visited, count, uncertain, rank, price, best
                        visited.Remove target
                    End If
                    path.Remove path.Count
                End If
            End If
        End If
NextRoutingEdge:
    Next edge
End Sub

Private Sub RoutingAirCount(ByVal rows As Variant, ByVal months As Variant, ByVal bag As Boolean, _
    ByVal allowance As Double, ByVal counts As Object, ByRef known As Long, ByRef unknown As Boolean)
    Dim period As Variant, row As Long, priority As Long, best As Long, current As Boolean
    Dim eligible As Boolean, number As Long, id As String, monthKnown As Long, monthUnknown As Boolean
    For Each period In months
        best = 0: monthKnown = 0: monthUnknown = False
        For row = 1 To UBound(rows, 1)
            If rows(row, pcMode) = "air" And rows(row, pcMonth) = Month(CDate(period)) And rows(row, pcBaseYear) <= Year(CDate(period)) Then
                current = Left$(CStr(rows(row, pcRateMethod)), 12) = "air_current_"
                If current Then eligible = (Not bag Or allowance >= 0) Else eligible = (CBool(rows(row, pcBaggage)) = bag)
                If eligible Then
                    priority = 1: If current Then priority = 2
                    If priority > best Then best = priority: monthKnown = 0: monthUnknown = False
                    If priority = best Then
                        number = -1: id = CStr(rows(row, pcQuoteId))
                        If counts.Exists(id) Then number = CLng(counts(id))
                        If number < 0 Then
                            monthUnknown = True
                        ElseIf number > monthKnown Then
                            monthKnown = number
                        End If
                    End If
                End If
            End If
        Next row
        If monthKnown > known Then known = monthKnown
        unknown = unknown Or monthUnknown
    Next period
End Sub

Public Function RoutingEstimate(ByVal origin As String, ByVal destination As String, ByVal launch As Variant, _
    ByVal cohort As String, ByVal semester As String, ByVal preferences As String, ByVal baggage As String, _
    ByVal cal As Variant, ByVal values As Variant, ByVal railClasses As String, ByVal countries As Variant, _
    ByVal allowanceRU As Double, ByVal allowanceAB As Double, ByVal transferValues As Variant, _
    Optional ByVal requestedMode As String = "", Optional ByVal cachePrefix As String = "") As Object
    Dim origins As Object, cities As Object, countryMap As Object, groups As Object, counts As Object, ranks As Object
    Dim row As Long, a As String, b As String, key As Variant, indices As Collection, n As Long, j As Long, index As Variant
    Dim subset() As Variant, months As Variant, estimates As Object, prices As Object, methods As Object, mode As Variant
    Dim edges As Collection, edge As Object, chosen As Object, result As Object, allowance As Double
    Dim graphKey As String, graphEntry As Object, calendarRow As Long
    Dim known As Long, unknown As Boolean, chosenCode As String, title As String, description As String, path As Collection
    origin = PlanningCityKey(origin): destination = PlanningCityKey(destination)
    If origin = "" Or destination = "" Or origin = destination Then Err.Raise 5, , "Укажите два разных города"
    Set ranks = RoutingRanks(preferences)
    months = PlanningSemesterMonths(launch, cohort, semester, cal)
    If cachePrefix <> "" Then
        graphKey = cachePrefix & "|" & destination & "|" & CStr(launch) & "|" & cohort & "|" & semester & "|" & baggage & "|" & railClasses & "|" & CStr(allowanceRU) & "|" & CStr(allowanceAB)
        For calendarRow = 1 To UBound(cal, 1): graphKey = graphKey & "|" & CStr(cal(calendarRow, 1)): Next calendarRow
    End If
    Set graphEntry = RoutingCachedGraph(graphKey)
    If graphEntry Is Nothing Then
        Set edges = New Collection
    Set origins = CreateObject("Scripting.Dictionary"): Set cities = CreateObject("Scripting.Dictionary")
    Set countryMap = CreateObject("Scripting.Dictionary"): Set groups = CreateObject("Scripting.Dictionary")
    Set counts = CreateObject("Scripting.Dictionary")
    If Not IsEmpty(countries) Then
        For row = 1 To UBound(countries, 1)
            a = PlanningCityKey(CStr(countries(row, 1)))
            cities(a) = True: countryMap(a) = CStr(countries(row, 2))
        Next row
    End If
    If Not IsEmpty(transferValues) Then
        For row = 1 To UBound(transferValues, 1)
            counts(CStr(transferValues(row, 1))) = CLng(transferValues(row, 2))
        Next row
    End If
    For row = 1 To UBound(values, 1)
        a = PlanningCityKey(CStr(values(row, pcOrigin))): b = PlanningCityKey(CStr(values(row, pcDestination)))
        origins(a) = True: cities(a) = True: cities(b) = True
    Next row
    If Not cities.Exists(origin) Then Err.Raise 5, , "Город отправления вне покрытия: " & origin
    If Not cities.Exists(destination) Then Err.Raise 5, , "Город прибытия вне покрытия: " & destination
    For row = 1 To UBound(values, 1)
        If PlanningModeCode(CStr(values(row, pcMode))) <> "" Then
            a = PlanningCityKey(CStr(values(row, pcOrigin))): b = PlanningCityKey(CStr(values(row, pcDestination)))
            If b = destination Or origins.Exists(b) Then
                key = CStr(Len(a)) & ":" & a & b
                If Not groups.Exists(key) Then
                    Set indices = New Collection: groups.Add key, indices
                End If
                Set indices = groups(key): indices.Add row
            End If
        End If
    Next row
    For Each key In groups.Keys
        Set indices = groups(key)
        ReDim subset(1 To indices.Count, 1 To 20)
        n = 0
        For Each index In indices
            n = n + 1
            For j = 1 To 20: subset(n, j) = values(CLng(index), j): Next j
        Next index
        a = PlanningCityKey(CStr(subset(1, pcOrigin))): b = PlanningCityKey(CStr(subset(1, pcDestination)))
        allowance = -1
        If countryMap.Exists(a) And countryMap.Exists(b) Then
            If countryMap(a) <> "" And countryMap(b) <> "" Then
                If countryMap(a) = "RU" And countryMap(b) = "RU" Then allowance = allowanceRU Else allowance = allowanceAB
            End If
        End If
        Set estimates = PlanningEstimate(a, b, launch, cohort, semester, "rail,air,bus", baggage, cal, subset, railClasses, Empty, "", allowance)
        Set prices = estimates("prices"): Set methods = estimates("methods")
        For Each mode In prices.Keys
            known = 0: unknown = False
            If mode = "air" Then RoutingAirCount subset, months, UCase$(Trim$(baggage)) = "ДА", allowance, counts, known, unknown
            Set edge = CreateObject("Scripting.Dictionary")
            edge.Add "origin", a: edge.Add "destination", b: edge.Add "mode", CStr(mode)
            edge.Add "origin_title", CStr(subset(1, pcOrigin)): edge.Add "destination_title", CStr(subset(1, pcDestination))
            edge.Add "price", prices(mode): edge.Add "transfers", known: edge.Add "unknown", unknown
            edge.Add "method", CStr(methods(mode))
            edges.Add edge
        Next mode
    Next key
        Set graphEntry = CreateObject("Scripting.Dictionary")
        graphEntry.Add "edges", edges: graphEntry.Add "cities", cities
        RoutingStoreGraph graphKey, graphEntry
    Else
        Set edges = graphEntry("edges"): Set cities = graphEntry("cities")
        If Not cities.Exists(origin) Then Err.Raise 5, , "Город отправления вне покрытия: " & origin
        If Not cities.Exists(destination) Then Err.Raise 5, , "Город прибытия вне покрытия: " & destination
    End If
    If requestedMode <> "" Then preferences = requestedMode
    Set chosen = RoutingFind(edges, origin, destination, preferences, 2)
    Set result = CreateObject("Scripting.Dictionary")
    Set prices = CreateObject("Scripting.Dictionary"): Set methods = CreateObject("Scripting.Dictionary")
    result.Add "chosen", "": result.Add "prices", prices: result.Add "methods", methods
    result.Add "reasons", "Нет полного ценового пути по всем месяцам семестра с учётом классов, багажа и предела двух известных пересадок"
    If chosen("status") = "ok" Then
        chosenCode = "route": If requestedMode <> "" Then chosenCode = requestedMode
        Set path = chosen("legs")
        description = "Автомаршрут: " & CStr(path(1)("origin_title"))
        For Each edge In path
            description = description & " -> " & CStr(edge("destination_title"))
            If title <> "" Then title = title & " -> "
            title = title & PlanningModeTitle(CStr(edge("mode")))
        Next edge
        description = description & "; известных пересадок: " & CStr(chosen("known_transfers"))
        If Not chosen("transfer_limit_verified") Then description = description & "; число авиапересадок неизвестно, предел двух НЕ ПРОВЕРЕН"
        If path.Count > 1 Or chosen("known_transfers") > 0 Or Not chosen("transfer_limit_verified") Then
            description = description & "; СУММА БИЛЕТОВ: переезды при пересадке и ночёвки не оценены; стыковки по времени не проверены"
        End If
        For Each edge In path: description = description & " | " & CStr(edge("method")): Next edge
        prices.Add chosenCode, chosen("price"): methods.Add chosenCode, description
        result("chosen") = chosenCode
        result.Add "mode_title", title
        result.Add "route", chosen
    End If
    Set RoutingEstimate = result
End Function
