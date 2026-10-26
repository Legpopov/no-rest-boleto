Attribute VB_Name = "modPriceUpdate"
' Обновление рыночных цен (авиа, ЖД, гостиницы, курсы ЦБ) на листе «Справочник»
' из data/prices.json сборщика github.com/Legpopov/no-rest-boleto.
'
' Все пороги, лимиты и надбавки читаются с листа «Параметры цен» — в коде их нет.
' Кнопка «Обновить цены» вызывает UpdatePrices.
Option Explicit

' ---- структура книги (не нормы) ----
Private Const SH_REF As String = "Справочник"
Private Const SH_PAR As String = "Параметры цен"
Private Const SH_LOG As String = "Журнал обновления"
Private Const FIRST_ROW As Long = 4
Private Const TOTAL_MARK As String = "уточняется"   ' строка средних по РФ (формулы) — таблица заканчивается над ней
Private Const HEADER_ROW As Long = 2
Private Const C_COUNTRY As String = "BJ"
Private Const C_CITY As String = "BK"
Private Const C_AIR_AVG As String = "BO"
Private Const C_AIR_MAX As String = "BP"
Private Const C_RATE As String = "BQ"
Private Const C_HOT_RUB As String = "BR"
Private Const C_HOT_CUR As String = "BS"
Private Const C_HOT_VAL As String = "BT"
Private Const C_HOT_MAX As String = "BU"
Private Const C_RAIL_AVG As String = "BV"
Private Const C_RAIL_MAX As String = "BW"

' ---- имена параметров (колонка A листа «Параметры цен») ----
Private Const K_URL As String = "Источник данных (URL или путь к файлу)"
Private Const K_BAG_RU As String = "Авиа: багаж РФ, руб. за перелёт в одну сторону"
Private Const K_BAG_AB As String = "Авиа: багаж зарубеж, руб. за перелёт в одну сторону"
Private Const K_NOAIR As String = "Авиа: нет аэропорта — очищать ячейки"
Private Const K_RAIL_CLS As String = "ЖД: допустимые классы"
Private Const K_RAIL_Q As String = "ЖД: перцентиль для «макс.»"
Private Const K_H_LIM_RU As String = "Отель: лимит РФ, руб./ночь"
Private Const K_H_LIM_AB As String = "Отель: лимит зарубеж, руб./ночь"
Private Const K_H_STARS As String = "Отель: звёзд не меньше"
Private Const K_H_RATING As String = "Отель: рейтинг не меньше"
Private Const K_H_DIST As String = "Отель: до центра не дальше, км"
Private Const K_H_APART As String = "Отель: исключить апартаменты"
Private Const K_H_KINDS As String = "Отель: исключить типы размещения"
Private Const K_H_NOSTAR As String = "Отель: города без фильтра звёзд и рейтинга"
Private Const K_FRESH As String = "Предупреждать, если снимок старше, дней"
Private Const K_ACCEPT As String = "Критерий приёмки, % городов"
Private Const K_BACKUP As String = "Бэкап файла перед обновлением, хранить копий"
Private Const K_CUR_HEAD As String = "Валюта в «Справочнике»"
Private Const K_DEP As String = "Дата отправления"
Private Const K_DAYS As String = "Дней в поездке"
Private Const K_AIR_TOL As String = "Даты: авиа — допуск до ближайшей недели, дней"
Private Const K_HOT_TOL As String = "Даты: отель — допуск до ближайшего заезда, дней"
Private Const O_RUN As String = "Последний запуск"
Private Const O_SNAP As String = "Снимок данных от"
Private Const O_FX As String = "Курс ЦБ на"
Private Const O_TOTAL As String = "Итог"
Private Const O_DATES As String = "Даты расчёта"
Private Const DATES_TAG As String = " (на даты "    ' начало приписки к заголовкам в режиме дат

Private Const BUS_TAG As String = "(автобус)"   ' пометка в числовом формате ЖД-ячеек, когда записан автобус
Private Const SKIP_COLOR As Long = 10092543    ' RGB(255, 255, 153) — светло-жёлтый «пропуск»

' результаты по полю
Private Const R_OK As String = "обновлено"
Private Const R_SKIP As String = "пропущено"
Private Const R_NA As String = "не применимо"

Private gQuiet As Boolean
Private gLastError As String
' режим дат: цены на введённую дату отправления (приблизительно, из того же снимка)
Private gDateMode As Boolean
Private gDep As Date, gRet As Date, gDays As Long
Private nAirDated As Long, nAirUndated As Long, nHotDated As Long, nHotUndated As Long
Private prm As Object        ' имя -> Array(значение, включено)
Private curMap As Object     ' валюта из листа (нормализованная) -> код ЦБ
Private logRows As Collection

' ======================= точки входа =======================

Public Sub UpdatePrices()
    gQuiet = False
    RunUpdate
End Sub

' Без окон сообщений — для автоматической проверки. Возвращает "" или текст ошибки.
Public Function UpdatePricesQuiet() As String
    gQuiet = True
    gLastError = ""
    RunUpdate
    UpdatePricesQuiet = gLastError
End Function

' ======================= основной сценарий =======================

Private Sub RunUpdate()
    Dim lastRow As Long
    Dim wb As Workbook, ws As Worksheet, wsPar As Worksheet
    Dim calcMode As XlCalculation, startedAt As Date
    Dim data As Object, idx As Object, rates As Object, fxDate As String
    Dim r As Long, nCities As Long, nFull As Long, nPart As Long, nNone As Long
    Dim cityStat As String, msg As String, warn As String, step As String
    Dim acceptPct As Double, snapAge As Double

    Set wb = ThisWorkbook
    startedAt = Now
    Set logRows = New Collection
    calcMode = Application.Calculation

    On Error GoTo Failed
    step = "чтение параметров"
    Set ws = wb.Worksheets(SH_REF)
    Set wsPar = wb.Worksheets(SH_PAR)
    LoadParams wsPar
    step = "даты поездки"
    ReadTripDates wsPar
    nAirDated = 0: nAirUndated = 0: nHotDated = 0: nHotUndated = 0

    step = "резервная копия"
    warn = MakeBackup(wb)
    If gDateMode And gDays > 30 Then
        warn = warn & "Поездка длиннее 30 дней: цены перелётов в данных собраны для пребывания 1–30 ночей." & vbLf
    End If

    step = "загрузка данных"
    Application.StatusBar = "Обновление цен: загрузка данных..."
    Dim txt As String
    txt = LoadText(CStr(PVal(K_URL)), wb.Path)

    step = "разбор JSON"
    Application.StatusBar = "Обновление цен: разбор данных..."
    Set data = JsonParse(txt)
    txt = vbNullString
    If Not data.Exists("cities") Then Err.Raise vbObjectError + 801, , "в данных нет списка cities"

    Set idx = BuildIndex(data("cities"))
    If IsObject(JGet(data, "fx.rates")) Then Set rates = JGet(data, "fx.rates")
    fxDate = CStr(JGet(data, "fx.date"))

    snapAge = SnapshotAgeDays(CStr(JGet(data, "generated_at")))
    If PEnabled(K_FRESH) And snapAge > PNum(K_FRESH) Then
        warn = warn & "Снимок данных старый: " & Format(snapAge, "0.0") & " дн. " & _
               "Новые цены могут быть не получены — проверьте ночной сбор." & vbLf
    End If

    step = "заполнение таблицы"
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    EnsureHeaders ws, fxDate

    lastRow = TableLastRow(ws)
    For r = FIRST_ROW To lastRow
        If Trim$(CStr(ws.Cells(r, C_CITY).Value)) <> "" Then
            If r Mod 20 = 0 Then Application.StatusBar = "Обновление цен: строка " & r & " из " & lastRow
            cityStat = ProcessRow(ws, r, idx, rates)
            nCities = nCities + 1
            Select Case cityStat
            Case "full": nFull = nFull + 1
            Case "part": nPart = nPart + 1
            Case Else: nNone = nNone + 1
            End Select
        End If
    Next

    Application.Calculation = calcMode
    Application.ScreenUpdating = True
    Application.StatusBar = False

    step = "журнал"
    If PEnabled(K_ACCEPT) Then acceptPct = PNum(K_ACCEPT)
    WriteLog wb, startedAt, data, fxDate, nCities, nFull, nPart, nNone, acceptPct, warn
    WriteSummary wsPar, startedAt, data, fxDate, nCities, nFull, nPart

    If Not gQuiet Then
        msg = "Обновлено полностью: " & nFull & " из " & nCities & " (" & Format(SafePct(nFull, nCities), "0") & "%)" & vbLf & _
              "Частично: " & nPart & ", не обновлено: " & nNone & vbLf & _
              "Снимок данных: " & IsoToText(CStr(JGet(data, "generated_at"))) & vbLf & _
              "Курс ЦБ на: " & fxDate & vbLf & _
              "Даты: " & DatesText() & vbLf & vbLf & _
              "Пропуски залиты жёлтым, подробности — на листе «" & SH_LOG & "»."
        If warn <> "" Then msg = msg & vbLf & vbLf & "Внимание:" & vbLf & warn
        MsgBox msg, IIf(warn = "", vbInformation, vbExclamation), "Обновление цен"
    End If
    Exit Sub

Failed:
    Dim eMsg As String
    eMsg = "Ошибка на шаге «" & step & "»: " & Err.Description
    Application.Calculation = calcMode
    Application.ScreenUpdating = True
    Application.StatusBar = False
    On Error Resume Next
    LogLine 0, "", "", "", "ошибка", Empty, eMsg
    WriteLog wb, startedAt, Nothing, "", 0, 0, 0, 0, 0, eMsg
    On Error GoTo 0
    If gQuiet Then
        gLastError = eMsg
    Else
        MsgBox eMsg & vbLf & vbLf & "Таблица не изменена (или изменена частично до ошибки).", vbCritical, "Обновление цен"
    End If
End Sub

' Обрабатывает одну строку «Справочника». Возвращает "full" / "part" / "none".
Private Function ProcessRow(ws As Worksheet, ByVal r As Long, idx As Object, rates As Object) As String
    Dim city As String, place As String, key As String, c As Object
    Dim nOk As Long, nSkip As Long, res As String
    Dim abroad As Boolean, rate As Double, curText As String

    city = Trim$(CStr(ws.Cells(r, C_CITY).Value))
    place = Trim$(CStr(ws.Cells(r, C_COUNTRY).Value))
    key = NormKey(city) & "|" & NormKey(place)

    If Not idx.Exists(key) Then
        LogLine r, place, city, "все поля", R_SKIP, Empty, "город не найден в данных"
        MarkSkip ws.Range(ws.Cells(r, C_AIR_AVG), ws.Cells(r, C_HOT_RUB))
        MarkSkip ws.Range(ws.Cells(r, C_HOT_VAL), ws.Cells(r, C_HOT_MAX))
        MarkSkip ws.Range(ws.Cells(r, C_RAIL_AVG), ws.Cells(r, C_RAIL_MAX))
        ProcessRow = "none"
        Exit Function
    End If
    Set c = idx(key)
    abroad = IsTrue(JGet(c, "abroad"))

    ' --- авиа ---
    res = DoAir(ws, r, place, city, c, abroad)
    Tally res, nOk, nSkip
    ' --- ЖД ---
    res = DoRail(ws, r, place, city, c)
    Tally res, nOk, nSkip
    ' --- курс ---
    curText = Trim$(CStr(ws.Cells(r, C_HOT_CUR).Value))
    res = DoRate(ws, r, place, city, curText, rates, rate)
    Tally res, nOk, nSkip
    ' --- гостиница ---
    res = DoHotel(ws, r, place, city, c, abroad, curText, rate)
    Tally res, nOk, nSkip

    If nSkip = 0 And nOk > 0 Then
        ProcessRow = "full"
    ElseIf nOk > 0 Then
        ProcessRow = "part"
    Else
        ProcessRow = "none"
    End If
End Function

' Последняя строка таблицы цен: над строкой «уточняется» (средние по РФ), иначе — последняя заполненная.
' Так новые города можно вставлять внутрь таблицы без правки макроса.
Private Function TableLastRow(ws As Worksheet) As Long
    Dim r As Long, last As Long
    last = ws.Cells(ws.Rows.Count, C_CITY).End(xlUp).Row
    For r = FIRST_ROW To last
        If LCase$(Trim$(CStr(ws.Cells(r, C_CITY).Value))) = TOTAL_MARK Then
            TableLastRow = r - 1
            Exit Function
        End If
    Next
    TableLastRow = last
End Function

Private Sub Tally(ByVal res As String, ByRef nOk As Long, ByRef nSkip As Long)
    If res = R_OK Then nOk = nOk + 1
    If res = R_SKIP Then nSkip = nSkip + 1
End Sub

' ======================= поля =======================

Private Function DoAir(ws As Worksheet, ByVal r As Long, place As String, city As String, c As Object, ByVal abroad As Boolean) As String
    Dim blk As Variant, why As String, med As Variant, p90 As Variant, bag As Double
    Dim rng As Range, dateNote As String
    Set rng = ws.Range(ws.Cells(r, C_AIR_AVG), ws.Cells(r, C_AIR_MAX))

    If IsObject(JGet(c, "air")) Then Set blk = JGet(c, "air")
    If Not IsObject(blk) Then
        DoAir = Skip(rng, r, place, city, "Авиа", "нет блока air в данных")
        Exit Function
    End If
    If CStr(JGet(blk, "status")) = "no_airport" Then
        If PEnabled(K_NOAIR) Then
            rng.ClearContents
            MarkOk rng
            LogLine r, place, city, "Авиа", R_NA, Empty, "нет аэропорта — ячейки очищены"
        Else
            LogLine r, place, city, "Авиа", R_NA, Empty, "нет аэропорта — значение не менялось"
        End If
        DoAir = R_NA
        Exit Function
    End If
    why = BlockProblem(blk)
    If why <> "" Then DoAir = Skip(rng, r, place, city, "Авиа", why): Exit Function

    If gDateMode Then
        If PickAirWeek(JGet(blk, "by_week"), med, p90, dateNote) Then
            nAirDated = nAirDated + 1
        Else
            nAirUndated = nAirUndated + 1
            med = JGet(blk, "all.median")
            p90 = JGet(blk, "all.p90")
        End If
    Else
        med = JGet(blk, "all.median")
        p90 = JGet(blk, "all.p90")
    End If
    If Not IsNumeric(med) Or Not IsNumeric(p90) Or IsEmpty(med) Or IsEmpty(p90) Then
        DoAir = Skip(rng, r, place, city, "Авиа", "нет медианы/90-го перцентиля в данных")
        Exit Function
    End If
    If abroad Then
        If PEnabled(K_BAG_AB) Then bag = PNum(K_BAG_AB)
    Else
        If PEnabled(K_BAG_RU) Then bag = PNum(K_BAG_RU)
    End If
    ws.Cells(r, C_AIR_AVG).Value = Round(med + 2 * bag, 0)
    ws.Cells(r, C_AIR_MAX).Value = Round(p90 + 2 * bag, 0)
    MarkOk rng
    LogLine r, place, city, "Авиа сред.", R_OK, ws.Cells(r, C_AIR_AVG).Value, "медиана " & med & " + багаж 2x" & bag & DateNoteText(dateNote) & AgeNote(blk)
    LogLine r, place, city, "Авиа макс.", R_OK, ws.Cells(r, C_AIR_MAX).Value, "90-й перцентиль " & p90 & " + багаж 2x" & bag
    DoAir = R_OK
End Function

Private Function DoRail(ws As Worksheet, ByVal r As Long, place As String, city As String, c As Object) As String
    Dim blk As Variant, why As String, st As String
    Dim there() As Double, back() As Double, nT As Long, nB As Long, q As Double
    Dim avgV As Double, maxV As Double, rng As Range
    Set rng = ws.Range(ws.Cells(r, C_RAIL_AVG), ws.Cells(r, C_RAIL_MAX))

    If IsObject(JGet(c, "rail")) Then Set blk = JGet(c, "rail")
    If Not IsObject(blk) Then
        DoRail = Skip(rng, r, place, city, "ЖД", "нет блока rail в данных")
        Exit Function
    End If
    st = CStr(JGet(blk, "status"))
    If st = "no_station" Then
        DoRail = RailOrBus(ws, r, place, city, c, R_NA, "нет ж/д станции в данных")
        Exit Function
    End If
    why = BlockProblem(blk)
    If why <> "" Then DoRail = RailOrBus(ws, r, place, city, c, R_SKIP, why): Exit Function

    nT = RailFares(JGet(blk, "there"), there)
    nB = RailFares(JGet(blk, "back"), back)
    If nT = 0 Or nB = 0 Then
        DoRail = RailOrBus(ws, r, place, city, c, R_SKIP, "нет мест допустимых классов (туда: " & nT & ", обратно: " & nB & ")")
        Exit Function
    End If
    q = PNum(K_RAIL_Q)
    avgV = WorksheetFunction.Median(there) + WorksheetFunction.Median(back)
    maxV = WorksheetFunction.Percentile(there, q) + WorksheetFunction.Percentile(back, q)
    ws.Cells(r, C_RAIL_AVG).Value = Round(avgV, 0)
    ws.Cells(r, C_RAIL_MAX).Value = Round(maxV, 0)
    SetBusFormat ws, r, False
    MarkOk rng
    LogLine r, place, city, "ЖД сред.", R_OK, Round(avgV, 0), "медиана туда+обратно; тарифов " & nT & "/" & nB & AgeNote(blk)
    LogLine r, place, city, "ЖД макс.", R_OK, Round(maxV, 0), "перцентиль " & q & " туда+обратно"
    DoRail = R_OK
End Function

' Поезда нет: если есть автобус (bus.status = ok) — пишем его в колонки ЖД с пометкой «(автобус)» в формате ячейки;
' иначе — прежний исход для ЖД (пропуск с жёлтой заливкой или «не применимо»).
Private Function RailOrBus(ws As Worksheet, ByVal r As Long, place As String, city As String, c As Object, _
                           ByVal failRes As String, ByVal railWhy As String) As String
    Dim blk As Variant, there() As Double, back() As Double, nT As Long, nB As Long, q As Double
    Dim avgV As Double, maxV As Double, rng As Range, busWhy As String, how As String
    Set rng = ws.Range(ws.Cells(r, C_RAIL_AVG), ws.Cells(r, C_RAIL_MAX))
    If IsObject(JGet(c, "bus")) Then Set blk = JGet(c, "bus")
    If IsObject(blk) Then
        busWhy = BlockProblem(blk)
        If busWhy = "" Then
            nT = BusFares(JGet(blk, "there"), there)
            nB = BusFares(JGet(blk, "back"), back)
            If nT = 0 Then busWhy = "нет рейсов туда"
        End If
        If busWhy = "" Then
            q = PNum(K_RAIL_Q)
            If nB > 0 Then
                avgV = WorksheetFunction.Median(there) + WorksheetFunction.Median(back)
                maxV = WorksheetFunction.Percentile(there, q) + WorksheetFunction.Percentile(back, q)
                how = "медиана туда+обратно; рейсов " & nT & "/" & nB
            Else
                avgV = 2 * WorksheetFunction.Median(there)
                maxV = 2 * WorksheetFunction.Percentile(there, q)
                how = "обратных рейсов нет — туда x2; рейсов " & nT
            End If
            ws.Cells(r, C_RAIL_AVG).Value = Round(avgV, 0)
            ws.Cells(r, C_RAIL_MAX).Value = Round(maxV, 0)
            SetBusFormat ws, r, True
            MarkOk rng
            LogLine r, place, city, "ЖД сред.", R_OK, Round(avgV, 0), "ЖД: автобус (нет поезда: " & railWhy & "); " & how & AgeNote(blk)
            LogLine r, place, city, "ЖД макс.", R_OK, Round(maxV, 0), "автобус, перцентиль " & q & IIf(nB > 0, " туда+обратно", " туда x2")
            RailOrBus = R_OK
            Exit Function
        End If
        If CStr(Nz(JGet(blk, "status"))) <> "not_needed" Then railWhy = railWhy & "; автобус: " & busWhy
    End If
    If failRes = R_NA Then
        LogLine r, place, city, "ЖД", R_NA, Empty, railWhy
        RailOrBus = R_NA
    Else
        RailOrBus = Skip(rng, r, place, city, "ЖД", railWhy)
    End If
End Function

' Цены автобусов из списка рейсов [название, часов, цена RUB].
Private Function BusFares(ByVal trips As Variant, ByRef out() As Double) As Long
    Dim i As Long, t As Variant, cnt As Long
    ReDim out(0 To 0)
    If Not IsArray(trips) Then Exit Function
    For i = 0 To UBound(trips)
        t = trips(i)
        If IsArray(t) Then
            If UBound(t) >= 2 Then
                If NumOr(t(2), 0) > 0 Then
                    If cnt > UBound(out) Then ReDim Preserve out(0 To 2 * cnt + 1)
                    out(cnt) = t(2)
                    cnt = cnt + 1
                End If
            End If
        End If
    Next
    If cnt > 0 Then ReDim Preserve out(0 To cnt - 1)
    BusFares = cnt
End Function

' Пометка «(автобус)» — только числовым форматом, значение остаётся числом (для средних в итоговой строке).
' Для обычного ЖД возвращается формат соседней колонки авиа той же строки.
Private Sub SetBusFormat(ws As Worksheet, ByVal r As Long, ByVal isBus As Boolean)
    Dim col As Variant, cell As Range
    For Each col In Array(C_RAIL_AVG, C_RAIL_MAX)
        Set cell = ws.Cells(r, col)
        If isBus Then
            cell.NumberFormat = "#,##0"" " & BUS_TAG & """"
        ElseIf InStr(cell.NumberFormat, BUS_TAG) > 0 Then
            cell.NumberFormat = ws.Cells(r, C_AIR_MAX).NumberFormat
        End If
    Next
End Sub

' Цены допустимых классов из списка поездов [номер, фирменный 0/1, часы, {класс: цена}].
Private Function RailFares(ByVal trains As Variant, ByRef out() As Double) As Long
    Dim rules As Variant, i As Long, j As Long, t As Variant, cls As Object, k As Variant
    Dim firm As Boolean, cnt As Long, rule() As String, ok As Boolean
    ReDim out(0 To 0)
    If Not IsArray(trains) Then Exit Function
    rules = Split(Replace(CStr(PVal(K_RAIL_CLS)), ";", ","), ",")
    For i = 0 To UBound(trains)
        t = trains(i)
        If IsArray(t) Then
            If UBound(t) >= 3 Then
                firm = (Val(CStr(t(1))) <> 0)
                If IsObject(t(3)) Then
                    Set cls = t(3)
                    For Each k In cls.Keys
                        ok = False
                        For j = 0 To UBound(rules)
                            rule = Split(LCase$(Trim$(CStr(rules(j)))), ":")
                            If UBound(rule) >= 0 Then
                                If Trim$(rule(0)) = LCase$(CStr(k)) Then
                                    If UBound(rule) = 0 Then
                                        ok = True
                                    ElseIf Trim$(rule(1)) = "firm" Then
                                        ok = firm
                                    End If
                                End If
                            End If
                            If ok Then Exit For
                        Next
                        If ok And IsNumeric(cls(k)) Then
                            If cls(k) > 0 Then
                                If cnt > UBound(out) Then ReDim Preserve out(0 To 2 * cnt + 1)
                                out(cnt) = cls(k)
                                cnt = cnt + 1
                            End If
                        End If
                    Next
                End If
            End If
        End If
    Next
    If cnt > 0 Then ReDim Preserve out(0 To cnt - 1)
    RailFares = cnt
End Function

' Курс: из fx.rates по валюте из колонки BS. Пустая валюта (РФ) — рубли, курс не пишется.
Private Function DoRate(ws As Worksheet, ByVal r As Long, place As String, city As String, _
                        ByVal curText As String, rates As Object, ByRef rate As Double) As String
    Dim code As String, cell As Range
    Set cell = ws.Cells(r, C_RATE)
    rate = 0
    If curText = "" Then
        rate = 1
        DoRate = R_NA
        Exit Function
    End If
    code = CurrencyCode(curText)
    If code = "" Then
        DoRate = Skip(cell, r, place, city, "Курс", "валюта «" & curText & "» не найдена в таблице сопоставления")
        Exit Function
    End If
    If code = "RUB" Then
        rate = 1
    ElseIf rates Is Nothing Then
        DoRate = Skip(cell, r, place, city, "Курс", "нет курсов ЦБ в данных")
        Exit Function
    ElseIf Not rates.Exists(code) Then
        DoRate = Skip(cell, r, place, city, "Курс", "нет курса " & code & " в данных ЦБ")
        Exit Function
    Else
        rate = rates(code)
    End If
    cell.Value = Round(rate, 4)
    MarkOk cell
    LogLine r, place, city, "Курс", R_OK, cell.Value, code
    DoRate = R_OK
End Function

Private Function DoHotel(ws As Worksheet, ByVal r As Long, place As String, city As String, c As Object, _
                         ByVal abroad As Boolean, ByVal curText As String, ByVal rate As Double) As String
    Dim blk As Variant, why As String, items As Variant, it As Variant, i As Long
    Dim total As Long, nPass As Long, sumP As Double, maxP As Double, avgP As Double
    Dim rng As Range, rubFormula As Boolean, foreign As Boolean, dateNote As String
    Dim noStar As Boolean
    Set rng = ws.Range(ws.Cells(r, C_HOT_RUB), ws.Cells(r, C_HOT_MAX))
    noStar = CityInParam(city, K_H_NOSTAR)
    If noStar Then dateNote = "без фильтра звёзд/рейтинга"

    If IsObject(JGet(c, "hotels")) Then Set blk = JGet(c, "hotels")
    If Not IsObject(blk) Then
        DoHotel = SkipHotel(ws, r, place, city, "нет блока hotels в данных")
        Exit Function
    End If
    why = BlockProblem(blk)
    If why <> "" Then DoHotel = SkipHotel(ws, r, place, city, why): Exit Function

    If gDateMode Then
        Dim winNote As String
        items = PickHotelWindow(JGet(blk, "dated"), winNote)
        dateNote = JoinNote(dateNote, winNote)
        If IsArray(items) Then nHotDated = nHotDated + 1 Else nHotUndated = nHotUndated + 1
    End If
    If Not IsArray(items) Then items = JGet(blk, "items")
    If Not IsArray(items) Then DoHotel = SkipHotel(ws, r, place, city, "нет списка отелей"): Exit Function
    For i = 0 To UBound(items)
        it = items(i)
        If IsArray(it) Then
            total = total + 1
            If HotelPasses(it, abroad, noStar) Then
                nPass = nPass + 1
                sumP = sumP + it(0)
                If it(0) > maxP Then maxP = it(0)
            End If
        End If
    Next
    If nPass = 0 Then
        DoHotel = SkipHotel(ws, r, place, city, "после фильтров не осталось вариантов (было " & total & ")" & DateNoteText(dateNote))
        Exit Function
    End If
    avgP = sumP / nPass

    foreign = (curText <> "")
    rubFormula = ws.Cells(r, C_HOT_RUB).HasFormula
    If foreign And rate <= 0 Then
        If rubFormula Then
            DoHotel = SkipHotel(ws, r, place, city, "нет курса валюты — пересчитать в валюту нельзя")
        Else
            ws.Cells(r, C_HOT_RUB).Value = Round(avgP, 0)
            MarkOk ws.Cells(r, C_HOT_RUB)
            MarkSkip ws.Cells(r, C_HOT_VAL)
            MarkSkip ws.Cells(r, C_HOT_MAX)
            LogLine r, place, city, "Гостиница сред. руб.", R_OK, Round(avgP, 0), "среднее " & nPass & " из " & total & DateNoteText(dateNote) & AgeNote(blk)
            LogLine r, place, city, "Гостиница в валюте", R_SKIP, Empty, "нет курса валюты"
            DoHotel = R_SKIP
        End If
        Exit Function
    End If

    If foreign Then
        ws.Cells(r, C_HOT_VAL).Value = Round(avgP / rate, 0)
        ws.Cells(r, C_HOT_MAX).Value = Round(maxP / rate, 0)
        ws.Cells(r, C_HOT_MAX).NumberFormat = "0"" " & Replace(curText, """", "") & """"
    Else
        ws.Cells(r, C_HOT_VAL).Value = Round(avgP, 0)
        ws.Cells(r, C_HOT_MAX).Value = Round(maxP, 0)
    End If
    If Not rubFormula Then ws.Cells(r, C_HOT_RUB).Value = Round(avgP, 0)
    MarkOk ws.Cells(r, C_HOT_RUB)
    MarkOk ws.Cells(r, C_HOT_VAL)
    MarkOk ws.Cells(r, C_HOT_MAX)
    LogLine r, place, city, "Гостиница сред. руб.", R_OK, Round(avgP, 0), _
            "среднее " & nPass & " из " & total & IIf(rubFormula, " (BR — формула, пересчитана через BT)", "") & DateNoteText(dateNote) & AgeNote(blk)
    LogLine r, place, city, "Гостиница макс.", R_OK, ws.Cells(r, C_HOT_MAX).Value, _
            "максимум после фильтров" & IIf(foreign, ", в валюте «" & curText & "»", ", руб.")
    DoHotel = R_OK
End Function

Private Function SkipHotel(ws As Worksheet, ByVal r As Long, place As String, city As String, ByVal why As String) As String
    MarkSkip ws.Cells(r, C_HOT_RUB)
    MarkSkip ws.Cells(r, C_HOT_VAL)
    MarkSkip ws.Cells(r, C_HOT_MAX)
    LogLine r, place, city, "Гостиница", R_SKIP, Empty, why
    SkipHotel = R_SKIP
End Function

' Отель: [цена/ночь RUB, звёзды, рейтинг, расстояние м, тип, апартаменты 0/1]
' noStar = True — для города из списка «без фильтра звёзд и рейтинга»: эти два фильтра не применяются.
Private Function HotelPasses(it As Variant, ByVal abroad As Boolean, ByVal noStar As Boolean) As Boolean
    Dim kinds As Variant, j As Long, kind As String
    If UBound(it) < 5 Then Exit Function
    If Not IsNumeric(it(0)) Or IsNull(it(0)) Then Exit Function
    If it(0) <= 0 Then Exit Function
    If abroad Then
        If PEnabled(K_H_LIM_AB) Then If it(0) > PNum(K_H_LIM_AB) Then Exit Function
    Else
        If PEnabled(K_H_LIM_RU) Then If it(0) > PNum(K_H_LIM_RU) Then Exit Function
    End If
    If Not noStar Then
        If PEnabled(K_H_STARS) Then If NumOr(it(1), -1) < PNum(K_H_STARS) Then Exit Function
        If PEnabled(K_H_RATING) Then If NumOr(it(2), -1) < PNum(K_H_RATING) Then Exit Function
    End If
    If PEnabled(K_H_DIST) Then
        If NumOr(it(3), -1) < 0 Then Exit Function
        If it(3) > PNum(K_H_DIST) * 1000 Then Exit Function
    End If
    If PEnabled(K_H_APART) Then If NumOr(it(5), 0) <> 0 Then Exit Function
    If PEnabled(K_H_KINDS) Then
        kind = LCase$(Trim$(CStr(Nz(it(4)))))
        kinds = Split(Replace(CStr(PVal(K_H_KINDS)), ";", ","), ",")
        For j = 0 To UBound(kinds)
            If kind <> "" And kind = LCase$(Trim$(CStr(kinds(j)))) Then Exit Function
        Next
    End If
    HotelPasses = True
End Function

' ======================= режим дат =======================

' Дата отправления и число дней — на листе параметров (ищутся по подписи, значение — в ячейке справа).
' Пустая дата — обычный режим (средние за весь период снимка).
Private Sub ReadTripDates(ws As Worksheet)
    Dim f As Range, v As Variant
    gDateMode = False: gDays = 0
    Set f = ws.UsedRange.Find(What:=K_DEP, LookIn:=xlValues, LookAt:=xlWhole)
    If f Is Nothing Then Exit Sub
    v = f.Offset(0, 1).Value
    If IsEmpty(v) Then Exit Sub
    If Trim$(CStr(v)) = "" Then Exit Sub
    If Not IsDate(v) Then Err.Raise vbObjectError + 810, , "«" & K_DEP & "»: не похоже на дату («" & CStr(v) & "»)"
    gDep = DateSerial(Year(CDate(v)), Month(CDate(v)), Day(CDate(v)))
    If gDep < Date Then Err.Raise vbObjectError + 811, , "«" & K_DEP & "» " & Format(gDep, "dd.mm.yyyy") & _
                                  " уже прошла — укажите будущую дату или очистите ячейку"
    Set f = ws.UsedRange.Find(What:=K_DAYS, LookIn:=xlValues, LookAt:=xlWhole)
    If f Is Nothing Then Err.Raise vbObjectError + 812, , "на листе «" & SH_PAR & "» нет ячейки «" & K_DAYS & "»"
    v = f.Offset(0, 1).Value
    If IsEmpty(v) Or Not IsNumeric(v) Then Err.Raise vbObjectError + 813, , "«" & K_DAYS & "» — укажите целое число дней"
    If CDbl(v) < 1 Or CDbl(v) <> Int(CDbl(v)) Then Err.Raise vbObjectError + 813, , "«" & K_DAYS & "» — целое число не меньше 1"
    gDays = CLng(v)
    gRet = gDep + gDays - 1
    RequireParams Array(K_AIR_TOL, K_HOT_TOL)
    gDateMode = True
End Sub

Private Function DatesText() As String
    If gDateMode Then
        DatesText = Format(gDep, "dd.mm.yyyy") & " – " & Format(gRet, "dd.mm.yyyy") & " (" & gDays & " дн.)"
    Else
        DatesText = "не заданы — средние за период снимка"
    End If
End Function

Private Function DateNoteText(ByVal note As String) As String
    If note <> "" Then DateNoteText = "; " & note
End Function

Private Function JoinNote(ByVal a As String, ByVal b As String) As String
    If a <> "" And b <> "" Then JoinNote = a & "; " & b Else JoinNote = a & b
End Function

' Город есть в списке параметра (через запятую или «;», без учёта регистра и «ё»). Параметр может отсутствовать.
Private Function CityInParam(ByVal city As String, ByVal k As String) As Boolean
    Dim lst As Variant, i As Long
    If Not prm.Exists(k) Or Not PEnabled(k) Then Exit Function
    lst = Split(Replace(CStr(PVal(k)), ";", ","), ",")
    For i = 0 To UBound(lst)
        If NormKey(CStr(lst(i))) <> "" And NormKey(CStr(lst(i))) = NormKey(city) Then CityInParam = True: Exit Function
    Next
End Function

' Допуск в днях; если параметр выключен — без ограничения.
Private Function Tolerance(ByVal k As String) As Double
    If PEnabled(k) Then Tolerance = PNum(k) Else Tolerance = 1E+30
End Function

' "YYYY-MM-DD" -> Date или Empty
Private Function IsoDay(ByVal v As Variant) As Variant
    Dim t As String
    IsoDay = Empty
    t = CStr(Nz(v))
    If Len(t) < 10 Then Exit Function
    On Error Resume Next
    IsoDay = DateSerial(CInt(Mid$(t, 1, 4)), CInt(Mid$(t, 6, 2)), CInt(Mid$(t, 9, 2)))
    If Err.Number <> 0 Then IsoDay = Empty
    On Error GoTo 0
End Function

' air.by_week: [[понедельник недели вылета, n, median, p90], ...].
' Берётся неделя, в которую попадает дата отправления, иначе ближайшая в пределах допуска.
Private Function PickAirWeek(ByVal byWeek As Variant, ByRef med As Variant, ByRef p90 As Variant, ByRef note As String) As Boolean
    Dim i As Long, w As Variant, wk As Variant, dist As Double, best As Double, bi As Long, tol As Double
    tol = Tolerance(K_AIR_TOL)
    best = -1: bi = -1
    If IsArray(byWeek) Then
        For i = 0 To UBound(byWeek)
            w = byWeek(i)
            If IsArray(w) Then
                If UBound(w) >= 3 Then
                    wk = IsoDay(w(0))
                    If Not IsEmpty(wk) And NumOr(w(2), -1) > 0 And NumOr(w(3), -1) > 0 Then
                        If gDep < wk Then
                            dist = wk - gDep
                        ElseIf gDep > wk + 6 Then
                            dist = gDep - (wk + 6)
                        Else
                            dist = 0
                        End If
                        If best < 0 Or dist < best Then best = dist: bi = i
                    End If
                End If
            End If
        Next
    End If
    If bi < 0 Then
        note = "опорная дата: в данных нет цен по неделям — взята медиана за весь период"
    ElseIf best > tol Then
        note = "опорная дата: ближайшая неделя с ценами в " & best & " дн. от даты (допуск " & tol & ") — взята медиана за весь период"
    Else
        w = byWeek(bi)
        med = w(2): p90 = w(3)
        note = "опорная дата: авиа, неделя с " & Format(IsoDay(w(0)), "dd.mm") & IIf(best > 0, " (±" & best & " дн.)", "") & _
               ", вариантов " & Nz(w(1))
        PickAirWeek = True
    End If
End Function

' hotels.dated: [{checkin, nights, n, items}, ...] — окно с ближайшим к дате отправления заездом.
' Возвращает items окна или Empty (тогда берутся цены без даты).
Private Function PickHotelWindow(ByVal dated As Variant, ByRef note As String) As Variant
    Dim i As Long, w As Object, ci As Variant, dist As Double, best As Double, bi As Long, tol As Double
    PickHotelWindow = Empty
    tol = Tolerance(K_HOT_TOL)
    best = -1: bi = -1
    If IsArray(dated) Then
        For i = 0 To UBound(dated)
            If IsObject(dated(i)) Then
                Set w = dated(i)
                ci = IsoDay(JGet(w, "checkin"))
                If Not IsEmpty(ci) And IsArray(JGet(w, "items")) Then
                    dist = Abs(ci - gDep)
                    If best < 0 Or dist < best Then best = dist: bi = i
                End If
            End If
        Next
    End If
    If bi < 0 Then
        note = "опорная дата: в данных нет цен на даты заезда — взяты цены без даты"
    ElseIf best > tol Then
        note = "опорная дата: ближайший заезд в " & best & " дн. от даты (допуск " & tol & ") — взяты цены без даты"
    Else
        Set w = dated(bi)
        ci = IsoDay(JGet(w, "checkin"))
        note = "опорная дата: отель, заезд " & Format(ci, "dd.mm") & ", ночей " & Nz(JGet(w, "nights")) & ", " & ShiftText(ci - gDep)
        PickHotelWindow = JGet(w, "items")
    End If
End Function

' Положение опорной даты относительно даты отправления: «за 12 дн. до отправления» / «через 5 дн. после» / «в день отправления».
Private Function ShiftText(ByVal diff As Double) As String
    If diff = 0 Then
        ShiftText = "в день отправления"
    ElseIf diff < 0 Then
        ShiftText = "за " & -diff & " дн. до отправления"
    Else
        ShiftText = "через " & diff & " дн. после отправления"
    End If
End Function

' ======================= вспомогательное: данные =======================

' Индекс городов: "город|страна" и "город|регион" (для строк РФ в колонке страны стоит регион).
Private Function BuildIndex(cities As Variant) As Object
    Dim d As Object, i As Long, c As Object, city As String, k As String
    Set d = CreateObject("Scripting.Dictionary")
    If Not IsArray(cities) Then Set BuildIndex = d: Exit Function
    For i = 0 To UBound(cities)
        If IsObject(cities(i)) Then
            Set c = cities(i)
            city = NormKey(CStr(Nz(JGet(c, "city"))))
            k = city & "|" & NormKey(CStr(Nz(JGet(c, "country"))))
            If Not d.Exists(k) Then d.Add k, c
            If CStr(Nz(JGet(c, "region"))) <> "" Then
                k = city & "|" & NormKey(CStr(JGet(c, "region")))
                If Not d.Exists(k) Then d.Add k, c
            End If
        End If
    Next
    Set BuildIndex = d
End Function

' Пустая строка — блок годен; иначе причина пропуска.
Private Function BlockProblem(blk As Variant) As String
    Dim st As String, e As String
    st = CStr(Nz(JGet(blk, "status")))
    If st <> "ok" Then
        e = CStr(Nz(JGet(blk, "error")))
        BlockProblem = "статус «" & st & "»" & IIf(e <> "", ": " & e, "")
        Exit Function
    End If
    If IsTrue(JGet(blk, "stale")) Then
        BlockProblem = "устаревшие данные (от " & IsoToText(CStr(Nz(JGet(blk, "updated_at")))) & _
                       ", последняя попытка: " & CStr(Nz(JGet(blk, "last_attempt.status"))) & ")"
    End If
End Function

Private Function AgeNote(blk As Variant) As String
    AgeNote = "; данные от " & IsoToText(CStr(Nz(JGet(blk, "updated_at"))))
End Function

Private Function CurrencyCode(ByVal curText As String) As String
    Dim k As String
    k = NormKey(curText)
    If curMap.Exists(k) Then CurrencyCode = curMap(k)
End Function

Private Function NormKey(ByVal s As String) As String
    s = LCase$(Replace(Replace(s, "_", " "), ChrW(160), " "))
    s = Replace(s, "ё", "е")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    NormKey = Trim$(s)
End Function

Private Function IsTrue(v As Variant) As Boolean
    If VarType(v) = vbBoolean Then IsTrue = v
End Function

Private Function Nz(v As Variant) As Variant
    If IsNull(v) Or IsEmpty(v) Then Nz = "" Else Nz = v
End Function

Private Function NumOr(v As Variant, ByVal dflt As Double) As Double
    If IsNull(v) Or IsEmpty(v) Then
        NumOr = dflt
    ElseIf IsNumeric(v) Then
        NumOr = CDbl(v)
    Else
        NumOr = dflt
    End If
End Function

Private Function SafePct(ByVal a As Long, ByVal b As Long) As Double
    If b > 0 Then SafePct = 100# * a / b
End Function

' "2026-10-20T13:26:06+00:00" -> Date в UTC
Private Function IsoToDate(ByVal iso As String) As Variant
    IsoToDate = Empty
    If Len(iso) < 19 Then Exit Function
    On Error GoTo Bad
    IsoToDate = DateSerial(CInt(Mid$(iso, 1, 4)), CInt(Mid$(iso, 6, 2)), CInt(Mid$(iso, 9, 2))) + _
                TimeSerial(CInt(Mid$(iso, 12, 2)), CInt(Mid$(iso, 15, 2)), CInt(Mid$(iso, 18, 2)))
    If Len(iso) >= 25 Then   ' смещение +HH:MM
        IsoToDate = IsoToDate - IIf(Mid$(iso, 20, 1) = "-", -1, 1) * _
                    TimeSerial(CInt(Mid$(iso, 21, 2)), CInt(Mid$(iso, 24, 2)), 0)
    End If
    Exit Function
Bad:
    IsoToDate = Empty
End Function

Private Function IsoToText(ByVal iso As String) As String
    Dim d As Variant
    d = IsoToDate(iso)
    If IsEmpty(d) Then IsoToText = iso Else IsoToText = Format(d, "dd.mm.yyyy hh:nn") & " UTC"
End Function

Private Function SnapshotAgeDays(ByVal iso As String) As Double
    Dim d As Variant
    d = IsoToDate(iso)
    If IsEmpty(d) Then SnapshotAgeDays = 9999 Else SnapshotAgeDays = Now - d   ' погрешность — часовой пояс ПК
End Function

' ======================= вспомогательное: загрузка =======================

Private Function LoadText(ByVal src As String, ByVal basePath As String) As String
    Dim bytes() As Byte
    src = Trim$(src)
    If src = "" Then Err.Raise vbObjectError + 802, , "не задан источник данных на листе «" & SH_PAR & "»"
    If LCase$(Left$(src, 4)) = "http" Then
        bytes = HttpGet(src)
        LoadText = Utf8Decode(bytes)
    Else
        If InStr(src, ":") = 0 And Left$(src, 2) <> "\\" Then src = basePath & Application.PathSeparator & src
        If Dir(src) = "" Then Err.Raise vbObjectError + 803, , "файл не найден: " & src
        LoadText = ReadUtf8File(src)
    End If
End Function

Private Function HttpGet(ByVal url As String) As Byte()
    Dim h As Object, firstErr As String, u As String
    u = url & IIf(InStr(url, "?") > 0, "&", "?") & "nocache=" & Format(Now, "yyyymmddhhnnss")
    ' 1) WinHTTP
    On Error Resume Next
    Set h = CreateObject("WinHttp.WinHttpRequest.5.1")
    h.SetTimeouts 20000, 20000, 60000, 180000   ' мс: DNS, соединение, отправка, приём
    h.Open "GET", u, False
    h.SetRequestHeader "User-Agent", "Excel-VBA price updater"
    h.Send
    If Err.Number = 0 Then
        If h.Status = 200 Then HttpGet = h.ResponseBody: Exit Function
        firstErr = "HTTP " & h.Status
    Else
        firstErr = Err.Description
    End If
    Err.Clear
    ' 2) MSXML через настройки прокси Windows (WinINet) — на случай корпоративного прокси
    Set h = Nothing
    Set h = CreateObject("MSXML2.XMLHTTP.6.0")
    h.Open "GET", u, False
    h.Send
    If Err.Number = 0 Then
        If h.Status = 200 Then HttpGet = h.ResponseBody: Exit Function
        firstErr = firstErr & "; HTTP " & h.Status
    Else
        firstErr = firstErr & "; " & Err.Description
    End If
    On Error GoTo 0
    Err.Raise vbObjectError + 804, , "не удалось скачать " & url & " (" & firstErr & "). " & _
              "Проверьте интернет/VPN и адрес на листе «" & SH_PAR & "»."
End Function

Private Function Utf8Decode(bytes() As Byte) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1: st.Open
    st.Write bytes
    st.Position = 0
    st.Type = 2: st.Charset = "utf-8"
    Utf8Decode = st.ReadText
    st.Close
End Function

Private Function ReadUtf8File(ByVal path As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8": st.Open
    st.LoadFromFile path
    ReadUtf8File = st.ReadText
    st.Close
End Function

' Копия файла перед изменением; хранится не больше N последних. Возвращает предупреждение или "".
Private Function MakeBackup(wb As Workbook) As String
    Dim keep As Long, dirPath As String, fso As Object, f As Object, names As Collection
    Dim arr() As String, i As Long, j As Long, t As String, ext As String
    If Not PEnabled(K_BACKUP) Then Exit Function
    keep = CLng(PNum(K_BACKUP))
    If keep < 1 Then Exit Function
    On Error GoTo Bad
    If wb.Path = "" Or LCase$(Left$(wb.Path, 4)) = "http" Then
        MakeBackup = "Бэкап не сделан: файл не сохранён на диске." & vbLf
        Exit Function
    End If
    Set fso = CreateObject("Scripting.FileSystemObject")
    dirPath = wb.Path & Application.PathSeparator & "_backup_цены"
    If Not fso.FolderExists(dirPath) Then fso.CreateFolder dirPath
    ext = Mid$(wb.Name, InStrRev(wb.Name, "."))
    wb.SaveCopyAs dirPath & Application.PathSeparator & _
                  Left$(wb.Name, InStrRev(wb.Name, ".") - 1) & " " & Format(Now, "yyyy-mm-dd_hhnnss") & ext
    ' удаляем старые копии сверх лимита (имена с датой сортируются по времени)
    Set names = New Collection
    For Each f In fso.GetFolder(dirPath).Files
        If LCase$(Right$(f.Name, Len(ext))) = LCase$(ext) Then names.Add f.Name
    Next
    If names.Count > keep Then
        ReDim arr(1 To names.Count)
        For i = 1 To names.Count: arr(i) = names(i): Next
        For i = 1 To UBound(arr) - 1
            For j = i + 1 To UBound(arr)
                If arr(j) < arr(i) Then t = arr(i): arr(i) = arr(j): arr(j) = t
            Next
        Next
        For i = 1 To UBound(arr) - keep
            fso.DeleteFile dirPath & Application.PathSeparator & arr(i), True
        Next
    End If
    Exit Function
Bad:
    MakeBackup = "Бэкап не сделан: " & Err.Description & vbLf
End Function

' ======================= вспомогательное: параметры =======================

Private Sub LoadParams(ws As Worksheet)
    Dim r As Long, k As String, en As String, lastR As Long, curCol As Long, c As Long
    Set prm = CreateObject("Scripting.Dictionary")
    Set curMap = CreateObject("Scripting.Dictionary")
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 1 To lastR
        k = Trim$(CStr(ws.Cells(r, 1).Value))
        If k <> "" And Not prm.Exists(k) Then
            en = UCase$(Trim$(CStr(ws.Cells(r, 3).Value)))
            prm.Add k, Array(ws.Cells(r, 2).Value, Not (en = "НЕТ" Or en = "NO" Or en = "0"))
        End If
    Next
    ' таблица сопоставления валют: заголовок K_CUR_HEAD, ниже пары «как в листе» — «код ЦБ»
    For c = 1 To 20
        For r = 1 To 10
            If Trim$(CStr(ws.Cells(r, c).Value)) = K_CUR_HEAD Then curCol = c: Exit For
        Next
        If curCol > 0 Then Exit For
    Next
    If curCol = 0 Then Err.Raise vbObjectError + 805, , "на листе «" & SH_PAR & "» нет таблицы «" & K_CUR_HEAD & "»"
    r = r + 1
    Do While Trim$(CStr(ws.Cells(r, curCol).Value)) <> ""
        k = NormKey(CStr(ws.Cells(r, curCol).Value))
        If Not curMap.Exists(k) Then curMap.Add k, UCase$(Trim$(CStr(ws.Cells(r, curCol + 1).Value)))
        r = r + 1
    Loop
    RequireParams Array(K_URL, K_BAG_RU, K_BAG_AB, K_NOAIR, K_RAIL_CLS, K_RAIL_Q, K_H_LIM_RU, K_H_LIM_AB, _
                        K_H_STARS, K_H_RATING, K_H_DIST, K_H_APART, K_H_KINDS)
End Sub

Private Sub RequireParams(keys As Variant)
    Dim i As Long, miss As String
    For i = 0 To UBound(keys)
        If Not prm.Exists(keys(i)) Then miss = miss & vbLf & "  • " & keys(i)
    Next
    If miss <> "" Then Err.Raise vbObjectError + 806, , "на листе «" & SH_PAR & "» не найдены параметры:" & miss
End Sub

Private Function PVal(ByVal k As String) As Variant
    If prm.Exists(k) Then PVal = prm(k)(0) Else PVal = Empty
End Function

Private Function PNum(ByVal k As String) As Double
    Dim v As Variant
    v = PVal(k)
    If IsNumeric(v) And Not IsEmpty(v) Then
        PNum = CDbl(v)
    Else
        Err.Raise vbObjectError + 807, , "параметр «" & k & "» должен быть числом (сейчас: «" & CStr(v) & "»)"
    End If
End Function

Private Function PEnabled(ByVal k As String) As Boolean
    If prm.Exists(k) Then PEnabled = prm(k)(1)
End Function

' ======================= вспомогательное: лист и журнал =======================

Private Sub EnsureHeaders(ws As Worksheet, ByVal fxDate As String)
    Dim col As Variant, t As String, k As Long
    If fxDate <> "" Then ws.Cells(HEADER_ROW, C_RATE).Value = "Курс ЦБ на " & fxDate
    ' приписка «на даты …» у колонок, зависящих от дат (снимается, когда даты очищены)
    For Each col In Array(C_AIR_AVG, C_AIR_MAX, C_HOT_RUB, C_HOT_VAL, C_HOT_MAX)
        t = CStr(ws.Cells(HEADER_ROW, col).Value)
        k = InStr(t, DATES_TAG)
        If k > 0 Then t = Left$(t, k - 1)
        If gDateMode Then t = t & DATES_TAG & Format(gDep, "dd.mm") & "–" & Format(gRet, "dd.mm") & ", приблизительно)"
        If t <> CStr(ws.Cells(HEADER_ROW, col).Value) Then ws.Cells(HEADER_ROW, col).Value = t
    Next
    If Trim$(CStr(ws.Cells(HEADER_ROW, C_RAIL_AVG).Value)) = "" Then
        ws.Cells(HEADER_ROW, C_RAIL_AVG).Value = "ЖД - сред. цена (туда-обратно)"
    End If
    If Trim$(CStr(ws.Cells(HEADER_ROW, C_RAIL_MAX).Value)) = "" Then
        ws.Cells(HEADER_ROW, C_RAIL_MAX).Value = "ЖД - макс. цена (туда-обратно)"
    End If
End Sub

Private Sub MarkSkip(rng As Range)
    rng.Interior.Color = SKIP_COLOR
End Sub

Private Sub MarkOk(rng As Range)
    Dim cell As Range
    For Each cell In rng.Cells
        If cell.Interior.Pattern <> xlNone And cell.Interior.Color = SKIP_COLOR Then cell.Interior.Pattern = xlNone
    Next
End Sub

Private Function Skip(rng As Range, ByVal r As Long, place As String, city As String, fld As String, why As String) As String
    MarkSkip rng
    LogLine r, place, city, fld, R_SKIP, Empty, why
    Skip = R_SKIP
End Function

Private Sub LogLine(ByVal r As Long, place As String, city As String, fld As String, res As String, v As Variant, why As String)
    logRows.Add Array(IIf(r > 0, r, ""), place, city, fld, res, v, why)
End Sub

Private Function GetLogSheet(wb As Workbook) As Worksheet
    On Error Resume Next
    Set GetLogSheet = wb.Worksheets(SH_LOG)
    On Error GoTo 0
    If GetLogSheet Is Nothing Then
        Set GetLogSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        GetLogSheet.Name = SH_LOG
    End If
End Function

Private Sub WriteLog(wb As Workbook, ByVal startedAt As Date, data As Object, ByVal fxDate As String, _
                     ByVal nCities As Long, ByVal nFull As Long, ByVal nPart As Long, ByVal nNone As Long, _
                     ByVal acceptPct As Double, ByVal warn As String)
    Dim ws As Worksheet, arr() As Variant, i As Long, j As Long, top As Long, pct As Double
    Dim gen As String, nOkF As Long, nSkipF As Long, nNaF As Long, rr As Variant
    Set ws = GetLogSheet(wb)
    ws.Cells.Clear
    If Not data Is Nothing Then gen = IsoToText(CStr(Nz(JGet(data, "generated_at"))))
    pct = SafePct(nFull, nCities)
    For i = 1 To logRows.Count
        Select Case logRows(i)(4)
        Case R_OK: nOkF = nOkF + 1
        Case R_SKIP: nSkipF = nSkipF + 1
        Case R_NA: nNaF = nNaF + 1
        End Select
    Next

    ws.Range("A1").Value = "Журнал обновления цен"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 13
    Dim head As Variant
    head = Array( _
        Array("Запуск", Format(startedAt, "dd.mm.yyyy hh:nn:ss")), _
        Array("Снимок данных (generated_at)", gen), _
        Array("Курс ЦБ на", fxDate), _
        Array("Даты поездки", DatesText()), _
        Array("Цены на даты", IIf(gDateMode, "авиа: " & nAirDated & " по неделе вылета, " & nAirUndated & " без даты; " & _
              "отели: " & nHotDated & " по дате заезда, " & nHotUndated & " без даты (приблизительно)", "режим дат выключен")), _
        Array("Городов в таблице", nCities), _
        Array("Обновлено полностью", nFull & " (" & Format(pct, "0.0") & "%)"), _
        Array("Обновлено частично", nPart), _
        Array("Не обновлено", nNone), _
        Array("Полей: обновлено / пропущено / не применимо", nOkF & " / " & nSkipF & " / " & nNaF), _
        Array("Критерий приёмки", IIf(acceptPct > 0, "не меньше " & acceptPct & "% городов полностью: " & _
              IIf(pct >= acceptPct, "ВЫПОЛНЕН", "НЕ ВЫПОЛНЕН"), "не задан")))
    For i = 0 To UBound(head)
        ws.Cells(2 + i, 1).Value = head(i)(0)
        ws.Cells(2 + i, 2).Value = head(i)(1)
    Next
    top = 2 + UBound(head)                      ' строка критерия приёмки (последняя в шапке)
    ws.Range(ws.Cells(2, 2), ws.Cells(top, 2)).HorizontalAlignment = xlLeft
    ws.Range(ws.Cells(2, 1), ws.Cells(top, 1)).Font.Bold = True
    If acceptPct > 0 Then ws.Cells(top, 2).Interior.Color = IIf(pct >= acceptPct, RGB(198, 239, 206), RGB(255, 199, 206))
    top = top + 1
    If warn <> "" Then
        ws.Cells(top, 1).Value = "Внимание"
        ws.Cells(top, 1).Font.Bold = True
        ws.Cells(top, 2).Value = Replace(warn, vbLf, " ")
        ws.Cells(top, 2).Font.Color = RGB(192, 0, 0)
        top = top + 1
    End If
    top = top + 1
    ws.Cells(top, 1).Resize(1, 7).Value = Array("Строка", "Страна / регион", "Город", "Поле", "Результат", "Значение", "Пояснение / причина")
    ws.Cells(top, 1).Resize(1, 7).Font.Bold = True
    ws.Cells(top, 1).Resize(1, 7).Interior.Color = RGB(221, 235, 247)
    If logRows.Count > 0 Then
        ReDim arr(1 To logRows.Count, 1 To 7)
        For i = 1 To logRows.Count
            rr = logRows(i)
            For j = 0 To 6
                arr(i, j + 1) = rr(j)
            Next
        Next
        ws.Cells(top + 1, 1).Resize(logRows.Count, 7).Value = arr
        For i = 1 To logRows.Count
            If arr(i, 5) = R_SKIP Or arr(i, 5) = "ошибка" Then ws.Cells(top + i, 5).Interior.Color = SKIP_COLOR
        Next
    End If
    ws.Columns("A:F").AutoFit
    ws.Columns("G").ColumnWidth = 90
    ws.Columns("A").ColumnWidth = 44
End Sub

Private Sub WriteSummary(ws As Worksheet, ByVal startedAt As Date, data As Object, ByVal fxDate As String, _
                         ByVal nCities As Long, ByVal nFull As Long, ByVal nPart As Long)
    SetOut ws, O_RUN, Format(startedAt, "dd.mm.yyyy hh:nn")
    SetOut ws, O_SNAP, IsoToText(CStr(Nz(JGet(data, "generated_at"))))
    SetOut ws, O_FX, fxDate
    SetOut ws, O_DATES, DatesText()
    SetOut ws, O_TOTAL, "полностью " & nFull & " из " & nCities & " (" & Format(SafePct(nFull, nCities), "0") & "%), частично " & nPart
End Sub

Private Sub SetOut(ws As Worksheet, ByVal k As String, ByVal v As Variant)
    Dim f As Range
    Set f = ws.Columns(1).Find(What:=k, LookIn:=xlValues, LookAt:=xlWhole)
    If Not f Is Nothing Then f.Offset(0, 1).Value = v
End Sub
