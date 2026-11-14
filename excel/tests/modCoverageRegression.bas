Attribute VB_Name = "modCoverageRegression"
Option Explicit

Private Sub Check(ByVal condition As Boolean, ByVal message As String, ByRef passed As Long)
    If Not condition Then Err.Raise 5, , "Coverage regression: " & message
    passed = passed + 1
End Sub

Public Function RunCoverageRegression() As Variant
    On Error GoTo Failed
    Dim cal(1 To 4, 1 To 1) As Variant, rows(1 To 10, 1 To 20) As Variant
    Dim hotels(1 To 1, 1 To 10) As Variant, result As Object, i As Long, j As Long
    Dim expected As Double, actual As Variant, passed As Long, reason As String
    cal(1, 1) = 9: cal(2, 1) = 5: cal(3, 1) = 2: cal(4, 1) = 5
    For i = 1 To 10
        rows(i, 1) = "A": rows(i, 2) = "B": rows(i, 3) = "air"
        rows(i, 4) = 2 + ((i - 1) Mod 5): rows(i, 5) = 2026
        rows(i, 6) = 100: rows(i, 7) = 0.1: rows(i, 8) = True
        rows(i, 9) = "fixture": rows(i, 13) = "route_reference_yoy_repeated_scenario"
        rows(i, 17) = 1
        If i > 5 Then
            rows(i, 5) = 2027: rows(i, 6) = 200: rows(i, 7) = 0.2: rows(i, 8) = False
            rows(i, 13) = "air_current_official_category_scenario"
        End If
    Next i
    Set result = PlanningEstimate("A", "B", 2027, "1", "2", "air", "ДА", cal, rows, "", Empty, "", 50)
    Check Abs(result("prices")("air") - 300) < 0.001, "current fare and baggage allowance once, actual 2027 base year", passed
    Set result = PlanningEstimate("A", "B", 2027, "1", "2", "air", "НЕТ", cal, rows)
    Check Abs(result("prices")("air") - 240) < 0.001, "unknown baggage does not imply an included bag", passed
    Set result = PlanningEstimate("A", "B", 2027, "1", "2", "air", "ДА", cal, rows)
    Check Abs(result("prices")("air") - 121) < 0.001, "without explicit allowance use known-baggage historical fallback", passed
    For i = 6 To 10: rows(i, 5) = 2029: Next i
    Set result = PlanningEstimate("A", "B", 2027, "1", "2", "air", "ДА", cal, rows, "", Empty, "", 50)
    Check Abs(result("prices")("air") - 121) < 0.001, "future price cannot be backdated", passed
    hotels(1, 1) = "B": hotels(1, 2) = 1000: hotels(1, 3) = 2026: hotels(1, 4) = 11
    hotels(1, 5) = 0.1: hotels(1, 6) = "price fixture": hotels(1, 7) = "index fixture"
    hotels(1, 8) = "2026-11-11": hotels(1, 9) = "hotel_official_category_scenario": hotels(1, 10) = "ok"
    For j = 15 To 19: expected = expected + 1000 * 1.1 ^ (j / 12#): Next j
    expected = Application.Round(6 * expected / 5, 2)
    actual = PlanningIndexedStay(2, "B", "пребывание", 3, 2027, "1", "2", cal, hotels, "price")
    Check Abs(CDbl(actual) - expected) < 0.001, "hotel index uses semester months, nights and people", passed
    actual = PlanningIndexedStay(2, "B", "пересадка", "", 2027, "1", "2", cal, hotels, "price")
    Check actual = 0, "transit without nights stays zero", passed
    hotels(1, 10) = "Нет свежих цен после фильтров книги"
    On Error Resume Next
    actual = PlanningIndexedStay(2, "B", "пребывание", 3, 2027, "1", "2", cal, hotels, "price")
    reason = Err.Description: Err.Clear
    On Error GoTo Failed
    Check InStr(reason, "нет свежих цен") > 0, "stale or filtered-out hotel is not indexed", passed
    RunCoverageRegression = passed
    Exit Function
Failed:
    RunCoverageRegression = "FAIL: " & Err.Description
End Function
