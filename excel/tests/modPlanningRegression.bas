Attribute VB_Name = "modPlanningRegression"
Option Explicit

Private Sub Check(ByVal condition As Boolean, ByVal message As String, ByRef passed As Long)
    If Not condition Then Err.Raise 5, , "Regression: " & message
    passed = passed + 1
End Sub

Public Function RunPlanningRegression(ByVal fixturePath As String, ByVal catalogPath As String) As Variant
    On Error GoTo Fatal
    Dim stream As Object, root As Object, cases As Variant, test As Variant, c As Object
    Dim cal(1 To 4, 1 To 1) As Variant, months As Variant, expected As Variant
    Dim fares As Variant, fare As Variant, f As Object, matrix() As Variant, count As Long
    Dim n As Long, i As Long, m As Variant, result As Object, prices As Object, prefs As Variant
    Dim passed As Long, scratch As Worksheet, values(1 To 1, 1 To 3) As Variant
    Dim badNumber As Variant, badFlag As Variant, saved As Variant, errorNumber As Long
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2: stream.Charset = "utf-8": stream.Open
    stream.LoadFromFile fixturePath
    Set root = JsonParse(stream.ReadText): stream.Close
    cases = root("cases")
    For Each test In cases
        Set c = test
        expected = c("calendar")
        For i = 0 To 3: cal(i + 1, 1) = expected(i): Next i
        months = PlanningSemesterMonths(c("launch"), CStr(c("cohort")), CStr(c("semester")), cal)
        expected = c("months")
        Check UBound(months) = UBound(expected), c("name") & " month count", passed
        For i = 0 To UBound(months)
            Check Format$(months(i), "yyyy-mm") = expected(i), c("name") & " calendar", passed
        Next i
        fares = c("fares"): count = 0
        For Each fare In fares
            Set f = fare: expected = f("months"): count = count + UBound(expected) + 1
        Next fare
        ReDim matrix(1 To count, 1 To pcUncertainty): n = 0
        For Each fare In fares
            Set f = fare: expected = f("months")
            For Each m In expected
                n = n + 1
                matrix(n, pcOrigin) = "A": matrix(n, pcDestination) = "B"
                matrix(n, pcMode) = f("mode"): matrix(n, pcMonth) = m
                matrix(n, pcBaseYear) = 2026: matrix(n, pcPrice) = f("price")
                matrix(n, pcAnnualRate) = f("rate"): matrix(n, pcBaggage) = True
                matrix(n, pcSource) = "shared fixture scenario": matrix(n, pcRailClass) = "coupe"
                matrix(n, pcFirm) = False: matrix(n, pcSampleCount) = 1
            Next m
        Next fare
        prefs = c("preferences")
        Set result = PlanningEstimate("A", "B", c("launch"), CStr(c("cohort")), _
                        CStr(c("semester")), Join(prefs, ","), "睦", cal, matrix, "coupe")
        Set prices = result("prices")
        If IsNull(c("mode")) Then
            Check result("chosen") = "", c("name") & " missing is not zero", passed
        Else
            Check result("chosen") = c("mode"), c("name") & " chosen mode", passed
            Check Abs(prices(result("chosen")) - c("price")) < 0.0001, c("name") & " price", passed
        End If
    Next test
    Check PlanningBackupMatches("想囗_20261111_010000.xlsm", "想囗.xlsm"), "own backup", passed
    Check PlanningBackupMatches("想囗_20261111_010000_001.xlsm", "想囗.xlsm"), "collision suffix", passed
    Check Not PlanningBackupMatches("想囗_2027_20261111_010000.xlsm", "想囗.xlsm"), "other workbook prefix", passed
    Check Not PlanningBackupMatches("想囗_20261111_010000.txt", "想囗.xlsm"), "wrong extension", passed
    Check PlanningBackupMatches("想囗 (1)+_20261111_010000.xlsm", "想囗 (1)+.xlsm"), "regex literal name", passed
    Check Not PlanningBackupMatches("想囗 11_20261111_010000.xlsm", "想囗 (1)+.xlsm"), "regex must not expand name", passed
    stream.Open: stream.LoadFromFile catalogPath
    Set root = JsonParse(stream.ReadText): stream.Close
    fares = root("records"): Set f = fares(0)
    saved = f("price"): f("price") = "123"
    On Error Resume Next
    badNumber = PlanningCatalogMatrix(root): errorNumber = Err.Number: Err.Clear
    On Error GoTo Fatal
    Check errorNumber <> 0, "numeric string rejected before import", passed
    f("price") = saved
    saved = f("baggage_included"): f("baggage_included") = "false": errorNumber = 0
    On Error Resume Next
    badFlag = PlanningCatalogMatrix(root): errorNumber = Err.Number: Err.Clear
    On Error GoTo Fatal
    Check errorNumber <> 0, "boolean string rejected before import", passed
    f("baggage_included") = saved
    f("source") = "=1+1": root("retrieved_at") = "=1+1"
    badNumber = PlanningCatalogMatrix(root)
    Set scratch = ThisWorkbook.Worksheets.Add
    On Error GoTo CleanFailed
    WritePlanningValues scratch.Range("A1").Resize(UBound(badNumber, 1), pcUncertainty), badNumber
    Check Not scratch.Cells(1, pcSource).HasFormula, "catalog source cannot become formula", passed
    Check scratch.Cells(1, pcSource).Value2 = "=1+1", "catalog source preserved literally", passed
    values(1, 1) = "=1+1": values(1, 2) = 123.45: values(1, 3) = True
    WritePlanningValues scratch.Range("U1:W1"), values
    Check Not scratch.Range("U1").HasFormula, "coverage text cannot become formula", passed
    Check VarType(scratch.Range("V1").Value2) = vbDouble, "numbers retain numeric type", passed
    Check VarType(scratch.Range("W1").Value2) = vbBoolean, "booleans retain boolean type", passed
    WritePlanningValues scratch.Range("U2"), PlanningText(root, "retrieved_at", True)
    Check Not scratch.Range("U2").HasFormula, "retrieval metadata cannot become formula", passed
    scratch.Delete
    RunPlanningRegression = passed
    Exit Function
CleanFailed:
    errorNumber = Err.Number: saved = Err.Description
    scratch.Delete
    RunPlanningRegression = "FAIL: " & saved
    Exit Function
Fatal:
    RunPlanningRegression = "FAIL: " & Err.Description
End Function
