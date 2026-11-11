Attribute VB_Name = "modPlanningStorage"
Option Explicit

' Catalog boundary: typed JSON values in, text-safe worksheet values out.
Public Function PlanningCatalogFields() As Variant
    PlanningCatalogFields = Array("origin", "destination", "mode", "month", "base_year", "price", _
        "annual_rate", "baggage_included", "source", "previous_source", "reference_date", _
        "previous_date", "rate_method", "quote_id", "rail_class", "firm", "sample_count", "price_method", "uncertainty")
End Function

Public Function PlanningCoverageFields() As Variant
    PlanningCoverageFields = Array("origin", "destination", "mode", "status", "checked_at", "source")
End Function

Public Function PlanningText(ByVal record As Object, ByVal key As String, Optional ByVal required As Boolean = False) As String
    If Not record.Exists(key) Then
        If required Then Err.Raise 5, , "Отсутствует поле каталога: " & key
        Exit Function
    End If
    If VarType(record(key)) <> vbString Then Err.Raise 5, , "Ожидается строка каталога: " & key
    PlanningText = record(key)
    If required And Len(Trim$(PlanningText)) = 0 Then Err.Raise 5, , "Пустое поле каталога: " & key
End Function

Private Function CatalogNumber(ByVal record As Object, ByVal key As String) As Double
    If Not record.Exists(key) Then Err.Raise 5, , "Отсутствует число каталога: " & key
    Select Case VarType(record(key))
    Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
        CatalogNumber = CDbl(record(key))
    Case Else
        Err.Raise 5, , "Ожидается число, не строка или флаг: " & key
    End Select
End Function

Private Function CatalogBoolean(ByVal record As Object, ByVal key As String, Optional ByVal required As Boolean = True) As Boolean
    If Not record.Exists(key) Then
        If required Then Err.Raise 5, , "Отсутствует флаг каталога: " & key
        Exit Function
    End If
    If VarType(record(key)) <> vbBoolean Then Err.Raise 5, , "Ожидается логический флаг: " & key
    CatalogBoolean = record(key)
End Function

Public Function PlanningCatalogMatrix(ByVal root As Object) As Variant
    Dim records As Variant, matrix() As Variant, fields As Variant, rec As Object
    Dim i As Long, j As Long, value As Double, mode As String
    records = root("records"): fields = PlanningCatalogFields()
    If UBound(records) < 0 Then Err.Raise 5, , "Пустой набор плановых цен"
    If UBound(records) >= 1048575 Then Err.Raise 5, , "Каталог превышает размер листа"
    ReDim matrix(1 To UBound(records) + 1, 1 To pcUncertainty)
    For i = 0 To UBound(records)
        Set rec = records(i)
        For j = 0 To UBound(fields)
            Select Case j + 1
            Case pcMonth, pcBaseYear, pcPrice, pcAnnualRate
                matrix(i + 1, j + 1) = CatalogNumber(rec, CStr(fields(j)))
            Case pcBaggage
                matrix(i + 1, j + 1) = CatalogBoolean(rec, CStr(fields(j)))
            Case pcFirm
                matrix(i + 1, j + 1) = CatalogBoolean(rec, CStr(fields(j)), False)
            Case pcSampleCount
                value = 1
                If rec.Exists("sample_count") Then value = CatalogNumber(rec, "sample_count")
                If value <= 0 Or value <> Fix(value) Then Err.Raise 5, , "Некорректный размер выборки"
                matrix(i + 1, j + 1) = value
            Case Else
                matrix(i + 1, j + 1) = PlanningText(rec, CStr(fields(j)), _
                    j + 1 = pcOrigin Or j + 1 = pcDestination Or j + 1 = pcMode Or j + 1 = pcSource)
            End Select
        Next j
        value = matrix(i + 1, pcMonth)
        If value < 1 Or value > 12 Or value <> Fix(value) Then Err.Raise 5, , "Месяц каталога вне 1..12"
        value = matrix(i + 1, pcBaseYear)
        If value < 1 Or value > 9999 Or value <> Fix(value) Then Err.Raise 5, , "Некорректный базовый год"
        If matrix(i + 1, pcPrice) <= 0 Or matrix(i + 1, pcAnnualRate) <= -1 Then Err.Raise 5, , "Некорректная цена или индекс"
        mode = matrix(i + 1, pcMode)
        If mode <> "air" And mode <> "rail" And mode <> "bus" Then Err.Raise 5, , "Неизвестный вид транспорта"
        If mode = "rail" And matrix(i + 1, pcRailClass) = "" Then Err.Raise 5, , "Не указан класс ЖД"
    Next i
    PlanningCatalogMatrix = matrix
End Function

Public Function PlanningCoverageMatrix(ByVal root As Object) As Variant
    Dim rows As Variant, fields As Variant, matrix() As Variant, i As Long, j As Long, rec As Object
    rows = root("coverage"): fields = PlanningCoverageFields()
    If UBound(rows) < 0 Or UBound(rows) >= 1048575 Then Err.Raise 5, , "Некорректный размер покрытия"
    ReDim matrix(1 To UBound(rows) + 1, 1 To 6)
    For i = 0 To UBound(rows)
        Set rec = rows(i)
        For j = 0 To 5
            matrix(i + 1, j + 1) = PlanningText(rec, CStr(fields(j)), True)
        Next j
    Next i
    PlanningCoverageMatrix = matrix
End Function

Public Sub WritePlanningValues(ByVal target As Range, ByVal values As Variant)
    ' Value2 alone still interprets strings beginning with = as formulas.
    ' Set text format BEFORE assignment. Numeric/Boolean variants retain their types.
    target.NumberFormat = "@"
    target.Value2 = values
End Sub

Private Function BackupFolder() As String
    If ThisWorkbook.Path = "" Then Err.Raise 5, , "Сначала сохраните книгу"
    BackupFolder = ThisWorkbook.Path & "\_backup_план"
End Function

Private Function RegexLiteral(ByVal value As String) As String
    Dim i As Long, ch As String
    For i = 1 To Len(value)
        ch = Mid$(value, i, 1)
        If InStr(1, "\^$.|?*+()[]{}", ch, vbBinaryCompare) > 0 Then RegexLiteral = RegexLiteral & "\"
        RegexLiteral = RegexLiteral & ch
    Next i
End Function

Public Function PlanningBackupMatches(ByVal fileName As String, ByVal workbookName As String) As Boolean
    Dim fs As Object, pattern As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set pattern = CreateObject("VBScript.RegExp")
    pattern.IgnoreCase = True
    pattern.Pattern = "^" & RegexLiteral(fs.GetBaseName(workbookName)) & _
        "_[0-9]{8}_[0-9]{6}(_[0-9]{3})?\." & RegexLiteral(fs.GetExtensionName(workbookName)) & "$"
    PlanningBackupMatches = pattern.Test(fileName)
End Function

Public Sub CreatePlanningBackup()
    Dim fs As Object, folder As String, stem As String, file As String, suffix As Long, extension As String
    Set fs = CreateObject("Scripting.FileSystemObject")
    folder = BackupFolder()
    If Not fs.FolderExists(folder) Then fs.CreateFolder folder
    extension = "." & fs.GetExtensionName(ThisWorkbook.Name)
    stem = folder & "\" & fs.GetBaseName(ThisWorkbook.Name) & "_" & Format(Now, "yyyymmdd_hhnnss")
    file = stem & extension
    Do While fs.FileExists(file)
        suffix = suffix + 1
        If suffix > 999 Then Err.Raise 5, , "Слишком много резервных копий за секунду"
        file = stem & "_" & Format$(suffix, "000") & extension
    Loop
    ThisWorkbook.SaveCopyAs file
End Sub

Public Sub PrunePlanningBackups()
    ' Called only after the workbook operation succeeds; never prune on failure.
    Dim fs As Object, folder As String, count As Long, item As Object, oldest As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    folder = BackupFolder()
    Do
        count = 0: Set oldest = Nothing
        For Each item In fs.GetFolder(folder).Files
            If PlanningBackupMatches(item.Name, ThisWorkbook.Name) Then
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
