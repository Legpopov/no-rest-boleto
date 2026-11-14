Attribute VB_Name = "modRoutingRegression"
Option Explicit

Private Sub Check(ByVal condition As Boolean, ByVal message As String, ByRef passed As Long)
    If Not condition Then Err.Raise 5, , "Routing regression: " & message
    passed = passed + 1
End Sub

Public Function RunRoutingRegression(ByVal fixturePath As String) As Variant
    On Error GoTo Failed
    Dim stream As Object, root As Object, samples As Variant, sample As Object, row As Variant
    Dim edges As Collection, edge As Object, result As Object, path As Collection, item As Object
    Dim order As String, signature As String, expected As String, cities As Variant, i As Long, passed As Long
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2: stream.Charset = "utf-8": stream.Open
    stream.LoadFromFile fixturePath
    Set root = JsonParse(stream.ReadText): stream.Close
    samples = root("cases")
    For i = 0 To UBound(samples)
        Set sample = samples(i): Set edges = New Collection
        For Each row In sample("edges")
            Set edge = CreateObject("Scripting.Dictionary")
            edge.Add "origin", CStr(row(0)): edge.Add "destination", CStr(row(1))
            edge.Add "mode", CStr(row(2)): edge.Add "price", CDbl(row(3)): edge.Add "transfers", row(4)
            edges.Add edge
        Next row
        order = Join(sample("preferences"), ",")
        Set result = RoutingFind(edges, CStr(sample("origin")), CStr(sample("destination")), order)
        Check result("status") = sample("status"), CStr(sample("name")) & " status", passed
        If result("status") = "ok" Then
            Check Abs(CDbl(result("price")) - CDbl(sample("price"))) < 0.001, CStr(sample("name")) & " price", passed
            Check result("known_transfers") = sample("known_transfers"), CStr(sample("name")) & " count", passed
            Check result("transfer_limit_verified") = sample("verified"), CStr(sample("name")) & " verified", passed
            Set path = result("legs")
            signature = CStr(path(1)("origin"))
            For Each item In path: signature = signature & "," & CStr(item("destination")): Next item
            cities = sample("cities"): expected = Join(cities, ",")
            Check signature = expected, CStr(sample("name")) & " path", passed
        Else
            Check IsNull(result("price")), CStr(sample("name")) & " missing price", passed
        End If
    Next i
    Dim values(1 To 2, 1 To 20) As Variant, matrix As Variant, failed As Boolean
    values(1, pcQuoteId) = "a": values(1, pcMode) = "air"
    values(2, pcQuoteId) = "b": values(2, pcMode) = "rail"
    Set root = CreateObject("Scripting.Dictionary")
    matrix = RoutingTransfersMatrix(root, values)
    Check matrix(1, 2) = -1 And matrix(2, 2) = 0, "legacy catalog preserves unknown air transfers", passed
    root.Add "transfer_counts", Array(Null, 0)
    matrix = RoutingTransfersMatrix(root, values)
    Check matrix(1, 2) = -1 And matrix(2, 2) = 0, "null is not a known direct flight", passed
    root("transfer_counts") = Array(True, 0)
    On Error Resume Next
    matrix = RoutingTransfersMatrix(root, values): failed = Err.Number <> 0: Err.Clear
    On Error GoTo Failed
    Check failed, "boolean transfer count rejected", passed
    root("transfer_counts") = Array(0)
    On Error Resume Next
    matrix = RoutingTransfersMatrix(root, values): failed = Err.Number <> 0: Err.Clear
    On Error GoTo Failed
    Check failed, "metadata length mismatch rejected", passed
    root("transfer_counts") = Array(0, 1): values(2, pcQuoteId) = "a"
    On Error Resume Next
    matrix = RoutingTransfersMatrix(root, values): failed = Err.Number <> 0: Err.Clear
    On Error GoTo Failed
    Check failed, "one quote cannot have conflicting counts", passed
    RunRoutingRegression = passed
    Exit Function
Failed:
    RunRoutingRegression = "FAIL: " & Err.Description
End Function
