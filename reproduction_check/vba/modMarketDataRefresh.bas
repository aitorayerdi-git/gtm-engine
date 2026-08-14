Attribute VB_Name = "modMarketDataRefresh"
Option Explicit

Public Sub Actualizar_MarketView()
    Dim oldCalc As XlCalculation, latestDate As Variant, errorMessage As String
    On Error GoTo ErrHandler
    oldCalc = Application.Calculation
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = "Actualizando MarketView..."
    ThisWorkbook.Worksheets("Historical_DA").UsedRange.Calculate
    Application.CalculateUntilAsyncQueriesDone
    DoEvents
    latestDate = LatestMarketViewDate()
    SetSourceControl "MarketView", Now, latestDate, MarketStatusText(latestDate)
CleanExit:
    On Error Resume Next
    Application.Calculation = oldCalc
    Application.StatusBar = False
    On Error GoTo 0
    If Len(errorMessage) > 0 Then
        MsgBox "No se pudo actualizar MarketView." & vbCrLf & errorMessage, vbCritical
    Else
        MsgBox "MarketView: " & DateResultText(latestDate), _
            IIf(MarketStatusText(latestDate) = "OK", vbInformation, vbExclamation)
    End If
    Exit Sub
ErrHandler:
    errorMessage = "Error " & Err.Number & ": " & Err.Description
    Resume CleanExit
End Sub

Public Sub Comprobar_Reuters_MarketView()
    Dim reutersDate As Variant, marketViewDate As Variant
    On Error GoTo ErrHandler
    reutersDate = LatestReutersDate()
    marketViewDate = LatestMarketViewDate()
    SetSourceControl "Reuters", Now, reutersDate, MarketStatusText(reutersDate)
    SetSourceControl "MarketView", Now, marketViewDate, MarketStatusText(marketViewDate)
    MsgBox "Reuters: " & DateResultText(reutersDate) & vbCrLf & _
        "MarketView: " & DateResultText(marketViewDate), vbInformation
    Exit Sub
ErrHandler:
    MsgBox "No se pudieron comprobar Reuters y MarketView." & vbCrLf & _
        "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub

' Backward-compatible entry points for existing assigned buttons.
Public Sub Actualizar_Reuters()
    Comprobar_Reuters_MarketView
End Sub

Public Sub Comprobar_MarketView()
    Comprobar_Reuters_MarketView
End Sub

Public Sub Actualizar_Control_Datos_Mercado()
    Comprobar_Reuters_MarketView
End Sub

Public Sub Inicializar_Control_Datos_Mercado()
    Dim value As Variant
    value = LatestReutersDate()
    SetSourceControl "Reuters", Empty, value, MarketStatusText(value)
    value = LatestMarketViewDate()
    SetSourceControl "MarketView", Empty, value, MarketStatusText(value)
End Sub

Public Function MarketControlDate(ByVal fieldName As String) As Date
    Dim value As Variant
    value = TableLookupValue("tblManualDates", "Field", fieldName, "Date")
    If Not IsDate(value) Then Err.Raise vbObjectError + 2401, , _
        fieldName & " no contiene una fecha valida en tblManualDates."
    MarketControlDate = DateValue(CDate(value))
End Function

Public Function MarketStatusText(ByVal latestDate As Variant) As String
    Dim targetDate As Date
    targetDate = MarketControlDate("Historical End Date")
    If Not IsDate(latestDate) Then
        MarketStatusText = "CHECK - NO VALID DATE"
    ElseIf DateValue(CDate(latestDate)) < targetDate Then
        MarketStatusText = "CHECK - DATA BEHIND TARGET"
    Else
        MarketStatusText = "OK"
    End If
End Function

Public Sub SetSourceControl( _
    ByVal sourceName As String, ByVal actionStamp As Variant, _
    ByVal latestDate As Variant, ByVal statusText As String)
    Dim table As ListObject, rowIndex As Long
    Set table = FindTable("tblMarketDataStatus")
    rowIndex = FindTableRow(table, "Source", sourceName)
    If rowIndex = 0 Then Err.Raise vbObjectError + 2402, , _
        "Falta Source '" & sourceName & "' en tblMarketDataStatus."
    If IsDate(actionStamp) Then _
        table.DataBodyRange.Cells(rowIndex, TableColumn(table, "Last Successful Action At")).Value = CDate(actionStamp)
    If IsDate(latestDate) Then
        table.DataBodyRange.Cells(rowIndex, TableColumn(table, "Latest Data Date")).Value = DateValue(CDate(latestDate))
    Else
        table.DataBodyRange.Cells(rowIndex, TableColumn(table, "Latest Data Date")).ClearContents
    End If
    table.DataBodyRange.Cells(rowIndex, TableColumn(table, "Status")).Value = statusText
    table.ListColumns("Last Successful Action At").DataBodyRange.NumberFormat = "dd/mm/yyyy hh:mm:ss"
    table.ListColumns("Latest Data Date").DataBodyRange.NumberFormat = "dd/mm/yyyy"
End Sub

Public Sub SetSourceStatusOnly(ByVal sourceName As String, ByVal statusText As String)
    Dim table As ListObject, rowIndex As Long
    Set table = FindTable("tblMarketDataStatus")
    rowIndex = FindTableRow(table, "Source", sourceName)
    If rowIndex = 0 Then Err.Raise vbObjectError + 2406, , _
        "Falta Source '" & sourceName & "' en tblMarketDataStatus."
    table.DataBodyRange.Cells(rowIndex, TableColumn(table, "Status")).Value = statusText
End Sub

Public Function SourceControlStatus(ByVal sourceName As String) As String
    SourceControlStatus = CStr(TableLookupValue( _
        "tblMarketDataStatus", "Source", sourceName, "Status"))
End Function

Public Function TableLookupValue( _
    ByVal tableName As String, ByVal keyHeader As String, _
    ByVal keyValue As String, ByVal resultHeader As String) As Variant
    Dim table As ListObject, rowIndex As Long
    Set table = FindTable(tableName)
    rowIndex = FindTableRow(table, keyHeader, keyValue)
    If rowIndex = 0 Then Err.Raise vbObjectError + 2403, , _
        "No se encontro '" & keyValue & "' en " & tableName & "."
    TableLookupValue = table.DataBodyRange.Cells(rowIndex, TableColumn(table, resultHeader)).Value
End Function

Private Function FindTable(ByVal tableName As String) As ListObject
    Dim ws As Worksheet, table As ListObject
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        Set table = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not table Is Nothing Then Set FindTable = table: Exit Function
    Next ws
    Err.Raise vbObjectError + 2404, , "No existe la tabla " & tableName & "."
End Function

Private Function TableColumn(ByVal table As ListObject, ByVal headerName As String) As Long
    On Error Resume Next
    TableColumn = table.ListColumns(headerName).Index
    On Error GoTo 0
    If TableColumn = 0 Then Err.Raise vbObjectError + 2405, , _
        "Falta la columna '" & headerName & "' en " & table.Name & "."
End Function

Private Function FindTableRow( _
    ByVal table As ListObject, ByVal keyHeader As String, ByVal keyValue As String) As Long
    Dim rowIndex As Long, columnIndex As Long
    columnIndex = TableColumn(table, keyHeader)
    If table.DataBodyRange Is Nothing Then Exit Function
    For rowIndex = 1 To table.DataBodyRange.Rows.Count
        If StrComp(Trim$(CStr(table.DataBodyRange.Cells(rowIndex, columnIndex).Value)), _
            keyValue, vbTextCompare) = 0 Then
            FindTableRow = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

Private Function DateResultText(ByVal latestDate As Variant) As String
    If IsDate(latestDate) Then
        DateResultText = Format(CDate(latestDate), "dd/mm/yyyy") & " - " & MarketStatusText(latestDate)
    Else
        DateResultText = "sin fecha valida - CHECK"
    End If
End Function

Public Function LatestReutersDate() As Variant
    ' Reuters status represents the cached provider curves. CURVE PRICES is
    ' rebuilt independently in the value-only snapshot before every output.
    LatestReutersDate = LatestReutersProviderDate()
End Function

Public Function LatestReutersProviderDate() As Variant
    Dim result As Variant
    result = LatestNumericDate(ThisWorkbook.Worksheets("TTF"), 4, 5, 42)
    result = EarlierDate(result, LatestNumericDate(ThisWorkbook.Worksheets("Brent Dated"), 4, 5, 42))
    result = EarlierDate(result, LatestNumericDate(ThisWorkbook.Worksheets("HH"), 4, 5, 42))
    result = EarlierDate(result, LatestNumericDate(ThisWorkbook.Worksheets("EURF"), 4, 5, 5))
    LatestReutersProviderDate = result
End Function

Public Function LatestNormalizedCurveDate() As Variant
    Dim table As ListObject, cell As Range, result As Date, found As Boolean
    Set table = FindTable("tblCurvePrices")
    If table.DataBodyRange Is Nothing Then Exit Function
    For Each cell In table.ListColumns("Market Date").DataBodyRange.Cells
        If IsDate(cell.Value) Then
            If Not found Or CDate(cell.Value) > result Then result = CDate(cell.Value): found = True
        End If
    Next cell
    If found Then LatestNormalizedCurveDate = result Else LatestNormalizedCurveDate = Empty
End Function

Private Function LatestMarketViewDate() As Variant
    Dim result As Variant, ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Historical_DA")
    result = LatestNumericDate(ws, 1, 2, 8)
    result = EarlierDate(result, LatestNumericDate(ws, 16, 17, 17))
    result = EarlierDate(result, LatestNumericDate(ws, 23, 24, 24))
    LatestMarketViewDate = result
End Function

Private Function LatestNumericDate( _
    ByVal ws As Worksheet, ByVal dateColumn As Long, _
    ByVal firstValueColumn As Long, ByVal lastValueColumn As Long) As Variant
    Dim rowNumber As Long, columnNumber As Long, lastRow As Long
    Dim candidateDate As Date, maximumDate As Date, foundDate As Boolean
    lastRow = ws.Cells(ws.Rows.Count, dateColumn).End(xlUp).Row
    For rowNumber = 1 To lastRow
        If IsDate(ws.Cells(rowNumber, dateColumn).Value) Then
            For columnNumber = firstValueColumn To lastValueColumn
                If IsNumeric(ws.Cells(rowNumber, columnNumber).Value) And _
                    Len(CStr(ws.Cells(rowNumber, columnNumber).Value)) > 0 Then
                    candidateDate = CDate(ws.Cells(rowNumber, dateColumn).Value)
                    If Not foundDate Or candidateDate > maximumDate Then
                        maximumDate = candidateDate
                        foundDate = True
                    End If
                    Exit For
                End If
            Next columnNumber
        End If
    Next rowNumber
    If foundDate Then LatestNumericDate = maximumDate Else LatestNumericDate = Empty
End Function

Private Function EarlierDate(ByVal firstDate As Variant, ByVal secondDate As Variant) As Variant
    If Not IsDate(firstDate) Or Not IsDate(secondDate) Then
        EarlierDate = Empty
    ElseIf CDate(firstDate) <= CDate(secondDate) Then
        EarlierDate = CDate(firstDate)
    Else
        EarlierDate = CDate(secondDate)
    End If
End Function
