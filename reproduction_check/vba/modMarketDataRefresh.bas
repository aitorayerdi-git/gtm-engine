Attribute VB_Name = "modMarketDataRefresh"
Option Explicit

Public Sub Actualizar_Reuters_MarketView()
    Dim oldCalc As XlCalculation
    Dim oldEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim errorMessage As String

    On Error GoTo ErrHandler
    oldCalc = Application.Calculation
    oldEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationAutomatic

    ThisWorkbook.Worksheets("TTF").Calculate
    ThisWorkbook.Worksheets("Brent Dated").Calculate
    ThisWorkbook.Worksheets("HH").Calculate
    ThisWorkbook.Worksheets("EURF").Calculate
    ThisWorkbook.Worksheets("Historical_DA").Calculate
    Application.CalculateUntilAsyncQueriesDone
    UpdateMarketDataControl True

CleanExit:
    On Error Resume Next
    Application.Calculation = oldCalc
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreenUpdating
    On Error GoTo 0
    If Len(errorMessage) > 0 Then
        MsgBox "No se pudo completar la actualización Reuters / MarketView." & vbCrLf & _
            errorMessage, vbCritical
    Else
        MsgBox "Reuters y MarketView actualizados. Revisa MARKET DATA STATUS en MANUAL CHANGES.", _
            vbInformation
    End If
    Exit Sub

ErrHandler:
    errorMessage = "Error " & Err.Number & ": " & Err.Description
    Resume CleanExit
End Sub

Public Sub Actualizar_Control_Datos_Mercado()
    UpdateMarketDataControl False
    MsgBox "Control de Reuters y MarketView recalculado sin modificar la fecha/hora de refresh.", _
        vbInformation
End Sub

Public Sub Inicializar_Control_Datos_Mercado()
    UpdateMarketDataControl False
End Sub

Private Sub UpdateMarketDataControl(ByVal stampRefresh As Boolean)
    Dim ws As Worksheet
    Dim reutersDate As Variant
    Dim marketViewDate As Variant
    Set ws = ThisWorkbook.Worksheets("MANUAL CHANGES")
    reutersDate = LatestReutersDate()
    marketViewDate = LatestMarketViewDate()

    If stampRefresh Then
        ws.Range("Y4").Value = Now
        ws.Range("Y5").Value = Now
    End If
    ws.Range("Z4").Value = reutersDate
    ws.Range("Z5").Value = marketViewDate
    ws.Range("AA4").Value = StatusText(reutersDate)
    ws.Range("AA5").Value = StatusText(marketViewDate)
    ws.Range("Y4:Y5").NumberFormat = "dd/mm/yyyy hh:mm:ss"
    ws.Range("Z4:Z5").NumberFormat = "dd/mm/yyyy"
End Sub

Private Function LatestReutersDate() As Variant
    Dim result As Variant
    result = LatestNumericDate(ThisWorkbook.Worksheets("TTF"), 4, 5, 42)
    result = EarlierDate(result, LatestNumericDate(ThisWorkbook.Worksheets("Brent Dated"), 4, 5, 42))
    result = EarlierDate(result, LatestNumericDate(ThisWorkbook.Worksheets("HH"), 4, 5, 42))
    result = EarlierDate(result, LatestNumericDate(ThisWorkbook.Worksheets("EURF"), 4, 5, 5))
    LatestReutersDate = result
End Function

Private Function LatestMarketViewDate() As Variant
    Dim result As Variant
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Historical_DA")
    result = LatestNumericDate(ws, 1, 2, 8)
    result = EarlierDate(result, LatestNumericDate(ws, 16, 17, 17))
    result = EarlierDate(result, LatestNumericDate(ws, 23, 24, 24))
    LatestMarketViewDate = result
End Function

Private Function LatestNumericDate( _
    ByVal ws As Worksheet, ByVal dateColumn As Long, _
    ByVal firstValueColumn As Long, ByVal lastValueColumn As Long) As Variant
    Dim rowNumber As Long
    Dim columnNumber As Long
    Dim lastRow As Long
    Dim candidateDate As Date
    Dim maximumDate As Date
    Dim foundDate As Boolean
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
    If foundDate Then
        LatestNumericDate = maximumDate
    Else
        LatestNumericDate = Empty
    End If
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

Private Function StatusText(ByVal latestDate As Variant) As String
    Dim targetDate As Variant
    targetDate = ThisWorkbook.Worksheets("MANUAL CHANGES").Range("P6").Value
    If Not IsDate(latestDate) Then
        StatusText = "CHECK - NO VALID DATE"
    ElseIf IsDate(targetDate) And CDate(latestDate) < CDate(targetDate) Then
        StatusText = "CHECK - DATA BEHIND TARGET"
    Else
        StatusText = "OK"
    End If
End Function
