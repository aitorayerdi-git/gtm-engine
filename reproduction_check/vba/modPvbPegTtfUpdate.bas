Attribute VB_Name = "modPvbPegTtfUpdate"
Option Explicit

Public Sub Actualizar_PVB_PEG_TTF()
    Const LOCATIONSP_URL As String = _
        "https://cepsacorp.sharepoint.com/sites/Shared-CGC2/Documentos compartidos/CGC2 Operaciones/Aprovisionamiento & Trading/Cotizaciones/PVBTTF from PRICER/LOCATIONSP.xlsx"

    Dim wbLocation As Workbook
    Dim wsPVB As Worksheet
    Dim wsPEG As Worksheet
    Dim oldCalc As XlCalculation
    Dim oldEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldAskToUpdateLinks As Boolean
    Dim lastRowPVB As Long
    Dim lastRowPEG As Long
    Dim lastColumnPVB As Long
    Dim lastColumnPEG As Long
    Dim lastDatePVB As Variant
    Dim lastDatePEG As Variant
    Dim errorMessage As String
    Dim currentStage As String

    On Error GoTo ErrHandler

    currentStage = "preparando las hojas de curvas"
    Set wsPVB = ThisWorkbook.Worksheets("PVB-TTF")
    Set wsPEG = ThisWorkbook.Worksheets("PEG-TTF")
    oldCalc = Application.Calculation
    oldEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    oldDisplayAlerts = Application.DisplayAlerts
    oldAskToUpdateLinks = Application.AskToUpdateLinks
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    Application.Calculation = xlCalculationManual

    lastRowPVB = wsPVB.Cells(wsPVB.Rows.Count, "A").End(xlUp).Row
    lastRowPEG = wsPEG.Cells(wsPEG.Rows.Count, "A").End(xlUp).Row
    lastColumnPVB = wsPVB.UsedRange.Column + wsPVB.UsedRange.Columns.Count - 1
    lastColumnPEG = wsPEG.UsedRange.Column + wsPEG.UsedRange.Columns.Count - 1
    currentStage = "abriendo LOCATIONSP.xlsx desde SharePoint"
    Set wbLocation = Workbooks.Open( _
        Filename:=LOCATIONSP_URL, ReadOnly:=True, UpdateLinks:=0, _
        AddToMru:=False, Notify:=False)

    currentStage = "recalculando PVB-TTF"
    wsPVB.Range(wsPVB.Cells(1, 3), wsPVB.Cells(lastRowPVB, lastColumnPVB)).Calculate
    currentStage = "recalculando PEG-TTF"
    wsPEG.Range(wsPEG.Cells(1, 3), wsPEG.Cells(lastRowPEG, lastColumnPEG)).Calculate
    PopulatePreviousMarketDates wsPVB, lastRowPVB
    PopulatePreviousMarketDates wsPEG, lastRowPEG
    currentStage = "cerrando LOCATIONSP.xlsx"
    wbLocation.Close SaveChanges:=False
    Set wbLocation = Nothing

    lastDatePVB = LastDateWithValueInAL(wsPVB)
    lastDatePEG = LastDateWithValueInAL(wsPEG)

CleanExit:
    On Error Resume Next
    If Not wbLocation Is Nothing Then wbLocation.Close SaveChanges:=False
    Application.Calculation = oldCalc
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreenUpdating
    Application.DisplayAlerts = oldDisplayAlerts
    Application.AskToUpdateLinks = oldAskToUpdateLinks
    If Len(errorMessage) = 0 Then
        SetSourceControl "PVB-TTF", Now, lastDatePVB, MarketStatusText(lastDatePVB)
        SetSourceControl "PEG-TTF", Now, lastDatePEG, MarketStatusText(lastDatePEG)
    Else
        SetSourceStatusOnly "PVB-TTF", "ERROR - " & errorMessage
        SetSourceStatusOnly "PEG-TTF", "ERROR - " & errorMessage
    End If
    On Error GoTo 0

    If Len(errorMessage) > 0 Then
        MsgBox "No se pudieron actualizar las curvas PVB/TTF y PEG/TTF." & vbCrLf & _
            errorMessage, vbCritical
    ElseIf IsDate(lastDatePVB) And IsDate(lastDatePEG) Then
        If CDate(lastDatePVB) = CDate(lastDatePEG) Then
            MsgBox "Curvas PVB/TTF y PEG/TTF actualizadas hasta Market Date " & _
                Format(CDate(lastDatePVB), "dd-mm-yy"), vbInformation
        Else
            MsgBox "Las curvas se actualizaron, pero sus últimos Market Dates no coinciden." & _
                vbCrLf & "PVB/TTF: " & Format(CDate(lastDatePVB), "dd-mm-yy") & _
                vbCrLf & "PEG/TTF: " & Format(CDate(lastDatePEG), "dd-mm-yy"), vbExclamation
        End If
    Else
        MsgBox "Las curvas se recalcularon, pero no se pudo identificar su último Market Date.", _
            vbExclamation
    End If
    Exit Sub

ErrHandler:
    errorMessage = "Fase: " & currentStage & vbCrLf & _
        "Error " & Err.Number & ": " & Err.Description
    Resume CleanExit
End Sub

Public Sub Rebuild_PVB_PEG_Market_Dates()
    Dim wsPVB As Worksheet, wsPEG As Worksheet
    Set wsPVB = ThisWorkbook.Worksheets("PVB-TTF")
    Set wsPEG = ThisWorkbook.Worksheets("PEG-TTF")
    PopulatePreviousMarketDates wsPVB, wsPVB.Cells(wsPVB.Rows.Count, "A").End(xlUp).Row
    PopulatePreviousMarketDates wsPEG, wsPEG.Cells(wsPEG.Rows.Count, "A").End(xlUp).Row
End Sub

Private Sub PopulatePreviousMarketDates(ByVal targetSheet As Worksheet, ByVal lastRow As Long)
    Dim calendarTable As ListObject, calendarDates As Range, marketFlags As Range
    Dim targetRow As Long, calendarRow As Long, curvesSent As Date
    Dim candidate As Date, previousDate As Date, found As Boolean
    Dim sentValues As Variant, dateValues As Variant, flagValues As Variant, outputValues() As Variant
    Set calendarTable = ThisWorkbook.Worksheets("MARKET CALENDAR").ListObjects("tblMarketCalendar")
    Set calendarDates = calendarTable.ListColumns("Date").DataBodyRange
    Set marketFlags = calendarTable.ListColumns("Is Market Day").DataBodyRange
    sentValues = targetSheet.Range("A3:A" & lastRow).Value2
    dateValues = calendarDates.Value2
    flagValues = marketFlags.Value2
    ReDim outputValues(1 To UBound(sentValues, 1), 1 To 1)
    For targetRow = 1 To UBound(sentValues, 1)
        found = False
        If IsNumeric(sentValues(targetRow, 1)) Then
            curvesSent = DateValue(CDate(sentValues(targetRow, 1)))
            For calendarRow = UBound(dateValues, 1) To 1 Step -1
                If IsNumeric(dateValues(calendarRow, 1)) Then
                    candidate = DateValue(CDate(dateValues(calendarRow, 1)))
                    If candidate < curvesSent And CBool(flagValues(calendarRow, 1)) Then
                        previousDate = candidate
                        found = True
                        Exit For
                    End If
                End If
            Next calendarRow
        End If
        If found Then
            outputValues(targetRow, 1) = previousDate
        Else
            outputValues(targetRow, 1) = Empty
        End If
    Next targetRow
    targetSheet.Range("B3:B" & lastRow).Value = outputValues
    targetSheet.Range("B3:B" & lastRow).NumberFormat = "dd/mm/yyyy"
End Sub

Private Function LastDateWithValueInAL(ByVal ws As Worksheet) As Variant
    Dim rowNumber As Long
    For rowNumber = ws.Cells(ws.Rows.Count, "AL").End(xlUp).Row To 1 Step -1
        If Len(Trim$(CStr(ws.Cells(rowNumber, "AL").Value))) > 0 Then
            ' Column A is the source observation date. Column B depends on the
            ' separate Foto FO workbook and must not determine refresh success.
            LastDateWithValueInAL = ws.Cells(rowNumber, "A").Value
            Exit Function
        End If
    Next rowNumber
    LastDateWithValueInAL = Empty
End Function
