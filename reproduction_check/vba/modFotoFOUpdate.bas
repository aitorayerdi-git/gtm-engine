Attribute VB_Name = "modFotoFOUpdate"
Option Explicit

Private Const FOTO_URL As String = "https://cepsacorp.sharepoint.com/sites/Shared-CGC2/Documentos compartidos/CGC2 Operaciones/Operativa Diaria - OPS/Balance/Operativa_diaria/Foto FO.xlsx"
Private Const SH_MANUAL As String = "MANUAL CHANGES"
Private Const SH_FLOWS As String = "OPERATING FLOWS"
Private Const SH_COSTS As String = "COSTS"
Private Const SH_LOG As String = "UPDATE LOG"
Private Const SH_CALENDAR As String = "MARKET CALENDAR"
Private Const SH_BOOKS As String = "BOOKS"
Private Const SH_STATE As String = "FOTO FO STATE"
Private mAppliedDateCache As Object

Public Sub UpdateFotoFO()
    UpdateFotoFOCore FOTO_URL, False
End Sub

Public Sub UpdateFotoFO_Automated()
    UpdateFotoFOCore FOTO_URL, True
End Sub

Public Sub UpdateFotoFO_LocalTest()
    Dim localPath As String
    localPath = ThisWorkbook.Path & "\..\verification\Foto FO.xlsx"
    UpdateFotoFOCore localPath, True
End Sub

Private Sub UpdateFotoFOCore(ByVal sourcePath As String, ByVal localTest As Boolean)
    Dim sourceBook As Workbook, openedHere As Boolean
    Dim targetDate As Date, previousDate As Date, startedAt As Date
    Dim oldCalc As XlCalculation, oldEvents As Boolean, oldScreen As Boolean, oldAlerts As Boolean
    Dim flowValues As Object, activeBooks As Object, book As Variant
    Dim sourceLabel As String, statusText As String, detailText As String
    Dim rowCount As Long, nonZeroCount As Long

    startedAt = Now
    oldCalc = Application.Calculation
    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    oldAlerts = Application.DisplayAlerts
    On Error GoTo Failed

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Actualizando Foto FO..."
    Set mAppliedDateCache = Nothing

    targetDate = ReadLastMarketDate()
    previousDate = PreviousMarketDate(targetDate)
    Set activeBooks = ReadActiveBooks()
    If activeBooks.Count = 0 Then Err.Raise vbObjectError + 2101, , "BOOKS no contiene BOOK activos."
    ValidateMissedMarketDates targetDate, activeBooks

    Set sourceBook = OpenFotoFO(sourcePath, openedHere)
    sourceLabel = sourceBook.FullName
    ValidateSource sourceBook, targetDate
    Set flowValues = CalculateFlows(sourceBook, targetDate, previousDate, activeBooks)
    ValidateCalculatedFlows flowValues, activeBooks
    PublishFlows flowValues, targetDate, activeBooks, rowCount, nonZeroCount
    PublishCosts flowValues, targetDate, activeBooks
    SaveState flowValues, targetDate, activeBooks

    statusText = "OK"
    detailText = "Carga completa: " & rowCount & " BOOK; " & nonZeroCount & " con importe distinto de cero."
    WriteAudit startedAt, Now, targetDate, sourceLabel, statusText, detailText, rowCount, activeBooks.Count
    UpdateManualStatus Now, targetDate, sourceLabel, statusText, detailText

    If openedHere Then sourceBook.Close SaveChanges:=False
    RestoreExcel oldCalc, oldEvents, oldScreen, oldAlerts
    If Not localTest Then MsgBox "Foto FO actualizada correctamente para Market Date " & Format(targetDate, "dd/mm/yyyy") & "." & vbCrLf & detailText, vbInformation, "Foto FO - OK"
    Exit Sub

Failed:
    statusText = "ERROR"
    detailText = "Error " & Err.Number & ": " & Err.Description
    On Error Resume Next
    If Not sourceBook Is Nothing Then If openedHere Then sourceBook.Close SaveChanges:=False
    WriteAudit startedAt, Now, targetDate, sourceLabel, statusText, detailText, rowCount, IIf(activeBooks Is Nothing, 0, activeBooks.Count)
    UpdateManualStatus Now, targetDate, sourceLabel, statusText, detailText
    RestoreExcel oldCalc, oldEvents, oldScreen, oldAlerts
    On Error GoTo 0
    If Not localTest Then MsgBox "No se ha publicado la actualización." & vbCrLf & detailText, vbCritical, "Foto FO - ERROR"
End Sub

Private Function OpenFotoFO(ByVal sourcePath As String, ByRef openedHere As Boolean) As Workbook
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If InStr(1, wb.Name, "Foto FO", vbTextCompare) > 0 And wb.Name <> ThisWorkbook.Name Then
            Set OpenFotoFO = wb
            openedHere = False
            Exit Function
        End If
    Next wb
    Set OpenFotoFO = Workbooks.Open(Filename:=sourcePath, UpdateLinks:=0, ReadOnly:=True)
    openedHere = True
End Function

Private Sub ValidateMissedMarketDates(ByVal targetDate As Date, ByVal books As Object)
    Dim stateDate As Date, book As Variant, bookState As Variant
    Dim ws As Worksheet, lastRow As Long, r As Long, candidate As Date
    Dim missingDates As String

    For Each book In books.Keys
        bookState = ReadState(CStr(book))
        If stateDate = 0 Then
            stateDate = CDate(bookState(0))
        ElseIf DateValue(CDate(bookState(0))) <> stateDate Then
            Err.Raise vbObjectError + 2105, , "FOTO FO STATE contiene fechas distintas por BOOK. No se puede determinar la cobertura manual."
        End If
    Next book
    ValidateBaseline targetDate, stateDate
    If BaselineDate() > stateDate Then stateDate = BaselineDate()
    If stateDate >= targetDate Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(SH_CALENDAR)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 5 To lastRow
        If IsDate(ws.Cells(r, 1).Value) And CBool(ws.Cells(r, 2).Value) Then
            candidate = DateValue(ws.Cells(r, 1).Value)
            If candidate > stateDate And candidate < targetDate Then
                If Not HasValidManualCostRow(candidate) Then
                    If Len(missingDates) > 0 Then missingDates = missingDates & ", "
                    missingDates = missingDates & Format$(candidate, "dd/mm/yyyy")
                End If
            End If
        End If
    Next r

    If Len(missingDates) > 0 Then
        Err.Raise vbObjectError + 2106, , "Hay Market Dates sin actualizar: " & missingDates & ". Antes de continuar, cree en COSTS una fila por fecha, seleccione SOURCE=MANUAL e introduzca el delta diario. Si el delta fue cero, deje B:P en cero y documentelo en COMMENT."
    End If
End Sub

Private Sub ValidateBaseline(ByVal targetDate As Date, ByVal savedStateDate As Date)
    Dim ws As Worksheet, lastRow As Long, r As Long, count As Long, baseline As Date, c As Long
    Set ws = ThisWorkbook.Worksheets(SH_COSTS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 5 To lastRow
        If StrComp(Trim$(CStr(ws.Cells(r, 18).Value)), "BASELINE", vbTextCompare) = 0 Then
            count = count + 1
            If Not IsDate(ws.Cells(r, 1).Value) Then Err.Raise vbObjectError + 2109, , "La fila BASELINE necesita una Market Date valida."
            baseline = DateValue(ws.Cells(r, 1).Value)
            If baseline >= targetDate Then Err.Raise vbObjectError + 2117, , "BASELINE debe ser anterior a Last Market Date."
            For c = 2 To 16
                If Len(Trim$(CStr(ws.Cells(r, c).Value))) > 0 And Not IsNumeric(ws.Cells(r, c).Value) Then Err.Raise vbObjectError + 2118, , "BASELINE contiene un importe no numerico."
            Next c
            If Len(Trim$(CStr(ws.Cells(r, 21).Value))) = 0 Then Err.Raise vbObjectError + 2119, , "BASELINE necesita un COMMENT explicativo."
        End If
    Next r
    If count > 1 Then Err.Raise vbObjectError + 2123, , "COSTS solo puede contener una fila BASELINE."
End Sub

Private Function BaselineDate() As Date
    Dim ws As Worksheet, lastRow As Long, r As Long
    Set ws = ThisWorkbook.Worksheets(SH_COSTS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 5 To lastRow
        If StrComp(Trim$(CStr(ws.Cells(r, 18).Value)), "BASELINE", vbTextCompare) = 0 And IsDate(ws.Cells(r, 1).Value) Then
            BaselineDate = DateValue(ws.Cells(r, 1).Value)
            Exit Function
        End If
    Next r
End Function

Private Function HasValidManualCostRow(ByVal marketDate As Date) As Boolean
    Dim ws As Worksheet, lastRow As Long, r As Long, c As Long
    Set ws = ThisWorkbook.Worksheets(SH_COSTS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 5 To lastRow
        If IsDate(ws.Cells(r, 1).Value) Then
            If DateValue(ws.Cells(r, 1).Value) = marketDate And StrComp(Trim$(CStr(ws.Cells(r, 18).Value)), "MANUAL", vbBinaryCompare) = 0 Then
                For c = 2 To 16
                    If Len(Trim$(CStr(ws.Cells(r, c).Value))) > 0 And Not IsNumeric(ws.Cells(r, c).Value) Then
                        Err.Raise vbObjectError + 2107, , "COSTS contiene un importe no numerico para " & Format$(marketDate, "dd/mm/yyyy") & "."
                    End If
                Next c
                If Len(Trim$(CStr(ws.Cells(r, 21).Value))) = 0 Then
                    Err.Raise vbObjectError + 2108, , "La fila MANUAL de " & Format$(marketDate, "dd/mm/yyyy") & " necesita un COMMENT explicativo."
                End If
                HasValidManualCostRow = True
                Exit Function
            End If
        End If
    Next r
End Function

Private Function ReadLastMarketDate() As Date
    Dim value As Variant
    value = ThisWorkbook.Worksheets(SH_MANUAL).Range("F5").Value
    If Not IsDate(value) Then Err.Raise vbObjectError + 2102, , "DAILY-001 / Last Market Date no contiene una fecha válida."
    ReadLastMarketDate = DateValue(CDate(value))
    If Not IsConfiguredMarketDate(ReadLastMarketDate) Then Err.Raise vbObjectError + 2103, , "Last Market Date no está marcado como Market Day en MARKET CALENDAR."
End Function

Private Function IsConfiguredMarketDate(ByVal targetDate As Date) As Boolean
    Dim ws As Worksheet, lastRow As Long, r As Long
    Set ws = ThisWorkbook.Worksheets(SH_CALENDAR)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 5 To lastRow
        If IsDate(ws.Cells(r, 1).Value) Then
            If DateValue(ws.Cells(r, 1).Value) = targetDate Then
                IsConfiguredMarketDate = CBool(ws.Cells(r, 2).Value)
                Exit Function
            End If
        End If
    Next r
End Function

Private Function PreviousMarketDate(ByVal targetDate As Date) As Date
    Dim ws As Worksheet, lastRow As Long, r As Long, candidate As Date
    Set ws = ThisWorkbook.Worksheets(SH_CALENDAR)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 5 To lastRow
        If IsDate(ws.Cells(r, 1).Value) And CBool(ws.Cells(r, 2).Value) Then
            candidate = DateValue(ws.Cells(r, 1).Value)
            If candidate < targetDate And candidate > PreviousMarketDate Then PreviousMarketDate = candidate
        End If
    Next r
    If PreviousMarketDate = 0 Then Err.Raise vbObjectError + 2104, , "No existe un Market Date anterior configurado."
End Function

Private Function ReadActiveBooks() As Object
    Dim result As Object, ws As Worksheet, lastRow As Long, r As Long, name As String
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    Set ws = ThisWorkbook.Worksheets(SH_BOOKS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 5 To lastRow
        name = Trim$(CStr(ws.Cells(r, 1).Value))
        If Len(name) > 0 And CBool(ws.Cells(r, 2).Value) Then result(name) = True
    Next r
    Set ReadActiveBooks = result
End Function

Private Sub ValidateSource(ByVal wb As Workbook, ByVal targetDate As Date)
    Dim required As Variant, item As Variant
    required = Array("Canones", "Delta de costes", "Optimizaciones", "Index replication", "MAIN")
    For Each item In required
        If Not SheetExists(CStr(item), wb) Then Err.Raise vbObjectError + 2110, , "Falta la hoja obligatoria '" & item & "' en Foto FO."
    Next item
    RequireHeader wb.Worksheets("Canones"), 1, 10, "Importe Total"
    RequireHeader wb.Worksheets("Optimizaciones"), 1, 3, "Trade date"
    RequireHeader wb.Worksheets("Optimizaciones"), 1, 6, "Tradebook"
    RequireHeader wb.Worksheets("Optimizaciones"), 1, 10, "Importe"
    RequireHeader wb.Worksheets("Optimizaciones"), 1, 12, "Strategy"
    RequireHeader wb.Worksheets("Index replication"), 1, 1, "Fecha delivery"
    RequireHeader wb.Worksheets("MAIN"), 1, 1, "Trade Date"
    If Not YearExists(wb.Worksheets("Delta de costes"), Year(targetDate)) Then Err.Raise vbObjectError + 2111, , "Delta de costes no contiene el año de Last Market Date."
End Sub

Private Function CalculateFlows(ByVal wb As Workbook, ByVal targetDate As Date, ByVal previousDate As Date, ByVal books As Object) As Object
    Dim result As Object, book As Variant, values As Variant
    Dim snapshots As Object, previousState As Variant, stateDate As Date
    Dim currentSnapshot As Double, feeAmount As Double, optim As Double, replication As Double
    Dim manualCost As Double, manualFees As Double, baselineCost As Double, baselineFees As Double
    Set result = CreateObject("Scripting.Dictionary"): result.CompareMode = vbTextCompare
    Set snapshots = LogisticsSnapshots(wb.Worksheets("Delta de costes"))
    For Each book In books.Keys
        values = Array(0#, 0#, 0#)
        previousState = ReadState(CStr(book))
        stateDate = CDate(previousState(0))
        If BaselineDate() > 0 And (BaselineDate() > stateDate Or IsFirstAutomaticDateAfterBaseline(targetDate)) Then
            stateDate = BaselineDate()
            previousState = Array(stateDate, 0#, 0#, 0#)
            baselineCost = CostRowAdjustment(CStr(book), "BASELINE", stateDate, False)
            baselineFees = CostRowAdjustment(CStr(book), "BASELINE", stateDate, True)
        Else
            baselineCost = 0#: baselineFees = 0#
        End If
        If stateDate > targetDate Then Err.Raise vbObjectError + 2116, , "El estado de Foto FO es posterior a Last Market Date."
        currentSnapshot = 0#
        If snapshots.Exists(CStr(book)) Then currentSnapshot = CDbl(snapshots(CStr(book)))
        feeAmount = FeeSnapshot(wb.Worksheets("Canones"), CStr(book))
        optim = OptimizationFlow(wb.Worksheets("Optimizaciones"), CStr(book), previousDate, targetDate)
        replication = ReplicationFlow(wb, CStr(book), targetDate)
        manualCost = ManualCostAdjustment(CStr(book), stateDate, targetDate, False)
        manualFees = ManualCostAdjustment(CStr(book), stateDate, targetDate, True)
        If stateDate = targetDate Then
            values(0) = ExistingFlowValue(CStr(book), targetDate, 3) + CDbl(previousState(1)) - currentSnapshot
            values(1) = ExistingFlowValue(CStr(book), targetDate, 4) + feeAmount - CDbl(previousState(2)) + optim - CDbl(previousState(3))
        Else
            values(0) = CDbl(previousState(1)) - currentSnapshot
            values(1) = feeAmount - CDbl(previousState(2)) + optim
        End If
        values(0) = Round(values(0) - manualCost - baselineCost, 2)
        values(1) = Round(values(1) - manualFees - baselineFees, 2)
        values(2) = Round(replication, 2)
        values = Array(values(0), values(1), values(2), currentSnapshot, feeAmount, optim)
        result(CStr(book)) = values
    Next book
    Set CalculateFlows = result
End Function

Private Function IsFirstAutomaticDateAfterBaseline(ByVal targetDate As Date) As Boolean
    Dim ws As Worksheet, lastRow As Long, r As Long, baseline As Date, rowDate As Date
    baseline = BaselineDate()
    If baseline = 0 Or targetDate <= baseline Then Exit Function
    Set ws = ThisWorkbook.Worksheets(SH_COSTS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 5 To lastRow
        If StrComp(Trim$(CStr(ws.Cells(r, 18).Value)), "AUTO", vbTextCompare) = 0 And IsDate(ws.Cells(r, 1).Value) Then
            rowDate = DateValue(ws.Cells(r, 1).Value)
            If rowDate > baseline And rowDate < targetDate Then Exit Function
        End If
    Next r
    IsFirstAutomaticDateAfterBaseline = True
End Function

Private Function CostRowAdjustment(ByVal book As String, ByVal sourceType As String, ByVal rowDate As Date, ByVal feeColumn As Boolean) As Double
    Dim ws As Worksheet, lastRow As Long, r As Long, columnIndex As Long
    columnIndex = CostColumn(book, feeColumn)
    If columnIndex = 0 Then Exit Function
    Set ws = ThisWorkbook.Worksheets(SH_COSTS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 5 To lastRow
        If IsDate(ws.Cells(r, 1).Value) Then
            If DateValue(ws.Cells(r, 1).Value) = rowDate And StrComp(Trim$(CStr(ws.Cells(r, 18).Value)), sourceType, vbTextCompare) = 0 Then
                CostRowAdjustment = NzNumber(ws.Cells(r, columnIndex).Value)
                Exit Function
            End If
        End If
    Next r
End Function

Private Function LogisticsSnapshots(ByVal ws As Worksheet) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary"): result.CompareMode = vbTextCompare
    result("CGA_SHT1") = -SumNumeric(ws.Range("B2:B4"))
    result("CGA_AVB") = -SumNumeric(ws.Range("E2:E4"))
    result("CGA_TVB") = -SumNumeric(ws.Range("D2:D4"))
    result("CGA_FS") = -SumNumeric(ws.Range("F2:F4"))
    result("PIRINEOS") = -SumNumeric(ws.Range("G2:G4"))
    result("CGC_BUNKER") = -SumNumeric(ws.Range("I2:I4"))
    Set LogisticsSnapshots = result
End Function

Private Function FeeSnapshot(ByVal ws As Worksheet, ByVal book As String) As Double
    Dim op1 As String, op2 As String
    If StrComp(book, "CGA_TVB", vbTextCompare) = 0 Then
        op1 = "PVB-TVB": op2 = "TVB-TVB"
    ElseIf StrComp(book, "CGA_AVB", vbTextCompare) = 0 Then
        op1 = "PVB-AVB"
    Else
        Exit Function
    End If
    FeeSnapshot = SumCanonesAfterCutoff(ws, DateSerial(2025, 12, 31), op1, op2)
End Function

Private Function OptimizationFlow(ByVal ws As Worksheet, ByVal book As String, ByVal previousDate As Date, ByVal targetDate As Date) As Double
    Dim strategy As String, lastRow As Long, r As Long, tradeDate As Date, data As Variant
    If StrComp(book, "CGA_TVB", vbTextCompare) = 0 Then strategy = "REGAS"
    If StrComp(book, "CGA_AVB", vbTextCompare) = 0 Then strategy = "AASS"
    If Len(strategy) = 0 Then Exit Function
    lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
    If lastRow < 2 Then Exit Function
    data = ws.Range("A2:L" & lastRow).Value2
    For r = 2 To lastRow
        If (IsDate(data(r - 1, 3)) Or IsNumeric(data(r - 1, 3))) And StrComp(Trim$(CStr(data(r - 1, 12))), strategy, vbTextCompare) = 0 Then
            tradeDate = DateValue(CDate(data(r - 1, 3)))
            If tradeDate > previousDate And tradeDate <= targetDate Then OptimizationFlow = OptimizationFlow + NzNumber(data(r - 1, 10))
        End If
    Next r
End Function

Private Function ReplicationFlow(ByVal wb As Workbook, ByVal book As String, ByVal targetDate As Date) As Double
    Dim ws As Worksheet, lastRow As Long, r As Long, economicDate As Date, data As Variant
    If StrComp(book, "CGTO", vbTextCompare) = 0 Or StrComp(book, "CGTINDEX", vbTextCompare) = 0 Then
        Set ws = wb.Worksheets("Index replication")
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If lastRow >= 2 Then data = ws.Range("A2:J" & lastRow).Value2
        For r = 2 To lastRow
            If IsDate(data(r - 1, 1)) Or IsNumeric(data(r - 1, 1)) Then
                economicDate = DateValue(CDate(data(r - 1, 1)))
                If AppliedMarketDate(economicDate) = targetDate Then
                    If StrComp(book, "CGTO", vbTextCompare) = 0 Then ReplicationFlow = ReplicationFlow + NzNumber(data(r - 1, 9))
                    If StrComp(book, "CGTINDEX", vbTextCompare) = 0 Then ReplicationFlow = ReplicationFlow + NzNumber(data(r - 1, 10))
                End If
            End If
        Next r
    End If
    If StrComp(book, "CGTINDEX", vbTextCompare) = 0 Then
        Set ws = wb.Worksheets("MAIN")
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If lastRow >= 2 Then data = ws.Range("A2:C" & lastRow).Value2
        For r = 2 To lastRow
            If IsDate(data(r - 1, 1)) Or IsNumeric(data(r - 1, 1)) Then
                economicDate = DateValue(CDate(data(r - 1, 1)))
                If AppliedMarketDate(economicDate) = targetDate Then ReplicationFlow = ReplicationFlow + NzNumber(data(r - 1, 3))
            End If
        Next r
    End If
End Function

Private Function AppliedMarketDate(ByVal economicDate As Date) As Date
    Dim ws As Worksheet, lastRow As Long, r As Long, candidate As Date, cacheKey As String
    If mAppliedDateCache Is Nothing Then
        Set mAppliedDateCache = CreateObject("Scripting.Dictionary")
    End If
    cacheKey = CStr(CLng(DateValue(economicDate)))
    If mAppliedDateCache.Exists(cacheKey) Then
        AppliedMarketDate = mAppliedDateCache(cacheKey)
        Exit Function
    End If
    Set ws = ThisWorkbook.Worksheets(SH_CALENDAR)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 5 To lastRow
        If IsDate(ws.Cells(r, 1).Value) And CBool(ws.Cells(r, 2).Value) Then
            candidate = DateValue(ws.Cells(r, 1).Value)
            If candidate >= economicDate Then
                AppliedMarketDate = candidate
                mAppliedDateCache(cacheKey) = AppliedMarketDate
                Exit Function
            End If
        End If
    Next r
    mAppliedDateCache(cacheKey) = 0
End Function

Private Sub ValidateCalculatedFlows(ByVal values As Object, ByVal books As Object)
    Dim book As Variant, parts As Variant, i As Long
    If values.Count <> books.Count Then Err.Raise vbObjectError + 2120, , "El cálculo no cubre todos los BOOK activos."
    For Each book In books.Keys
        If Not values.Exists(CStr(book)) Then Err.Raise vbObjectError + 2121, , "Falta el BOOK " & book & "."
        parts = values(CStr(book))
        For i = 0 To 2
            If Not IsNumeric(parts(i)) Then Err.Raise vbObjectError + 2122, , "Importe no numérico para " & book & "."
        Next i
    Next book
End Sub

Private Sub PublishFlows(ByVal values As Object, ByVal targetDate As Date, ByVal books As Object, ByRef rowCount As Long, ByRef nonZeroCount As Long)
    Dim ws As Worksheet, lo As ListObject, r As Long, book As Variant, parts As Variant, newRow As ListRow
    Set ws = ThisWorkbook.Worksheets(SH_FLOWS)
    Set lo = ws.ListObjects("tblOperatingFlows")
    If Not lo.DataBodyRange Is Nothing Then
        For r = lo.ListRows.Count To 1 Step -1
            If IsDate(lo.DataBodyRange.Cells(r, 1).Value) Then
                If DateValue(lo.DataBodyRange.Cells(r, 1).Value) = targetDate Then lo.ListRows(r).Delete
            End If
        Next r
    End If
    For Each book In books.Keys
        parts = values(CStr(book))
        Set newRow = lo.ListRows.Add
        newRow.Range.Cells(1, 1).Value = targetDate
        newRow.Range.Cells(1, 2).Value = CStr(book)
        newRow.Range.Cells(1, 3).Value = CDbl(parts(0))
        newRow.Range.Cells(1, 4).Value = CDbl(parts(1))
        newRow.Range.Cells(1, 5).Value = CDbl(parts(2))
        newRow.Range.Cells(1, 6).Value = "FOTO-FO-" & Format$(targetDate, "yyyymmdd") & "-" & CStr(book)
        rowCount = rowCount + 1
        If Abs(CDbl(parts(0))) + Abs(CDbl(parts(1))) + Abs(CDbl(parts(2))) > 0.0000001 Then nonZeroCount = nonZeroCount + 1
    Next book
    lo.ListColumns(1).DataBodyRange.NumberFormat = "yyyy-mm-dd"
End Sub

Private Function CostColumn(ByVal book As String, ByVal feeColumn As Boolean) As Long
    If feeColumn Then
        If StrComp(book, "CGA_TVB", vbTextCompare) = 0 Then CostColumn = 15
        If StrComp(book, "CGA_AVB", vbTextCompare) = 0 Then CostColumn = 16
        Exit Function
    End If
    Select Case UCase$(book)
        Case "CGA_SHT1": CostColumn = 2
        Case "CGA_AVB": CostColumn = 3
        Case "CGTINDEX": CostColumn = 4
        Case "CGA_TVB": CostColumn = 5
        Case "CGA_FS": CostColumn = 6
        Case "CGTO": CostColumn = 7
        Case "PROP.TRADING": CostColumn = 8
        Case "PIRINEOS": CostColumn = 9
        Case "CGC_BUNKER": CostColumn = 10
        Case "PVB FLOW": CostColumn = 11
        Case "BIOMETHANE": CostColumn = 12
        Case "COBERTURAS CLIENTES": CostColumn = 13
        Case "CGA_GS": CostColumn = 14
    End Select
End Function

Private Function ManualCostAdjustment(ByVal book As String, ByVal stateDate As Date, ByVal targetDate As Date, ByVal feeColumn As Boolean) As Double
    Dim ws As Worksheet, lastRow As Long, r As Long, columnIndex As Long, rowDate As Date
    columnIndex = CostColumn(book, feeColumn)
    If columnIndex = 0 Then Exit Function
    Set ws = ThisWorkbook.Worksheets(SH_COSTS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 5 To lastRow
        If IsDate(ws.Cells(r, 1).Value) And StrComp(Trim$(CStr(ws.Cells(r, 18).Value)), "MANUAL", vbTextCompare) = 0 Then
            rowDate = DateValue(ws.Cells(r, 1).Value)
            If rowDate > stateDate And rowDate < targetDate Then
                ManualCostAdjustment = ManualCostAdjustment + NzNumber(ws.Cells(r, columnIndex).Value)
            End If
        End If
    Next r
End Function

Private Sub PublishCosts(ByVal values As Object, ByVal targetDate As Date, ByVal books As Object)
    Dim ws As Worksheet, r As Long, targetRow As Long, lastRow As Long
    Dim book As Variant, parts As Variant, columnIndex As Long
    Set ws = ThisWorkbook.Worksheets(SH_COSTS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 5 Then lastRow = 4
    For r = 5 To lastRow
        If IsDate(ws.Cells(r, 1).Value) Then
            If DateValue(ws.Cells(r, 1).Value) = targetDate Then targetRow = r: Exit For
        End If
    Next r
    If targetRow = 0 Then
        targetRow = lastRow + 1
        If targetRow < 5 Then targetRow = 5
    End If
    If StrComp(Trim$(CStr(ws.Cells(targetRow, 18).Value)), "MANUAL", vbTextCompare) = 0 Or StrComp(Trim$(CStr(ws.Cells(targetRow, 18).Value)), "BASELINE", vbTextCompare) = 0 Then
        Err.Raise vbObjectError + 2150, , "COSTS ya contiene una fila MANUAL o BASELINE para el Market Date objetivo. No se sobrescribe."
    End If
    ws.Range(ws.Cells(targetRow, 2), ws.Cells(targetRow, 16)).ClearContents
    ws.Cells(targetRow, 1).Value = targetDate
    For Each book In books.Keys
        parts = values(CStr(book))
        columnIndex = CostColumn(CStr(book), False)
        If columnIndex > 0 Then ws.Cells(targetRow, columnIndex).Value = CDbl(parts(0))
    Next book
    ws.Cells(targetRow, 17).FormulaR1C1 = "=SUM(RC[-15]:RC[-1])"
    ws.Cells(targetRow, 18).Value = "AUTO"
    ws.Cells(targetRow, 19).Value = Now
    ws.Cells(targetRow, 20).Value = Environ$("Username")
    ws.Cells(targetRow, 21).Value = "Foto FO: " & Format$(targetDate, "yyyy-mm-dd")
    ws.Cells(targetRow, 1).NumberFormat = "yyyy-mm-dd"
    ws.Cells(targetRow, 19).NumberFormat = "yyyy-mm-dd hh:mm:ss"
End Sub

Private Function ExistingFlowValue(ByVal book As String, ByVal targetDate As Date, ByVal columnIndex As Long) As Double
    Dim lo As ListObject, r As Long
    Set lo = ThisWorkbook.Worksheets(SH_FLOWS).ListObjects("tblOperatingFlows")
    If lo.DataBodyRange Is Nothing Then Exit Function
    For r = 1 To lo.ListRows.Count
        If IsDate(lo.DataBodyRange.Cells(r, 1).Value) Then
            If DateValue(lo.DataBodyRange.Cells(r, 1).Value) = targetDate And StrComp(Trim$(CStr(lo.DataBodyRange.Cells(r, 2).Value)), book, vbTextCompare) = 0 Then
                ExistingFlowValue = NzNumber(lo.DataBodyRange.Cells(r, columnIndex).Value)
                Exit Function
            End If
        End If
    Next r
End Function

Private Function ReadState(ByVal book As String) As Variant
    Dim ws As Worksheet, lastRow As Long, r As Long
    Set ws = ThisWorkbook.Worksheets(SH_STATE)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), book, vbTextCompare) = 0 Then
            If Not IsDate(ws.Cells(r, 2).Value) Then Err.Raise vbObjectError + 2140, , "Estado sin fecha para " & book & "."
            ReadState = Array(DateValue(ws.Cells(r, 2).Value), NzNumber(ws.Cells(r, 3).Value), NzNumber(ws.Cells(r, 4).Value), NzNumber(ws.Cells(r, 5).Value))
            Exit Function
        End If
    Next r
    Err.Raise vbObjectError + 2141, , "Falta estado inicial de Foto FO para " & book & "."
End Function

Private Sub SaveState(ByVal values As Object, ByVal targetDate As Date, ByVal books As Object)
    Dim ws As Worksheet, lastRow As Long, r As Long, book As Variant, parts As Variant, found As Boolean
    Set ws = ThisWorkbook.Worksheets(SH_STATE)
    For Each book In books.Keys
        found = False
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        For r = 2 To lastRow
            If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), CStr(book), vbTextCompare) = 0 Then found = True: Exit For
        Next r
        If Not found Then r = lastRow + 1
        parts = values(CStr(book))
        ws.Cells(r, 1).Value = CStr(book)
        ws.Cells(r, 2).Value = targetDate
        ws.Cells(r, 3).Value = CDbl(parts(3))
        ws.Cells(r, 4).Value = CDbl(parts(4))
        ws.Cells(r, 5).Value = CDbl(parts(5))
        ws.Cells(r, 6).Value = Now
    Next book
End Sub

Private Function SumCanonesAfterCutoff(ByVal ws As Worksheet, ByVal cutoffDate As Date, ByVal operation1 As String, Optional ByVal operation2 As String = "") As Double
    Dim lastRow As Long, r As Long, startDate As Date, operation As String, data As Variant
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Function
    data = ws.Range("A2:J" & lastRow).Value2
    For r = 2 To lastRow
        If IsDate(data(r - 1, 1)) Or IsNumeric(data(r - 1, 1)) Then
            startDate = DateValue(CDate(data(r - 1, 1)))
            operation = Trim$(CStr(data(r - 1, 3)))
            If startDate > cutoffDate And (StrComp(operation, operation1, vbTextCompare) = 0 Or (Len(operation2) > 0 And StrComp(operation, operation2, vbTextCompare) = 0)) Then
                SumCanonesAfterCutoff = SumCanonesAfterCutoff + NzNumber(data(r - 1, 10))
            End If
        End If
    Next r
End Function

Private Function SumNumeric(ByVal rng As Range) As Double
    Dim cell As Range
    For Each cell In rng.Cells
        SumNumeric = SumNumeric + NzNumber(cell.Value)
    Next cell
End Function

Private Function NzNumber(ByVal value As Variant) As Double
    If Not IsError(value) And IsNumeric(value) Then NzNumber = CDbl(value)
End Function

Private Function SheetExists(ByVal name As String, ByVal wb As Workbook) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(name)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Private Sub RequireHeader(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal columnIndex As Long, ByVal expected As String)
    If StrComp(Trim$(CStr(ws.Cells(rowIndex, columnIndex).Value)), expected, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 2130, , "Cabecera inesperada en " & ws.Name & "!" & ws.Cells(rowIndex, columnIndex).Address(False, False) & "."
End Sub

Private Function YearExists(ByVal ws As Worksheet, ByVal targetYear As Long) As Boolean
    Dim lastRow As Long, r As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If IsNumeric(ws.Cells(r, 1).Value) Then If CLng(ws.Cells(r, 1).Value) = targetYear Then YearExists = True: Exit Function
    Next r
End Function

Private Function MaxDateInColumn(ByVal ws As Worksheet, ByVal columnIndex As Long, ByVal firstRow As Long) As Date
    Dim lastRow As Long, r As Long, d As Date
    lastRow = ws.Cells(ws.Rows.Count, columnIndex).End(xlUp).Row
    For r = firstRow To lastRow
        If IsDate(ws.Cells(r, columnIndex).Value) Then d = DateValue(ws.Cells(r, columnIndex).Value): If d > MaxDateInColumn Then MaxDateInColumn = d
    Next r
End Function

Private Sub UpdateManualStatus(ByVal stamp As Date, ByVal targetDate As Date, ByVal source As String, ByVal status As String, ByVal detail As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_MANUAL)
    ws.Range("I6").Value = stamp
    ws.Range("L6").Value = targetDate
    ws.Range("M6").Value = status & " - " & detail
    ws.Range("J6").Value = Environ$("Username")
    ws.Range("F6").Value = source
    ws.Range("I6").NumberFormat = "yyyy-mm-dd hh:mm:ss"
    ws.Range("L6").NumberFormat = "yyyy-mm-dd"
End Sub

Private Sub WriteAudit(ByVal startedAt As Date, ByVal finishedAt As Date, ByVal targetDate As Date, ByVal source As String, ByVal status As String, ByVal detail As String, ByVal rowsWritten As Long, ByVal expectedBooks As Long)
    Dim ws As Worksheet, nextRow As Long
    Set ws = ThisWorkbook.Worksheets(SH_LOG)
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < 5 Then nextRow = 5
    ws.Cells(nextRow, 1).Value = Format$(startedAt, "yyyymmdd-hhnnss")
    ws.Cells(nextRow, 2).Value = startedAt
    ws.Cells(nextRow, 3).Value = finishedAt
    ws.Cells(nextRow, 4).Value = targetDate
    ws.Cells(nextRow, 5).Value = Environ$("Username")
    ws.Cells(nextRow, 6).Value = source
    ws.Cells(nextRow, 7).Value = status
    ws.Cells(nextRow, 8).Value = rowsWritten
    ws.Cells(nextRow, 9).Value = expectedBooks
    ws.Cells(nextRow, 10).Value = detail
End Sub

Private Sub RestoreExcel(ByVal calcMode As XlCalculation, ByVal eventsValue As Boolean, ByVal screenValue As Boolean, ByVal alertsValue As Boolean)
    Application.Calculation = calcMode
    Application.EnableEvents = eventsValue
    Application.ScreenUpdating = screenValue
    Application.DisplayAlerts = alertsValue
    Application.StatusBar = False
End Sub
