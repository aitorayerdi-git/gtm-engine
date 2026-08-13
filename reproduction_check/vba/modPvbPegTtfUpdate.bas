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
    Dim lastRowPVB As Long
    Dim lastRowPEG As Long
    Dim formulasPVB As Variant
    Dim formulasPEG As Variant
    Dim lastDatePVB As Variant
    Dim lastDatePEG As Variant
    Dim errorMessage As String

    On Error GoTo ErrHandler

    Set wsPVB = ThisWorkbook.Worksheets("PVB-TTF")
    Set wsPEG = ThisWorkbook.Worksheets("PEG-TTF")
    oldCalc = Application.Calculation
    oldEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    lastRowPVB = wsPVB.Cells(wsPVB.Rows.Count, "B").End(xlUp).Row
    lastRowPEG = wsPEG.Cells(wsPEG.Rows.Count, "B").End(xlUp).Row
    If lastRowPVB >= 1 Then formulasPVB = wsPVB.Range("B1:B" & lastRowPVB).FormulaR1C1
    If lastRowPEG >= 1 Then formulasPEG = wsPEG.Range("B1:B" & lastRowPEG).FormulaR1C1

    Set wbLocation = Workbooks.Open( _
        Filename:=LOCATIONSP_URL, ReadOnly:=True, UpdateLinks:=True)
    wsPVB.Calculate
    wsPEG.Calculate

    If lastRowPVB >= 1 Then wsPVB.Range("B1:B" & lastRowPVB).FormulaR1C1 = formulasPVB
    If lastRowPEG >= 1 Then wsPEG.Range("B1:B" & lastRowPEG).FormulaR1C1 = formulasPEG
    wbLocation.Close SaveChanges:=False
    Set wbLocation = Nothing

    lastDatePVB = LastDateWithValueInAL(wsPVB)
    lastDatePEG = LastDateWithValueInAL(wsPEG)

CleanExit:
    On Error Resume Next
    If Not wbLocation Is Nothing Then wbLocation.Close SaveChanges:=False
    If lastRowPVB >= 1 Then wsPVB.Range("B1:B" & lastRowPVB).FormulaR1C1 = formulasPVB
    If lastRowPEG >= 1 Then wsPEG.Range("B1:B" & lastRowPEG).FormulaR1C1 = formulasPEG
    Application.Calculation = oldCalc
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreenUpdating
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
    errorMessage = "Error " & Err.Number & ": " & Err.Description
    Resume CleanExit
End Sub

Private Function LastDateWithValueInAL(ByVal ws As Worksheet) As Variant
    Dim rowNumber As Long
    For rowNumber = ws.Cells(ws.Rows.Count, "AL").End(xlUp).Row To 1 Step -1
        If Len(Trim$(CStr(ws.Cells(rowNumber, "AL").Value))) > 0 Then
            LastDateWithValueInAL = ws.Cells(rowNumber, "B").Value
            Exit Function
        End If
    Next rowNumber
    LastDateWithValueInAL = Empty
End Function
