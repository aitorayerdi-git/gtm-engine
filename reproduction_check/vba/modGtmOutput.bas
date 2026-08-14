Attribute VB_Name = "modGtmOutput"
Option Explicit

Public Sub Generar_GTM_Output()
    GenerateGtmOutputCore True
End Sub

Public Sub Generar_GTM_Output_Silent()
    GenerateGtmOutputCore False
End Sub

Private Sub GenerateGtmOutputCore(ByVal showMessages As Boolean)
    Dim requiredSources As Variant, source As Variant
    Dim repositoryPath As String, executablePath As String, outputPath As String
    Dim logPath As String, inputSnapshot As String, command As String, exitCode As Long
    Dim shell As Object, targetDate As Date, refreshCommand As String

    On Error GoTo ErrHandler
    requiredSources = Array("MarketView", "Foto FO", "PVB-TTF", "PEG-TTF")
    For Each source In requiredSources
        If StrComp(Trim$(SourceControlStatus(CStr(source))), "OK", vbTextCompare) <> 0 Then
            If showMessages Then MsgBox "No se puede generar el output." & vbCrLf & _
                CStr(source) & " no tiene Status OK.", vbExclamation, "GTM Output - CHECK"
            Exit Sub
        End If
    Next source

    targetDate = MarketControlDate("Historical End Date")
    If MarketStatusText(LatestReutersProviderDate()) <> "OK" Then
        If showMessages Then MsgBox "No se puede generar el output." & vbCrLf & _
            "Los datos fuente de Reuters no llegan a " & Format$(targetDate, "dd/mm/yyyy") & ".", _
            vbExclamation, "GTM Output - CHECK"
        Exit Sub
    End If

    If Not ThisWorkbook.Saved Then ThisWorkbook.Save
    repositoryPath = ThisWorkbook.Path & "\.."
    executablePath = repositoryPath & "\.venv\Scripts\gtm-engine.exe"
    outputPath = repositoryPath & "\outputs\gtm_excel_runs"
    logPath = outputPath & "\GTM_button_last_run.log"
    inputSnapshot = outputPath & "\GTM_Button_Input.xlsx"
    If Len(Dir$(executablePath)) = 0 Then Err.Raise vbObjectError + 2501, , _
        "No existe el ejecutable del motor: " & executablePath
    If Len(Dir$(outputPath, vbDirectory)) = 0 Then MkDir outputPath
    ExportValueOnlyInput inputSnapshot

    refreshCommand = "cmd.exe /c """"" & executablePath & _
        """ excel-refresh-curves --workbook """ & inputSnapshot & _
        """ --historical-end " & Format$(targetDate, "yyyy-mm-dd") & _
        " > """ & logPath & """ 2>&1""
    Application.StatusBar = "Publicando CURVE PRICES..."
    Set shell = CreateObject("WScript.Shell")
    exitCode = shell.Run(refreshCommand, 0, True)
    If exitCode <> 0 Then Err.Raise vbObjectError + 2504, , _
        "No se pudo publicar CURVE PRICES. Revisa " & logPath
    SetSourceControl "Reuters", Now, targetDate, "OK"
    ThisWorkbook.Save

    command = "cmd.exe /c """"" & executablePath & """ excel-build --workbook """ & _
        inputSnapshot & """ --output """ & outputPath & _
        """ > """ & logPath & """ 2>&1"""
    Application.StatusBar = "Generando GTM Output..."
    exitCode = shell.Run(command, 0, True)
    Application.StatusBar = False

    If exitCode = 0 And Len(Dir$(outputPath & "\GTM_LATEST.xlsx")) > 0 Then
        targetDate = MarketControlDate("Historical End Date")
        SetSourceControl "GTM Output", Now, targetDate, "OK"
        ThisWorkbook.Save
        If showMessages Then MsgBox "Output generado correctamente:" & vbCrLf & _
            outputPath & "\GTM_LATEST.xlsx", vbInformation, "GTM Output - OK"
    Else
        SetSourceStatusOnly "GTM Output", "ERROR - REVIEW LOG"
        ThisWorkbook.Save
        If showMessages Then MsgBox "El motor no publico un output valido." & vbCrLf & _
            "Revisa: " & logPath, vbCritical, "GTM Output - ERROR"
    End If
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    On Error Resume Next
    SetSourceStatusOnly "GTM Output", "ERROR - " & Err.Description
    ThisWorkbook.Save
    On Error GoTo 0
    If showMessages Then MsgBox "No se pudo generar el output." & vbCrLf & _
        "Error " & Err.Number & ": " & Err.Description, vbCritical, "GTM Output - ERROR"
End Sub

Private Sub ImportCurvePrices(ByVal snapshotPath As String)
    Dim snapshot As Workbook, sourceTable As ListObject, targetTable As ListObject
    Dim rowCount As Long, targetRange As Range, values As Variant, errorMessage As String
    On Error GoTo ImportFailed
    Set snapshot = Workbooks.Open(snapshotPath, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    Set sourceTable = FindSnapshotTable(snapshot, "tblCurvePrices")
    Set targetTable = FindSnapshotTable(ThisWorkbook, "tblCurvePrices")
    rowCount = sourceTable.DataBodyRange.Rows.Count
    values = sourceTable.DataBodyRange.Value
    Set targetRange = targetTable.HeaderRowRange.Resize(rowCount + 1, targetTable.ListColumns.Count)
    targetTable.Resize targetRange
    targetTable.DataBodyRange.Value = values
    snapshot.Close SaveChanges:=False
    Exit Sub
ImportFailed:
    errorMessage = Err.Description
    On Error Resume Next
    If Not snapshot Is Nothing Then snapshot.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise vbObjectError + 2505, , "No se pudo importar CURVE PRICES publicada: " & errorMessage
End Sub

Private Sub ExportValueOnlyInput(ByVal destination As String)
    Dim temporaryMacro As String, snapshot As Workbook, tableName As Variant
    Dim table As ListObject, oldAlerts As Boolean, oldCalc As XlCalculation
    Dim inputTables As Variant, errorMessage As String
    inputTables = Array("tblControl", "tblManualDates", "tblBooks", "tblUnderlyings", _
        "tblMarketCalendar", "tblInitialExposure", "tblInitialPnl", "tblTrades", _
        "tblDeliveryElections", "tblCurvePrices", "tblFixingPrices", "tblFxRates", _
        "tblOperatingFlows")
    temporaryMacro = Left$(destination, Len(destination) - 5) & ".xlsm"
    oldAlerts = Application.DisplayAlerts
    oldCalc = Application.Calculation
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    On Error GoTo ExportFailed
    If Len(Dir$(temporaryMacro)) > 0 Then Kill temporaryMacro
    If Len(Dir$(destination)) > 0 Then Kill destination
    ThisWorkbook.SaveCopyAs temporaryMacro
    Set snapshot = Workbooks.Open(temporaryMacro, UpdateLinks:=0, ReadOnly:=False, AddToMru:=False)
    For Each tableName In inputTables
        Set table = FindSnapshotTable(snapshot, CStr(tableName))
        If Not table.DataBodyRange Is Nothing Then table.DataBodyRange.Value = table.DataBodyRange.Value
    Next tableName
    FreezeSheetValues snapshot, "TTF"
    FreezeSheetValues snapshot, "Brent Dated"
    FreezeSheetValues snapshot, "HH"
    FreezeSheetValues snapshot, "PVB-TTF"
    FreezeSheetValues snapshot, "PEG-TTF"
    snapshot.SaveAs Filename:=destination, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    snapshot.Close SaveChanges:=False
    Set snapshot = Nothing
    Kill temporaryMacro
    Application.Calculation = oldCalc
    Application.DisplayAlerts = oldAlerts
    Exit Sub
ExportFailed:
    errorMessage = Err.Description
    On Error Resume Next
    If Not snapshot Is Nothing Then snapshot.Close SaveChanges:=False
    If Len(Dir$(temporaryMacro)) > 0 Then Kill temporaryMacro
    Application.Calculation = oldCalc
    Application.DisplayAlerts = oldAlerts
    On Error GoTo 0
    Err.Raise vbObjectError + 2502, , "No se pudo crear la fotografia value-only: " & errorMessage
End Sub

Private Sub FreezeSheetValues(ByVal workbook As Workbook, ByVal sheetName As String)
    With workbook.Worksheets(sheetName).UsedRange
        .Value = .Value
    End With
End Sub

Private Function FindSnapshotTable(ByVal workbook As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet, table As ListObject
    For Each ws In workbook.Worksheets
        On Error Resume Next
        Set table = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not table Is Nothing Then Set FindSnapshotTable = table: Exit Function
    Next ws
    Err.Raise vbObjectError + 2503, , "Falta la tabla de entrada " & tableName & "."
End Function
