$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$source = Join-Path $root 'reproduction_check\Input data.xlsx'
$destination = Join-Path $root 'reproduction_check\Input data.xlsm'
$module = Join-Path $root 'reproduction_check\vba\modFotoFOUpdate.bas'

$excel = $null
$workbook = $null
$sheet = $null
$button = $null
$templateWorkbook = $null
$costSheet = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AskToUpdateLinks = $false

    try {
        $workbook = $excel.Workbooks.Open($source, 0, $false)
    }
    catch {
        # xlRepairFile=1: Excel repairs only its own OpenXML metadata on a copy
        # before the macro-enabled destination is written.
        $workbook = $excel.Workbooks.Open($source, 0, $false, 5, '', '', $true, 2, '', $false, $false, 0, $true, 1)
    }
    $workbook.SaveAs($destination, 52)

    $costSheet = $workbook.Worksheets.Add([Type]::Missing, $workbook.Worksheets.Item('OPERATING FLOWS'))
    $costSheet.Name = 'COSTS'
    $costSheet.Range('A1:U1').Merge()
    $costSheet.Range('A1').Value = 'COSTS - Foto FO daily deltas. Use MANUAL rows only for missed Market Dates; the next automatic update compensates them.'
    $costSheet.Range('B2:N2').Merge()
    $costSheet.Range('O2:P2').Merge()
    $costSheet.Range('R2:U2').Merge()
    $costSheet.Range('A2').Value = 'Market Date'
    $costSheet.Range('B2').Value = 'LOGISTICS COST (EUR)'
    $costSheet.Range('O2').Value = 'CANONES INTERCAMBIOS (EUR)'
    $costSheet.Range('Q2').Value = "TOTAL COSTS`n+ CANONES"
    $costSheet.Range('R2').Value = 'CONTROL / MANUAL ADJUSTMENT'
    $costHeaders = @('CGA_SHT1','CGA_AVB','CGTINDEX','CGA_TVB','CGA_FS','CGTO','PROP.TRADING','PIRINEOS','CGC_BUNKER','PVB FLOW','BIOMETHANE','COBERTURAS CLIENTES','CGA_GS','CGA_TVB','CGA_AVB')
    for ($headerIndex = 0; $headerIndex -lt $costHeaders.Count; $headerIndex++) {
        $costSheet.Cells.Item(3, $headerIndex + 2).Value = $costHeaders[$headerIndex]
    }
    $costSheet.Range('Q3').Value = 'Total (EUR)'
    $costSheet.Range('R3').Value = 'SOURCE'
    $costSheet.Range('S3').Value = 'UPDATED AT'
    $costSheet.Range('T3').Value = 'UPDATED BY'
    $costSheet.Range('U3').Value = 'COMMENT'
    $costSheet.Range('R4').Value = 'AUTO / MANUAL'
    $costSheet.Range('S4').ClearContents()
    $costSheet.Range('T4').ClearContents()
    $costSheet.Range('U4').Value = 'For a missed date: enter date + delta(s), set SOURCE=MANUAL.'
    $costSheet.Range('B4:P4').FormulaR1C1 = '=SUM(R[1]C:R[500]C)'
    $costSheet.Range('Q4').Formula = '=SUM(B4:P4)'
    $costSheet.Range('Q5:Q504').FormulaR1C1 = '=SUM(RC[-15]:RC[-1])'
    $costSheet.Range('R5:R504').Validation.Delete()
    $costSheet.Range('R5:R504').Validation.Add(3, 1, 1, 'AUTO,MANUAL')
    $costSheet.Range('A5:A504').NumberFormat = 'yyyy-mm-dd'
    $costSheet.Range('S5:S504').NumberFormat = 'yyyy-mm-dd hh:mm:ss'
    $costSheet.Columns.Item('R').ColumnWidth = 12
    $costSheet.Columns.Item('S').ColumnWidth = 20
    $costSheet.Columns.Item('T').ColumnWidth = 16
    $costSheet.Columns.Item('U').ColumnWidth = 45
    $costSheet.Columns.Item('A').ColumnWidth = 14
    $costSheet.Range('B:P').ColumnWidth = 13
    $costSheet.Columns.Item('Q').ColumnWidth = 14
    $costSheet.Range('A1:U1').Interior.Color = 5253199
    $costSheet.Range('A1:U1').Font.Color = 16777215
    $costSheet.Range('A1:U1').Font.Bold = $true
    $costSheet.Range('A1:U1').Font.Size = 14
    $costSheet.Range('A2:U3').Font.Bold = $true
    $costSheet.Range('A2:N2').Interior.Color = 12611584
    $costSheet.Range('O2:Q2').Interior.Color = 49407
    $costSheet.Range('R2:U2').Interior.Color = 5287936
    $costSheet.Range('A2:U3').HorizontalAlignment = -4108
    $costSheet.Range('A2:U3').VerticalAlignment = -4108
    $costSheet.Range('A2:U3').WrapText = $true
    $costSheet.Range('A4:U4').Font.Bold = $true
    $costSheet.Range('A4:U4').Interior.Color = 14277081
    $costSheet.Range('A2:U504').Borders.LineStyle = 1
    $costSheet.Activate()
    $costSheet.Application.ActiveWindow.SplitRow = 4
    $costSheet.Application.ActiveWindow.FreezePanes = $true

    try {
        $component = $workbook.VBProject.VBComponents.Import($module)
    }
    catch {
        throw "Excel blocked programmatic access to the VBA project. Enable 'Trust access to the VBA project object model' and retry. Detail: $($_.Exception.Message)"
    }

    $sheet = $workbook.Worksheets.Item('MANUAL CHANGES')
    foreach ($existing in @($sheet.Buttons())) {
        if ($existing.Name -eq 'btnUpdateFotoFO') { $existing.Delete() }
    }

    $anchor = $sheet.Range('O5:R6')
    $button = $sheet.Buttons().Add($anchor.Left, $anchor.Top, $anchor.Width, $anchor.Height)
    $button.Name = 'btnUpdateFotoFO'
    $button.Characters().Text = 'UPDATE FOTO FO'
    $button.OnAction = 'UpdateFotoFO'
    $button.Font.Bold = $true
    $button.Font.Size = 11

    $sheet.Columns.Item('O:R').ColumnWidth = 12
    $workbook.Save()
    $workbook.Close($true)
    $workbook = $null
    Write-Output "CREATED=$destination"
}
finally {
    if ($null -ne $templateWorkbook) { $templateWorkbook.Close($false) }
    if ($null -ne $workbook) { $workbook.Close($false) }
    if ($null -ne $excel) { $excel.Quit() }
    foreach ($object in @($costSheet, $button, $sheet, $templateWorkbook, $workbook, $excel)) {
        if ($null -ne $object) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($object) }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
