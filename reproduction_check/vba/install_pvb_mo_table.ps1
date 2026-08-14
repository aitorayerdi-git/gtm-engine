param(
    [string]$TargetPath = (Join-Path $PSScriptRoot '..\Input data.xlsm')
)

$ErrorActionPreference = 'Stop'
function Invoke-ExcelAction {
    param([scriptblock]$Action)
    for ($attempt = 1; $attempt -le 20; $attempt++) {
        try { return & $Action }
        catch {
            if ($attempt -eq 20) { throw }
            Start-Sleep -Milliseconds 500
        }
    }
}

$target = (Resolve-Path -LiteralPath $TargetPath).Path
$repository = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..')).Path
if (-not $target.StartsWith($repository + '\', [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "Target workbook is outside the repository: $target"
}

$backup = Join-Path ([System.IO.Path]::GetDirectoryName($target)) 'Input data pre PVB MO table 20260814.xlsm'
if (-not (Test-Path -LiteralPath $backup)) {
    Copy-Item -LiteralPath $target -Destination $backup
}

$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 3
    $excel.ScreenUpdating = $false
    $temporaryWorkbook = $excel.Workbooks.Add()
    $excel.Calculation = -4135
    $excel.CalculateBeforeSave = $false
    $temporaryWorkbook.Close($false)
    $workbook = $excel.Workbooks.Open($target, 0, $false)
    if ($workbook.ReadOnly) { throw "Workbook opened read-only: $target" }
    Start-Sleep -Seconds 10

    $manualSheet = Invoke-ExcelAction { $workbook.Worksheets.Item('MANUAL CHANGES') }
    $manualTable = Invoke-ExcelAction { $manualSheet.ListObjects.Item('tblManualDates') }
    $lastManualRow = $manualTable.ListRows.Item($manualTable.ListRows.Count)
    if ([string]$lastManualRow.Range.Cells.Item(1, 1).Value2 -eq 'Market Date INITIAL POSITION PVB MO') {
        $firstRow = $manualTable.Range.Row
        $firstColumn = $manualTable.Range.Column
        $lastRow = $firstRow + $manualTable.Range.Rows.Count - 2
        $lastColumn = $firstColumn + $manualTable.Range.Columns.Count - 1
        $reducedRange = $manualSheet.Range(
            $manualSheet.Cells.Item($firstRow, $firstColumn),
            $manualSheet.Cells.Item($lastRow, $lastColumn)
        )
        Invoke-ExcelAction { $manualTable.Resize($reducedRange) }
        Start-Sleep -Milliseconds 1000
        Invoke-ExcelAction { $manualSheet.Range('O9:Q9').Value2 = $null }
    }

    try {
        $sheet = $workbook.Worksheets.Item('INITIAL POSITION PVB MO')
        while ($sheet.ListObjects.Count -gt 0) { $sheet.ListObjects.Item(1).Delete() }
        $sheet.Cells.Clear()
    }
    catch {
        $afterSheet = $workbook.Worksheets.Item('INITIAL PNL')
        $sheet = $workbook.Worksheets.Add([System.Type]::Missing, $afterSheet)
        $sheet.Name = 'INITIAL POSITION PVB MO'
    }

    $sheet.Cells.Item(1, 1).Value2 = 'INITIAL POSITION PVB MO'
    $sheet.Cells.Item(1, 1).Font.Bold = $true
    $sheet.Cells.Item(1, 1).Font.Size = 16
    $sheet.Cells.Item(2, 1).Value2 = 'Una fila por Delivery Month. Market Date es la fecha de cierre de la informacion de MO. Index PVB Aggregated es control; Exposure utiliza el desglose para evitar doble contabilizacion.'

    $headers = @(
        'Delivery Month',
        'Market Date',
        'PVB',
        'Index PVB Aggregated',
        'GWDES Auction',
        'D+1 Auction',
        'Mibgas Index ES',
        'MIBGAS D+1 Daily Reference',
        'MIBGAS LPI',
        'PVB Heren',
        'Mibgas API DA',
        'Mibgas MA'
    )
    for ($column = 1; $column -le $headers.Count; $column++) {
        $cell = $sheet.Cells.Item(4, $column)
        $cell.Value2 = $headers[$column - 1]
        $cell.Font.Bold = $true
        $cell.Font.Color = 16777215
        $cell.Interior.Color = 11625472
        $cell.WrapText = $true
    }

    $historicalEnd = $null
    foreach ($row in $manualTable.ListRows) {
        if ([string]$row.Range.Cells.Item(1, 1).Value2 -eq 'Historical End Date') {
            $historicalEnd = [DateTime]::FromOADate([double]$row.Range.Cells.Item(1, 2).Value2)
            break
        }
    }
    if ($null -eq $historicalEnd) { throw 'Historical End Date not found in tblManualDates.' }
    $firstDeliveryMonth = ([DateTime]::new($historicalEnd.Year, $historicalEnd.Month, 1)).AddMonths(1)

    $tableRange = $sheet.Range($sheet.Cells.Item(4, 1), $sheet.Cells.Item(6, $headers.Count))
    $table = $sheet.ListObjects.Add(1, $tableRange, $null, 1)
    $table.Name = 'tblPvbMoInitialPosition'
    $table.TableStyle = 'TableStyleMedium2'
    for ($column = 1; $column -le $headers.Count; $column++) {
        $table.HeaderRowRange.Cells.Item(1, $column).Value2 = $headers[$column - 1]
    }
    $table.DataBodyRange.Cells.Item(1, 1).Value = $firstDeliveryMonth
    $table.DataBodyRange.Cells.Item(2, 1).Value = $firstDeliveryMonth.AddMonths(1)

    $sheet.Range('A5:B1048576').NumberFormat = 'yyyy-mm-dd'
    $sheet.Range($sheet.Cells.Item(5, 1), $sheet.Cells.Item(6, $headers.Count)).Font.Color = 15636889
    $sheet.Columns.Item(1).ColumnWidth = 16
    $sheet.Columns.Item(2).ColumnWidth = 16
    for ($column = 3; $column -le $headers.Count; $column++) {
        $sheet.Columns.Item($column).ColumnWidth = 20
        $sheet.Columns.Item($column).NumberFormat = '#,##0.000'
    }
    $sheet.Columns.Item(8).ColumnWidth = 28
    $sheet.Range('A4').Select()
    $excel.ActiveWindow.FreezePanes = $false
    $sheet.Range('A5').Select()
    $excel.ActiveWindow.FreezePanes = $true
    $sheet.Tab.Color = 11625472

    [void]($workbook.ForceFullCalculation = $false)
    Invoke-ExcelAction { $workbook.Save() }
    Write-Output "Installed tblPvbMoInitialPosition in $target"
    Write-Output "Backup: $backup"
}
finally {
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
