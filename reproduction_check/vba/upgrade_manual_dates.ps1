param([string]$Path)
$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
if (-not $Path) { $Path = Join-Path $root 'reproduction_check\Input data.xlsm' }
$module = Join-Path $root 'reproduction_check\vba\modFotoFOUpdate.bas'
$excel = $null
$workbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $workbook = $excel.Workbooks.Open($Path, 0, $false)

    $manual = $workbook.Worksheets.Item('MANUAL CHANGES')
    $control = $workbook.Worksheets.Item('CONTROL')
    $savedDates = @{}
    foreach ($existingTable in @($manual.ListObjects)) {
        if ($existingTable.Name -eq 'tblManualDates') {
            foreach ($existingRow in @($existingTable.ListRows)) {
                $savedDates[[string]$existingRow.Range.Cells.Item(1, 1).Value2] = $existingRow.Range.Cells.Item(1, 2).Value2
            }
        }
    }
    $initialDate = $savedDates['Initial Market Date']
    $historicalStart = $savedDates['Historical Start Date']
    $lastMarketDate = $savedDates['Historical End Date']
    $reportStart = $savedDates['Exposure Report Start Month']
    $reportEnd = $savedDates['Exposure Report End Month']
    if (-not $initialDate) { $initialDate = $control.Range('B9').Value2 }
    if (-not $historicalStart) { $historicalStart = $control.Range('B10').Value2 }
    if (-not $lastMarketDate) { $lastMarketDate = $manual.Range('F5').Value2 }
    if (-not $reportStart) { $reportStart = [datetime]'2026-08-01' }
    if (-not $reportEnd) { $reportEnd = [datetime]'2028-12-01' }

    foreach ($table in @($manual.ListObjects)) {
        if ($table.Name -eq 'tblManualDates') { $table.Delete() }
    }
    $manual.Range('A507:C512').ClearContents()
    $manual.Range('O3:Q8').ClearContents()
    $manual.Range('O3').Value = 'Field'
    $manual.Range('P3').Value = 'Date'
    $manual.Range('Q3').Value = 'Explanation'
    $rows = @(
        @('Initial Market Date', $initialDate, 'Closing date represented by INITIAL EXPOSURE and INITIAL PNL.'),
        @('Historical Start Date', $historicalStart, 'First Market Date reconstructed by the engine.'),
        @('Historical End Date', $lastMarketDate, 'Single Last Market Date for Foto FO, Fixings, Costs, Fees, Optimizations, Replication, Exposure and P&L.'),
        @('Exposure Report Start Month', $reportStart, 'First Delivery Month displayed in the Exposure report; enter the first day of the month.'),
        @('Exposure Report End Month', $reportEnd, 'Last Delivery Month displayed in the Exposure report; enter the first day of the month.')
    )
    for ($index = 0; $index -lt $rows.Count; $index++) {
        $row = 4 + $index
        $manual.Cells.Item($row, 15).Value2 = [string]$rows[$index][0]
        $dateValue = $rows[$index][1]
        if ($dateValue -is [datetime]) { $dateValue = $dateValue.ToOADate() }
        $manual.Cells.Item($row, 16).Value2 = [double]$dateValue
        $manual.Cells.Item($row, 17).Value2 = [string]$rows[$index][2]
    }
    $table = $manual.ListObjects.Add(1, $manual.Range('O3:Q8'), $null, 1)
    $table.Name = 'tblManualDates'
    $table.TableStyle = 'TableStyleMedium2'
    $manual.Range('P4:P8').NumberFormat = 'dd/mm/yyyy'
    $manual.Columns.Item('O').ColumnWidth = 30
    $manual.Columns.Item('P').ColumnWidth = 16
    $manual.Columns.Item('Q').ColumnWidth = 75
    $manual.Range('A2:M2').UnMerge()
    $manual.Range('A2:M2').Merge()
    $manual.Range('A2').Value = 'Edit all operative dates in MODEL DATES below. Last Market Date is entered once and is shared by Foto FO and the Python engine.'
    $manual.Range('F4').Value = 'LAST MARKET DATE (DISPLAY)'
    $manual.Range('F5').Formula = '=P6'
    foreach ($button in @($manual.Buttons())) {
        if ($button.Name -eq 'btnUpdateFotoFO') {
            $anchor = $manual.Range('S3:V4')
            $button.Left = $anchor.Left
            $button.Top = $anchor.Top
            $button.Width = $anchor.Width
            $button.Height = $anchor.Height
        }
    }

    $controlTable = $control.ListObjects.Item('tblControl')
    for ($index = $controlTable.ListRows.Count; $index -ge 1; $index--) {
        $field = [string]$controlTable.ListRows.Item($index).Range.Cells.Item(1, 1).Value2
        if ($field -in @('Initial Market Date', 'Historical Start Date', 'Historical End Date')) {
            $controlTable.ListRows.Item($index).Delete()
        }
    }

    foreach ($component in @($workbook.VBProject.VBComponents)) {
        if ($component.Name -eq 'modFotoFOUpdate') {
            $workbook.VBProject.VBComponents.Remove($component)
            break
        }
    }
    $workbook.VBProject.VBComponents.Import($module) | Out-Null
    $workbook.Save()
    Write-Output "Updated $Path"
}
finally {
    if ($workbook) { $workbook.Close($false) }
    if ($excel) { $excel.Quit() }
    foreach ($item in @($workbook, $excel)) {
        if ($item) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($item) }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
