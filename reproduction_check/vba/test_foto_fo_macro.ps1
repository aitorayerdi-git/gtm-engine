param([Parameter(Mandatory = $true)][string]$Path)
$ErrorActionPreference = 'Stop'
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.EnableEvents = $false
$excel.AskToUpdateLinks = $false
# Test isolation only: prevent RTD/MarketView formulas from recalculating while
# the Foto FO macro is validated. This does not alter the production workbook.
$controlWorkbook = $excel.Workbooks.Add()
$excel.Calculation = -4135 # xlCalculationManual
$marketViewAddIn = $null
try {
    $marketViewAddIn = $excel.COMAddIns.Item('AStorm.Ui.StormUi')
    $marketViewAddIn.Connect = $false
}
catch {
    Write-Output "MARKETVIEW_ADDIN=not loaded"
}
$workbook = $null
try {
    $workbook = $excel.Workbooks.Open($Path, 0, $false)
    $controlWorkbook.Close($false)
    $controlWorkbook = $null
    $excel.Run("'$($workbook.Name)'!UpdateFotoFO_LocalTest")
    $status = $workbook.Worksheets.Item('MANUAL CHANGES').Range('M6').Text
    $stamp = $workbook.Worksheets.Item('MANUAL CHANGES').Range('I6').Text
    $rows = $workbook.Worksheets.Item('OPERATING FLOWS').ListObjects.Item('tblOperatingFlows').ListRows.Count
    $workbook.Save()
    Write-Output "STATUS=$status"
    Write-Output "STAMP=$stamp"
    Write-Output "FLOW_ROWS=$rows"
    if (-not $status.StartsWith('OK')) { throw "Macro validation failed: $status" }
    $workbook.Close($true)
    $workbook = $null
}
finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    if ($null -ne $controlWorkbook) { $controlWorkbook.Close($false) }
    $excel.Quit()
    if ($null -ne $marketViewAddIn) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($marketViewAddIn) }
}
