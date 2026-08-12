param([Parameter(Mandatory = $true)][string]$BasePath)
$ErrorActionPreference = 'Stop'
$testDirectory = Split-Path -Parent $BasePath

function Invoke-BaselineCase([string]$Path, [double]$BaselineAmount) {
    Copy-Item -LiteralPath $BasePath -Destination $Path -Force
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.EnableEvents = $false
    $control = $excel.Workbooks.Add(); $excel.Calculation = -4135
    $workbook = $null
    try {
        try { $excel.COMAddIns.Item('AStorm.Ui.StormUi').Connect = $false } catch {}
        $workbook = $excel.Workbooks.Open($Path, 0, $false)
        $control.Close($false); $control = $null
        $targetRaw = $workbook.Worksheets.Item('MANUAL CHANGES').Range('F5').Value2
        $targetDate = if ($targetRaw -is [datetime]) { $targetRaw } else { [datetime]::FromOADate([double]$targetRaw) }
        $calendar = $workbook.Worksheets.Item('MARKET CALENDAR')
        $lastRow = $calendar.Cells($calendar.Rows.Count, 1).End(-4162).Row
        $priorDates = @()
        for ($r = 5; $r -le $lastRow; $r++) {
            $raw = $calendar.Cells.Item($r, 1).Value2
            if ($null -ne $raw -and [bool]$calendar.Cells.Item($r, 2).Value2) {
                $d = if ($raw -is [datetime]) { $raw } else { [datetime]::FromOADate([double]$raw) }
                if ($d -lt $targetDate) { $priorDates += $d }
            }
        }
        $priorDates = @($priorDates | Sort-Object)
        $baselineDate = $priorDates[-1]; $oldStateDate = $priorDates[-2]
        $state = $workbook.Worksheets.Item('FOTO FO STATE')
        $stateLastRow = $state.Cells($state.Rows.Count, 1).End(-4162).Row
        $state.Range("B2:B$stateLastRow").Value2 = $oldStateDate.ToOADate()
        $costs = $workbook.Worksheets.Item('COSTS')
        $costs.Range('A5:U504').ClearContents()
        $costs.Range('A5').Value2 = $baselineDate.ToOADate()
        $costs.Range('B5').Value2 = $BaselineAmount
        $costs.Range('Q5').Formula = '=SUM(B5:P5)'
        $costs.Range('R5').Value2 = 'BASELINE'
        $costs.Range('S5').Value2 = (Get-Date).ToOADate()
        $costs.Range('T5').Value2 = [string]$env:USERNAME
        $costs.Range('U5').Value2 = 'Opening cumulative balance regression'
        $table = $workbook.Worksheets.Item('OPERATING FLOWS').ListObjects.Item('tblOperatingFlows')
        $baselineRowsBefore = 0
        for ($r = 1; $r -le $table.ListRows.Count; $r++) {
            $raw = $table.DataBodyRange.Cells($r, 1).Value2
            if ($null -ne $raw) {
                $d = if ($raw -is [datetime]) { $raw } else { [datetime]::FromOADate([double]$raw) }
                if ($d.Date -eq $baselineDate.Date) { $baselineRowsBefore++ }
            }
        }
        $excel.Run("'$($workbook.Name)'!UpdateFotoFO_LocalTest")
        $status = [string]$workbook.Worksheets.Item('MANUAL CHANGES').Range('M6').Text
        if (-not $status.StartsWith('OK')) { throw "Baseline macro failed: $status" }
        $flow = $null; $baselineRowsAfter = 0
        for ($r = 1; $r -le $table.ListRows.Count; $r++) {
            $raw = $table.DataBodyRange.Cells($r, 1).Value2
            if ($null -eq $raw) { continue }
            $d = if ($raw -is [datetime]) { $raw } else { [datetime]::FromOADate([double]$raw) }
            if ($d.Date -eq $baselineDate.Date) { $baselineRowsAfter++ }
            if ($d.Date -eq $targetDate.Date -and [string]$table.DataBodyRange.Cells($r, 2).Value2 -eq 'CGA_SHT1') {
                $flow = [double]$table.DataBodyRange.Cells($r, 3).Value2
            }
        }
        if ($baselineRowsAfter -ne $baselineRowsBefore) { throw 'BASELINE changed OPERATING FLOWS rows for its own date.' }
        if ($null -eq $flow) { throw 'Target CGA_SHT1 flow not found.' }
        return $flow
    }
    finally {
        if ($null -ne $workbook) { $workbook.Close($false) }
        if ($null -ne $control) { $control.Close($false) }
        $excel.Quit()
    }
}

$zeroOutput = @(Invoke-BaselineCase (Join-Path $testDirectory 'baseline_zero.xlsm') 0)
$adjustedOutput = @(Invoke-BaselineCase (Join-Path $testDirectory 'baseline_adjusted.xlsm') 123.45)
$zero = [double]$zeroOutput[-1]
$adjusted = [double]$adjustedOutput[-1]
$difference = [math]::Round($zero - $adjusted, 2)
if ([math]::Abs($difference - 123.45) -gt 0.01) { throw "Expected baseline difference 123.45, got $difference" }
Write-Output "ZERO_BASELINE_FLOW=$zero"
Write-Output "ADJUSTED_BASELINE_FLOW=$adjusted"
Write-Output "BASELINE_DEDUCTION=$difference"
Write-Output 'BASELINE_NOT_PUBLISHED=OK'
Write-Output 'BASELINE_TEST=OK'
