param([Parameter(Mandatory = $true)][string]$BasePath)
$ErrorActionPreference = 'Stop'
$addInKey = 'HKCU:\Software\Microsoft\Office\Excel\Addins\AStorm.Ui.StormUi'
$originalLoadBehavior = (Get-ItemProperty -LiteralPath $addInKey -Name LoadBehavior).LoadBehavior
$testDirectory = Split-Path -Parent $BasePath
$baselinePath = Join-Path $testDirectory 'manual_adjustment_baseline.xlsm'
$manualPath = Join-Path $testDirectory 'manual_adjustment_case.xlsm'

function Invoke-CostCase([string]$Path, [double]$ManualAmount) {
    Copy-Item -LiteralPath $BasePath -Destination $Path -Force
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AskToUpdateLinks = $false
    $control = $excel.Workbooks.Add()
    $excel.Calculation = -4135
    $addIn = $null
    $workbook = $null
    $stage = 'open'
    try {
        try { $addIn = $excel.COMAddIns.Item('AStorm.Ui.StormUi'); $addIn.Connect = $false } catch {}
        $workbook = $excel.Workbooks.Open($Path, 0, $false)
        $control.Close($false); $control = $null
        $stage = 'target-date'
        $targetRaw = $workbook.Worksheets.Item('MANUAL CHANGES').Range('F5').Value2
        $targetDate = if ($targetRaw -is [datetime]) { $targetRaw } else { [datetime]::FromOADate([double]$targetRaw) }
        $stage = 'calendar'
        $calendar = $workbook.Worksheets.Item('MARKET CALENDAR')
        $lastCalendarRow = $calendar.Cells($calendar.Rows.Count, 1).End(-4162).Row
        $priorDates = @()
        for ($r = 5; $r -le $lastCalendarRow; $r++) {
            $dateValue = $calendar.Cells.Item($r, 1).Value2
            $marketValue = $calendar.Cells.Item($r, 2).Value2
            if ($null -ne $dateValue -and ($marketValue -eq $true -or $marketValue -eq 1 -or [string]$marketValue -eq 'TRUE')) {
                $calendarDate = if ($dateValue -is [datetime]) { $dateValue } else { [datetime]::FromOADate([double]$dateValue) }
                if ($calendarDate -lt $targetDate) { $priorDates += $calendarDate }
            }
        }
        $priorDates = @($priorDates | Sort-Object)
        $manualDate = $priorDates[-1]
        $stateDate = $priorDates[-2]
        $stage = 'prepare-state'
        $stateSheet = $workbook.Worksheets.Item('FOTO FO STATE')
        $lastStateRow = $stateSheet.Cells($stateSheet.Rows.Count, 1).End(-4162).Row
        $stateSheet.Range("B2:B$lastStateRow").Value2 = $stateDate.ToOADate()
        if ($ManualAmount -ne 0) {
            $stage = 'prepare-manual-row'
            $costs = $workbook.Worksheets.Item('COSTS')
            $costs.Range('A5:U5').ClearContents()
            $costs.Range('A5').Value2 = $manualDate.ToOADate()
            $costs.Range('B5').Value2 = $ManualAmount
            $costs.Range('Q5').Formula = '=SUM(B5:P5)'
            $costs.Range('R5').Value2 = 'MANUAL'
            $costs.Range('S5').Value2 = (Get-Date).ToOADate()
            $costs.Range('T5').Value2 = [string]$env:USERNAME
            $costs.Range('U5').Value2 = 'Automated compensation test'
        }
        $stage = 'run-macro'
        $excel.Run("'$($workbook.Name)'!UpdateFotoFO_LocalTest")
        $status = $workbook.Worksheets.Item('MANUAL CHANGES').Range('M6').Text
        if (-not $status.StartsWith('OK')) { throw "Macro failed: $status" }
        $table = $workbook.Worksheets.Item('OPERATING FLOWS').ListObjects.Item('tblOperatingFlows')
        $stage = 'read-flow'
        $flow = $null
        for ($r = 1; $r -le $table.ListRows.Count; $r++) {
            $rowDate = $table.DataBodyRange.Cells($r, 1).Value2
            $rowBook = [string]$table.DataBodyRange.Cells($r, 2).Value2
            $flowDate = if ($rowDate -is [datetime]) { $rowDate } elseif ($null -ne $rowDate) { [datetime]::FromOADate([double]$rowDate) } else { $null }
            if ($null -ne $flowDate -and $flowDate.Date -eq $targetDate.Date -and $rowBook -eq 'CGA_SHT1') {
                $flow = [double]$table.DataBodyRange.Cells($r, 3).Value2
                break
            }
        }
        if ($null -eq $flow) { throw 'CGA_SHT1 target flow not found.' }
        $workbook.Save()
        $workbook.Close($true); $workbook = $null
        return $flow
    }
    catch {
        throw "Stage=$stage; line=$($_.InvocationInfo.ScriptLineNumber); $($_.Exception.Message)"
    }
    finally {
        if ($null -ne $workbook) { $workbook.Close($false) }
        if ($null -ne $control) { $control.Close($false) }
        $excel.Quit()
        foreach ($item in @($addIn, $workbook, $control, $excel)) {
            if ($null -ne $item) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($item) }
        }
    }
}

try {
    Set-ItemProperty -LiteralPath $addInKey -Name LoadBehavior -Value 0
    $baselineOutput = @(Invoke-CostCase -Path $baselinePath -ManualAmount 0)
    $manualOutput = @(Invoke-CostCase -Path $manualPath -ManualAmount 123.45)
    $baseline = [double]$baselineOutput[-1]
    $manual = [double]$manualOutput[-1]
    $difference = [math]::Round($baseline - $manual, 2)
    Write-Output "BASELINE_FLOW=$baseline"
    Write-Output "MANUAL_FLOW=$manual"
    Write-Output "COMPENSATED_AMOUNT=$difference"
    if ([math]::Abs($difference - 123.45) -gt 0.01) { throw "Expected compensation 123.45, got $difference" }
    Write-Output 'MANUAL_ADJUSTMENT_TEST=OK'
}
finally {
    Set-ItemProperty -LiteralPath $addInKey -Name LoadBehavior -Value $originalLoadBehavior
    Write-Output "MARKETVIEW_LOADBEHAVIOR_RESTORED=$originalLoadBehavior"
}
