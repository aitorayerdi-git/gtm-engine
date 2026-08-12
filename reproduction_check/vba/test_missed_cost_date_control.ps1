param([Parameter(Mandatory = $true)][string]$BasePath)
$ErrorActionPreference = 'Stop'
$testPath = Join-Path (Split-Path -Parent $BasePath) 'missed_cost_date_control.xlsm'
Copy-Item -LiteralPath $BasePath -Destination $testPath -Force

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.EnableEvents = $false
$excel.AskToUpdateLinks = $false
$control = $excel.Workbooks.Add()
$excel.Calculation = -4135
$workbook = $null
try {
    try { $excel.COMAddIns.Item('AStorm.Ui.StormUi').Connect = $false } catch {}
    $workbook = $excel.Workbooks.Open($testPath, 0, $false)
    $control.Close($false)
    $control = $null

    $targetRaw = $workbook.Worksheets.Item('MANUAL CHANGES').Range('F5').Value2
    $targetDate = if ($targetRaw -is [datetime]) { $targetRaw } else { [datetime]::FromOADate([double]$targetRaw) }
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
    $missedDate = $priorDates[-1]
    $stateDate = $priorDates[-2]
    $stateSheet = $workbook.Worksheets.Item('FOTO FO STATE')
    $lastStateRow = $stateSheet.Cells($stateSheet.Rows.Count, 1).End(-4162).Row
    $stateSheet.Range("B2:B$lastStateRow").Value2 = $stateDate.ToOADate()
    $workbook.Worksheets.Item('COSTS').Range('A5:U504').ClearContents()

    $validationFormula = [string]$workbook.Worksheets.Item('COSTS').Range('R5').Validation.Formula1
    if ($validationFormula -ne '=$X$1:$X$2') { throw "Expected range-backed MANUAL/BASELINE dropdown, got '$validationFormula'." }
    if ([string]$workbook.Worksheets.Item('COSTS').Range('X1').Value2 -ne 'MANUAL' -or [string]$workbook.Worksheets.Item('COSTS').Range('X2').Value2 -ne 'BASELINE') { throw 'Dropdown source values are invalid.' }

    $excel.Run("'$($workbook.Name)'!UpdateFotoFO_LocalTest")
    $status = [string]$workbook.Worksheets.Item('MANUAL CHANGES').Range('M6').Text
    if (-not $status.StartsWith('ERROR') -or $status -notlike '*Market Dates sin actualizar*') {
        throw "Expected missed-date error, got '$status'."
    }
    $savedStateRaw = $stateSheet.Range('B2').Value2
    $savedStateDate = if ($savedStateRaw -is [datetime]) { $savedStateRaw } else { [datetime]::FromOADate([double]$savedStateRaw) }
    if ($savedStateDate.Date -ne $stateDate.Date) { throw 'State advanced despite the missed-date block.' }

    Write-Output "MISSED_DATE=$($missedDate.ToString('yyyy-MM-dd'))"
    Write-Output "STATUS=$status"
    Write-Output 'DROPDOWN=MANUAL,BASELINE'
    Write-Output 'STATE_NOT_ADVANCED=OK'
}
finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    if ($null -ne $control) { $control.Close($false) }
    $excel.Quit()
}
