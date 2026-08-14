param(
    [string]$TargetPath = (Join-Path $PSScriptRoot '..\Input data.xlsm')
)

$ErrorActionPreference = 'Stop'
$target = (Resolve-Path -LiteralPath $TargetPath).Path
$repository = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..')).Path
if (-not $target.StartsWith($repository + '\', [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "Target workbook is outside the repository: $target"
}
if ([System.IO.Path]::GetExtension($target) -ne '.xlsm') {
    throw "Target workbook must be an .xlsm file: $target"
}

$backup = Join-Path ([System.IO.Path]::GetDirectoryName($target)) 'Input data pre PVB MO market date 20260814.xlsm'
if (-not (Test-Path -LiteralPath $backup)) {
    Copy-Item -LiteralPath $target -Destination $backup
}

$excel = $null
$workbook = $null
$sheet = $null
$table = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $workbook = $excel.Workbooks.Open($target, 0, $false)
    $sheet = $workbook.Worksheets.Item('MANUAL CHANGES')
    $table = $sheet.ListObjects.Item('tblManualDates')

    $existing = $null
    foreach ($row in $table.ListRows) {
        if ([string]$row.Range.Cells.Item(1, 1).Value2 -eq 'Market Date INITIAL POSITION PVB MO') {
            $existing = $row
            break
        }
    }

    if ($null -eq $existing) {
        $firstRow = $table.Range.Row
        $firstColumn = $table.Range.Column
        $lastRow = $firstRow + $table.Range.Rows.Count
        $lastColumn = $firstColumn + $table.Range.Columns.Count - 1
        $expandedRange = $sheet.Range(
            $sheet.Cells.Item($firstRow, $firstColumn),
            $sheet.Cells.Item($lastRow, $lastColumn)
        )
        $table.Resize($expandedRange)
        $newRow = $table.ListRows.Item($table.ListRows.Count)
        $newRow.Range.Cells.Item(1, 1).Value2 = 'Market Date INITIAL POSITION PVB MO'
        $newRow.Range.Cells.Item(1, 2).ClearContents()
        $newRow.Range.Cells.Item(1, 2).NumberFormat = 'yyyy-mm-dd'
        $newRow.Range.Cells.Item(1, 3).Value2 = 'Fecha que muestra la informacion de MO a su cierre. Solo se anaden a Exposure trades con Trade Date posterior; los de fecha igual o anterior ya estan incluidos.'
    }

    $workbook.Save()
    Write-Output "Installed Market Date INITIAL POSITION PVB MO in $target"
    Write-Output "Backup: $backup"
}
finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    if ($null -ne $excel) { $excel.Quit() }
    foreach ($object in @($table, $sheet, $workbook, $excel)) {
        if ($null -ne $object) {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($object)
        }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
