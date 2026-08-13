param(
    [Parameter(Mandatory = $true)][string]$TargetPath,
    [Parameter(Mandatory = $true)][string]$ExportDir
)
$ErrorActionPreference = 'Stop'
$TargetPath = [System.IO.Path]::GetFullPath($TargetPath)
$ExportDir = [System.IO.Path]::GetFullPath($ExportDir)

function Write-CsvTable($sheet, $tableName, $csvPath, $columnCount) {
    $rows = @(Import-Csv -LiteralPath $csvPath -Header (1..$columnCount | ForEach-Object { "C$_" }))
    $lastUsed = [Math]::Max(5, $sheet.UsedRange.Rows.Count)
    $sheet.Range($sheet.Cells.Item(5, 1), $sheet.Cells.Item($lastUsed, $columnCount)).ClearContents()
    if ($rows.Count -gt 0) {
        $data = New-Object 'object[,]' $rows.Count, $columnCount
        for ($r = 0; $r -lt $rows.Count; $r++) {
            for ($c = 0; $c -lt $columnCount; $c++) {
                $value = $rows[$r].("C$($c + 1)")
                if ($c -eq 0 -or ($columnCount -eq 8 -and $c -eq 2) -or $c -eq ($columnCount - 1)) {
                    $parsed = [datetime]::MinValue
                    if ([datetime]::TryParse($value, [ref]$parsed)) { $value = $parsed }
                } elseif (($columnCount -eq 8 -and $c -eq 3) -or ($columnCount -eq 5 -and $c -eq 2)) {
                    $value = [double]::Parse($value, [Globalization.CultureInfo]::InvariantCulture)
                }
                $data[$r, $c] = $value
            }
        }
        $sheet.Range($sheet.Cells.Item(5, 1), $sheet.Cells.Item(4 + $rows.Count, $columnCount)).Value2 = $data
        $sheet.Range($sheet.Cells.Item(5, 1), $sheet.Cells.Item(4 + $rows.Count, 1)).NumberFormat = 'yyyy-mm-dd'
        if ($columnCount -eq 8) {
            $sheet.Range($sheet.Cells.Item(5, 3), $sheet.Cells.Item(4 + $rows.Count, 3)).NumberFormat = 'yyyy-mm-dd'
        }
        $sheet.Range($sheet.Cells.Item(5, $columnCount), $sheet.Cells.Item(4 + $rows.Count, $columnCount)).NumberFormat = 'yyyy-mm-dd hh:mm:ss'
    }
    $sheet.ListObjects.Item($tableName).Resize($sheet.Range($sheet.Cells.Item(4, 1), $sheet.Cells.Item([Math]::Max(5, 4 + $rows.Count), $columnCount)))
    return $rows.Count
}

$excel = $null; $workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.EnableEvents = $false
    $workbook = $excel.Workbooks.Open($TargetPath, 0, $false)
    $curveCount = Write-CsvTable $workbook.Worksheets.Item('CURVE PRICES') 'tblCurvePrices' (Join-Path $ExportDir 'curve_prices.csv') 8
    $fxCount = Write-CsvTable $workbook.Worksheets.Item('FX RATES') 'tblFxRates' (Join-Path $ExportDir 'fx_rates.csv') 5
    $workbook.Save()
    Write-Output "Published $curveCount curve prices and $fxCount FX rates in $TargetPath"
}
finally {
    if ($workbook) { $workbook.Close($false) }
    if ($excel) { $excel.Quit() }
}
