param([Parameter(Mandatory = $true)][string]$Path)
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$sourcePath = Join-Path $root 'verification\Gas_Trading_Model 070826.xlsm'
$Path = [System.IO.Path]::GetFullPath($Path)
$module = Join-Path $PSScriptRoot 'modMarketDataRefresh.bas'
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.EnableEvents = $false
$source = $null
$target = $null
try {
    $source = $excel.Workbooks.Open($sourcePath, 0, $true)
    $target = $excel.Workbooks.Open($Path, 0, $false)
    $excel.Calculation = -4135 # xlCalculationManual: preserve cached MarketView values while copying.
    foreach ($sheet in @($target.Worksheets)) {
        if ($sheet.Name -eq 'Historical_DA') { $sheet.Delete(); break }
    }
    $source.Worksheets.Item('Historical_DA').Copy($target.Worksheets.Item('CURVE PRICES'))
    $target.Worksheets.Item('Historical_DA').Tab.Color = 49407

    $existing = $null
    try { $existing = $target.VBProject.VBComponents.Item('modMarketDataRefresh') } catch {}
    if ($null -ne $existing) { $target.VBProject.VBComponents.Remove($existing) }
    [void]$target.VBProject.VBComponents.Import($module)

    $sheet = $target.Worksheets.Item('MANUAL CHANGES')
    foreach ($table in @($sheet.ListObjects)) {
        if ($table.Name -eq 'tblMarketDataStatus') { $table.Unlist() }
    }
    $sheet.Range('X2:AA8').ClearContents()
    $sheet.Range('X2:AA2').Merge()
    $sheet.Range('X2').Value = 'MARKET DATA STATUS'
    $sheet.Range('X3').Value = 'Source'
    $sheet.Range('Y3').Value = 'Last Refresh At'
    $sheet.Range('Z3').Value = 'Latest Data Date'
    $sheet.Range('AA3').Value = 'Status'
    $sheet.Range('X4').Value = 'Reuters'
    $sheet.Range('X5').Value = 'MarketView'
    $sourceTimestamp = (Get-Item -LiteralPath $sourcePath).LastWriteTime
    $sheet.Range('Y4:Y5').Value = $sourceTimestamp
    $statusTable = $sheet.ListObjects.Add(1, $sheet.Range('X3:AA5'), $null, 1)
    $statusTable.Name = 'tblMarketDataStatus'
    $statusTable.TableStyle = 'TableStyleMedium2'
    $sheet.Range('Y4:Y5').NumberFormat = 'dd/mm/yyyy hh:mm:ss'
    $sheet.Range('Z4:Z5').NumberFormat = 'dd/mm/yyyy'
    $sheet.Columns.Item('X').ColumnWidth = 16
    $sheet.Columns.Item('Y').ColumnWidth = 22
    $sheet.Columns.Item('Z').ColumnWidth = 18
    $sheet.Columns.Item('AA').ColumnWidth = 28

    foreach ($button in @($sheet.Buttons())) {
        if ($button.Name -in @('btnUpdateReutersMarketView', 'btnRefreshMarketDataStatus')) { $button.Delete() }
    }
    $anchor = $sheet.Range('X7:Z8')
    $button = $sheet.Buttons().Add($anchor.Left, $anchor.Top, $anchor.Width, $anchor.Height)
    $button.Name = 'btnUpdateReutersMarketView'
    $button.Characters().Text = 'UPDATE REUTERS + MARKETVIEW'
    $button.OnAction = 'Actualizar_Reuters_MarketView'
    $button.Font.Bold = $true
    $button.Font.Size = 10
    $anchor = $sheet.Range('AA7:AA8')
    $button = $sheet.Buttons().Add($anchor.Left, $anchor.Top, $anchor.Width, $anchor.Height)
    $button.Name = 'btnRefreshMarketDataStatus'
    $button.Characters().Text = 'REFRESH STATUS'
    $button.OnAction = 'Actualizar_Control_Datos_Mercado'
    $button.Font.Bold = $true
    $button.Font.Size = 9

    $excel.Run("'" + $target.Name + "'!Inicializar_Control_Datos_Mercado")
    $target.Save()
    Write-Output "INSTALLED=$Path"
}
finally {
    if ($target) { $target.Close($false) }
    if ($source) { $source.Close($false) }
    $excel.Quit()
}
