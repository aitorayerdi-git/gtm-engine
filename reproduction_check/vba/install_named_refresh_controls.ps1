param([Parameter(Mandatory = $true)][string]$Path)
$ErrorActionPreference = 'Stop'
$Path = [System.IO.Path]::GetFullPath($Path)
$addInKey = 'HKCU:\Software\Microsoft\Office\Excel\Addins\AStorm.Ui.StormUi'
$originalLoadBehavior = (Get-ItemProperty -LiteralPath $addInKey -Name LoadBehavior -ErrorAction Stop).LoadBehavior
Set-ItemProperty -LiteralPath $addInKey -Name LoadBehavior -Value 0
$backup = Join-Path ([System.IO.Path]::GetDirectoryName($Path)) 'Input data v2 pre named refresh controls 20260814.xlsm'
if (-not (Test-Path -LiteralPath $backup)) { Copy-Item -LiteralPath $Path -Destination $backup }

$modules = @(
    (Join-Path $PSScriptRoot 'modMarketDataRefresh.bas'),
    (Join-Path $PSScriptRoot 'modFotoFOUpdate.bas'),
    (Join-Path $PSScriptRoot 'modPvbPegTtfUpdate.bas'),
    (Join-Path $PSScriptRoot 'modGtmOutput.bas')
)
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.EnableEvents = $false
$excel.ScreenUpdating = $false
$excel.AskToUpdateLinks = $false
$excel.AutomationSecurity = 3
$bootstrap = $null
$workbook = $null
$stage = 'starting Excel'
try {
    $stage = 'creating bootstrap workbook'
    $bootstrap = $excel.Workbooks.Add()
    $excel.Calculation = -4135
    $stage = 'opening target workbook'
    $workbook = $excel.Workbooks.Open($Path, 0, $false)
    $bootstrap.Close($false)
    $bootstrap = $null
    if ($workbook.ReadOnly) { throw "Workbook opened read-only: $Path" }
    Start-Sleep -Seconds 10

    $stage = 'importing VBA modules'
    foreach ($module in $modules) {
        $componentName = [System.IO.Path]::GetFileNameWithoutExtension($module)
        $existing = $null
        try { $existing = $workbook.VBProject.VBComponents.Item($componentName) } catch {}
        if ($null -ne $existing) { $workbook.VBProject.VBComponents.Remove($existing) }
        [void]$workbook.VBProject.VBComponents.Import($module)
    }

    $stage = 'rebuilding status table'
    $sheet = $workbook.Worksheets.Item('MANUAL CHANGES')
    $stage = 'reading existing status table'
    $oldTable = $sheet.ListObjects.Item('tblMarketDataStatus')
    $stage = 'resizing existing status table'
    $oldTable.Resize($sheet.Range('J5:M11'))
    $table = $oldTable
    $stage = 'setting status table style'
    $table.TableStyle = 'TableStyleMedium2'
    $headers = @('Source', 'Last Successful Action At', 'Latest Data Date', 'Status')
    $stage = 'writing status headers'
    for ($column = 1; $column -le 4; $column++) { $table.HeaderRowRange.Cells.Item(1, $column).Value2 = $headers[$column - 1] }
    $sources = @('Reuters', 'MarketView', 'Foto FO', 'PVB-TTF', 'PEG-TTF', 'GTM Output')
    $stage = 'writing status sources'
    for ($row = 1; $row -le 6; $row++) { $sheet.Cells.Item(5 + $row, 10).Value = [string]$sources[$row - 1] }
    $stage = 'formatting status table'
    $table.ListColumns('Last Successful Action At').DataBodyRange.NumberFormat = 'dd/mm/yyyy hh:mm:ss'
    $table.ListColumns('Latest Data Date').DataBodyRange.NumberFormat = 'dd/mm/yyyy'
    $sheet.Columns.Item('J').ColumnWidth = 16
    $sheet.Columns.Item('K').ColumnWidth = 25
    $sheet.Columns.Item('L').ColumnWidth = 18
    $sheet.Columns.Item('M').ColumnWidth = 30

    $stage = 'rebuilding buttons'
    foreach ($button in @($sheet.Buttons())) {
        if ($button.Name -in @('btnUpdateReutersMarketView', 'btnRefreshMarketDataStatus',
                'btnUpdateReuters', 'btnCheckMarketView', 'btnUpdateMarketView',
                'btnCheckReutersMarketView', 'btnUpdateFotoFO', 'btnUpdatePvbPegTtf',
                'btnGenerateGtmOutput')) {
            $button.Delete()
        }
    }
    $specs = @(
        @('btnUpdateMarketView', 'UPDATE MARKETVIEW', 'Actualizar_MarketView', 'J12:K13'),
        @('btnCheckReutersMarketView', 'CHECK REUTERS + MARKETVIEW', 'Comprobar_Reuters_MarketView', 'L12:M13'),
        @('btnUpdateFotoFO', 'UPDATE FOTO FO', 'UpdateFotoFO', 'J15:K16'),
        @('btnUpdatePvbPegTtf', 'UPDATE PVB/TTF + PEG/TTF', 'Actualizar_PVB_PEG_TTF', 'L15:M16'),
        @('btnGenerateGtmOutput', 'GENERATE GTM OUTPUT', 'Generar_GTM_Output', 'J18:M19')
    )
    foreach ($spec in $specs) {
        $anchor = $sheet.Range($spec[3])
        $button = $sheet.Buttons().Add($anchor.Left, $anchor.Top, $anchor.Width, $anchor.Height)
        $button.Name = $spec[0]
        $button.Characters().Text = $spec[1]
        $button.OnAction = $spec[2]
        $button.Font.Bold = $true
        $button.Font.Size = 9
    }

    $stage = 'saving workbook'
    $workbook.ForceFullCalculation = $false
    $workbook.Save()
    $workbook.Close($false)
    $workbook = $null
    Write-Output "INSTALLED=$Path"
    Write-Output "BACKUP=$backup"
}
catch {
    throw "$stage`: $($_.Exception.Message)"
}
finally {
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch {} }
    if ($null -ne $bootstrap) { try { $bootstrap.Close($false) } catch {} }
    try { $excel.Quit() } catch {}
    Set-ItemProperty -LiteralPath $addInKey -Name LoadBehavior -Value $originalLoadBehavior
}
