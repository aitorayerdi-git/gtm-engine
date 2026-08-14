param([Parameter(Mandatory = $true)][string]$Path)
$ErrorActionPreference = 'Stop'
$Path = [System.IO.Path]::GetFullPath($Path)
$modules = @(
    (Join-Path $PSScriptRoot 'modPvbPegTtfUpdate.bas'),
    (Join-Path $PSScriptRoot 'modMarketDataRefresh.bas')
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
try {
    $bootstrap = $excel.Workbooks.Add()
    $excel.Calculation = -4135
    $workbook = $excel.Workbooks.Open($Path, 0, $false)
    $bootstrap.Close($false)
    $bootstrap = $null
    if ($workbook.ReadOnly) { throw "Excel abrió el libro en modo de solo lectura: $Path" }

    foreach ($module in $modules) {
        $componentName = [System.IO.Path]::GetFileNameWithoutExtension($module)
        $existing = $null
        try { $existing = $workbook.VBProject.VBComponents.Item($componentName) } catch {}
        if ($null -ne $existing) { $workbook.VBProject.VBComponents.Remove($existing) }
        [void]$workbook.VBProject.VBComponents.Import($module)
    }

    $sheet = $workbook.Worksheets.Item('MANUAL CHANGES')
    $sheet.Range('Y3').Value = 'Last Successful Check At'
    foreach ($button in @($sheet.Buttons())) {
        if ($button.Name -in @('btnUpdateReutersMarketView', 'btnRefreshMarketDataStatus',
                'btnUpdateReuters', 'btnCheckMarketView')) { $button.Delete() }
    }
    $anchor = $sheet.Range('X7:Y8')
    $button = $sheet.Buttons().Add($anchor.Left, $anchor.Top, $anchor.Width, $anchor.Height)
    $button.Name = 'btnUpdateReuters'
    $button.Characters().Text = 'UPDATE + CHECK REUTERS'
    $button.OnAction = 'Actualizar_Reuters'
    $button.Font.Bold = $true
    $button.Font.Size = 9
    $anchor = $sheet.Range('Z7:AA8')
    $button = $sheet.Buttons().Add($anchor.Left, $anchor.Top, $anchor.Width, $anchor.Height)
    $button.Name = 'btnCheckMarketView'
    $button.Characters().Text = 'CHECK MARKETVIEW'
    $button.OnAction = 'Comprobar_MarketView'
    $button.Font.Bold = $true
    $button.Font.Size = 9

    $workbook.Save()
    $workbook.Close($true)
    $workbook = $null
    Write-Output "INSTALLED=$Path"
}
finally {
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch {} }
    if ($null -ne $bootstrap) { try { $bootstrap.Close($false) } catch {} }
    try { $excel.Quit() } catch {}
}
