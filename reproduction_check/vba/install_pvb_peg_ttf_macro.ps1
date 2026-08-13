param([Parameter(Mandatory = $true)][string]$Path)
$ErrorActionPreference = 'Stop'
$Path = [System.IO.Path]::GetFullPath($Path)
$module = Join-Path $PSScriptRoot 'modPvbPegTtfUpdate.bas'
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.EnableEvents = $false
$workbook = $null
try {
    $workbook = $excel.Workbooks.Open($Path, 0, $false)
    foreach ($componentName in @('modPvbPegTtfUpdate', 'Módulo4')) {
        $existing = $null
        try { $existing = $workbook.VBProject.VBComponents.Item($componentName) } catch {}
        if ($null -ne $existing) { $workbook.VBProject.VBComponents.Remove($existing) }
    }
    [void]$workbook.VBProject.VBComponents.Import($module)

    $sheet = $workbook.Worksheets.Item('MANUAL CHANGES')
    foreach ($button in @($sheet.Buttons())) {
        if ($button.Name -eq 'btnUpdatePvbPegTtf') { $button.Delete() }
    }
    $anchor = $sheet.Range('S6:V7')
    $button = $sheet.Buttons().Add($anchor.Left, $anchor.Top, $anchor.Width, $anchor.Height)
    $button.Name = 'btnUpdatePvbPegTtf'
    $button.Characters().Text = 'UPDATE PVB/TTF + PEG/TTF'
    $button.OnAction = 'Actualizar_PVB_PEG_TTF'
    $button.Font.Bold = $true
    $button.Font.Size = 10

    $workbook.Save()
    $workbook.Close($true)
    $workbook = $null
    Write-Output "INSTALLED=$Path"
}
finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    $excel.Quit()
}
