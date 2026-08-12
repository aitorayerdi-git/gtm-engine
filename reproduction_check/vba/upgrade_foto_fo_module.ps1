param([Parameter(Mandatory = $true)][string]$Path)
$ErrorActionPreference = 'Stop'
$module = Join-Path $PSScriptRoot 'modFotoFOUpdate.bas'
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.EnableEvents = $false
$workbook = $null
try {
    $workbook = $excel.Workbooks.Open($Path, 0, $false)
    $existing = $null
    try { $existing = $workbook.VBProject.VBComponents.Item('modFotoFOUpdate') } catch {}
    if ($null -ne $existing) { $workbook.VBProject.VBComponents.Remove($existing) }
    [void]$workbook.VBProject.VBComponents.Import($module)
    $workbook.Save()
    $workbook.Close($true); $workbook = $null
    Write-Output "UPDATED=$Path"
}
finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    $excel.Quit()
}
