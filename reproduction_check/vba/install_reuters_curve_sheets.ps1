param(
    [string]$TargetPath
)
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$sourcePath = Join-Path $root 'verification\Gas_Trading_Model 070826.xlsm'
if (-not $TargetPath) { $TargetPath = Join-Path $root 'reproduction_check\Input data.xlsm' }
$TargetPath = [System.IO.Path]::GetFullPath($TargetPath)
$names = @('CURVES Calendar','TTF','Brent Dated','HH','EURF','PVB-TTF','PEG-TTF')
$excel = $null; $source = $null; $target = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.EnableEvents = $false
    $source = $excel.Workbooks.Open($sourcePath, 0, $true)
    $target = $excel.Workbooks.Open($TargetPath, 0, $false)
    foreach ($name in $names) {
        foreach ($sheet in @($target.Worksheets)) { if ($sheet.Name -eq $name) { $sheet.Delete(); break } }
    }
    # Copy as one group so internal CURVES Calendar references remain internal.
    $source.Worksheets.Item($names).Copy($target.Worksheets.Item('CURVE PRICES'))
    foreach ($name in $names) { $target.Worksheets.Item($name).Tab.Color = 49407 }
    $target.Save()
    Write-Output "Installed Reuters curve staging sheets in $TargetPath"
}
finally {
    if ($target) { $target.Close($false) }; if ($source) { $source.Close($false) }; if ($excel) { $excel.Quit() }
}
