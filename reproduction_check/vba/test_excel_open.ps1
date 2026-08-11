param([Parameter(Mandatory = $true)][string]$Path)
$ErrorActionPreference = 'Stop'
$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Open($Path, 0, $true)
    Write-Output "OPENED=$($workbook.Name)"
    $workbook.Close($false)
    $workbook = $null
}
finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    if ($null -ne $excel) { $excel.Quit() }
}
