param([Parameter(Mandatory = $true)][string]$Path)
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$workbook = $null
try {
    $workbook = $excel.Workbooks.Open($Path, 0, $true)
    foreach ($component in @($workbook.VBProject.VBComponents)) {
        if ($component.Name -eq 'modFotoFOUpdate') {
            $code = $component.CodeModule.Lines(1, $component.CodeModule.CountOfLines)
            Write-Output "LINES=$($component.CodeModule.CountOfLines)"
            Write-Output "HAS_REBUILD=$($code.Contains('RebuildHistoricalFlows'))"
            Write-Output "HAS_OLD_DETAIL=$($code.Contains('Carga completa:'))"
        }
    }
}
finally {
    if ($workbook) { $workbook.Close($false) }
    $excel.Quit()
}
