param([Parameter(Mandatory = $true)][string]$Directory)
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    foreach ($name in @('test_manual.xlsx', 'test_log.xlsx', 'test_state.xlsx')) {
        $path = Join-Path $Directory $name
        try {
            $wb = $excel.Workbooks.Open($path, 0, $true)
            Write-Output "OPENED=$name"
            $wb.Close($false)
        }
        catch {
            Write-Output "FAILED=$name :: $($_.Exception.Message)"
        }
    }
}
finally { $excel.Quit() }
