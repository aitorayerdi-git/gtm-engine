param([Parameter(Mandatory = $true)][string]$Path)
$ErrorActionPreference = 'Stop'
$Path = [IO.Path]::GetFullPath($Path)
$key = 'HKCU:\Software\Microsoft\Office\Excel\Addins\AStorm.Ui.StormUi'
$old = (Get-ItemProperty -LiteralPath $key -Name LoadBehavior).LoadBehavior
Set-ItemProperty -LiteralPath $key -Name LoadBehavior -Value 0
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.EnableEvents = $false
$bootstrap = $excel.Workbooks.Add()
$excel.Calculation = -4135
$book = $null
try {
    $book = $excel.Workbooks.Open($Path, 0, $false)
    $bootstrap.Close($false)
    $bootstrap = $null
    foreach ($sheetName in @('PVB-TTF', 'PEG-TTF')) {
        $sheet = $book.Worksheets.Item($sheetName)
        $start = [datetime]::FromOADate([double]$sheet.Cells.Item(3, 2).Value2)
        $lastRow = $sheet.Cells.Item($sheet.Rows.Count, 2).End(-4162).Row
        if ($lastRow -lt 412) { $lastRow = 412 }
        $values = New-Object 'object[,]' ($lastRow - 2), 1
        for ($index = 0; $index -lt ($lastRow - 2); $index++) {
            $values[$index, 0] = $start.AddDays($index)
        }
        $range = $sheet.Range($sheet.Cells.Item(3, 2), $sheet.Cells.Item($lastRow, 2))
        $range.Value = $values
        $range.NumberFormat = 'dd/mm/yyyy'
    }
    $book.ForceFullCalculation = $false
    $book.Save()
    $book.Close($false)
    $book = $null
}
finally {
    if ($null -ne $book) { try { $book.Close($false) } catch {} }
    if ($null -ne $bootstrap) { try { $bootstrap.Close($false) } catch {} }
    try { $excel.Quit() } catch {}
    Set-ItemProperty -LiteralPath $key -Name LoadBehavior -Value $old
}
