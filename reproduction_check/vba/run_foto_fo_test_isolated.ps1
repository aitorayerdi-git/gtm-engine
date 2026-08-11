param([Parameter(Mandatory = $true)][string]$Path)
$ErrorActionPreference = 'Stop'
$addInKey = 'HKCU:\Software\Microsoft\Office\Excel\Addins\AStorm.Ui.StormUi'
$originalLoadBehavior = (Get-ItemProperty -LiteralPath $addInKey -Name LoadBehavior).LoadBehavior
try {
    Set-ItemProperty -LiteralPath $addInKey -Name LoadBehavior -Value 0
    & (Join-Path $PSScriptRoot 'test_foto_fo_macro.ps1') -Path $Path
}
finally {
    Set-ItemProperty -LiteralPath $addInKey -Name LoadBehavior -Value $originalLoadBehavior
    Write-Output "MARKETVIEW_LOADBEHAVIOR_RESTORED=$originalLoadBehavior"
}
