Get-Process -Name *Excel* | Stop-Process

Write-Host "Start: $(Get-Date)"

$stopwatch =  [system.diagnostics.stopwatch]::StartNew()

$splitFiles = Get-ChildItem -Path ".\splits\" -Filter "split*.txt" -File

$processList = @()

Start-Sleep 5

for($k = 0 ; $k -lt $splitFiles.Length ; $k++)
{
    $processList += Start-Process powershell -ArgumentList ".\ExcelVbaModifier.ps1",".\splits\$(($splitFiles[$k]).Name)" -PassThru #-WindowStyle Hidden
}

$processList | Wait-Process

$stopwatch.Stop()

Write-Host "End: $(Get-Date)"

$stopwatch