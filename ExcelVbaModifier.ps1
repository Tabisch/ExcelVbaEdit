Param(
    $Datei
)

Write-host "Loading file $($Datei)"

$filePaths = Get-Content $Datei -Encoding utf8

$files = $filePaths | ForEach-Object{
    try
    {
        return Get-Item -Path $_
    }
    catch
    {

    }
}

$vbaCode = Get-Content -Path ".\new.vba" -Raw -Encoding utf8

$vbaCode = $vbaCode + "'$((Get-Date).ToString("yyyyMMdd"))'"

$excelApplication = New-Object -ComObject ("Excel.Application")
$excelApplication.Visible = $false
$excelApplication.DisplayAlerts = $false

for($k = 0 ; $k -lt $files.length ; $k++)
{
    Write-host "File : $($k + 1) of $($files.length) : $([math]::Round(($k + 1)/$files.length *100,2))% : $(($files[$k]).Fullname)"
    $lastwrite = ($files[$k]).LastWriteTime 

    $workbook = $null

    if(!$excelApplication.Ready)
    {
        $excelApplication = New-Object -ComObject ("Excel.Application")
    }

    while($True -and $excelApplication.Ready)
    {
        try {
            $workbook = $excelApplication.Workbooks.Open($files.Fullname[$k])
            break
        }
        catch {
            if(!$excelApplication.Ready)
            {
                Write-Host "Excel crashed"
                break
            }
            Write-Host "Retry Open"
        }
    }
        
    $component = $workbook.VBProject.VBComponents | Where-Object{ $_.Name -eq "*Modulename*" }

    $workbook.VBProject.VBComponents.Remove($component)
    
    $module = $workbook.VBProject.VBComponents.Add(1)
    
    $module.CodeModule.AddFromString($vbaCode)

    while($True -and $excelApplication.Ready)
    {
        try {
            $workbook.Save()
            break
        }
        catch {
            if(!$excelApplication.Ready)
            {
                Write-Host "Excel crashed"
                break
            }
            Write-Host "Retry Save"
        }
    }
    
    $workbook.Close($True)

    while($True)
    {
        try {
            ($files[$k]).LastWriteTime = $lastwrite
            break
        }
        catch {
            Write-Host "Retry LastWrite"
            Write-Warning $Error[0]
            $workbook.Close($True)
            Start-Sleep 5
            $Error.Clear()
        }
    }

    while($True -and $excelApplication.Ready)
    {
        try {
            ($files[$k]).Fullname | Out-File -FilePath .\finished.txt -Append -Encoding utf8
            break
        }
        catch {
            Write-Host "Logging"
            Start-Sleep 1
        }
    }

    if(!$excelApplication.Ready)
    {
        $k--
        Write-host "Retrying File : $(($files[$k]).Fullname)"
        continue
    }
}

$excelApplication.Quit()
