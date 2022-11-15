Param(
    $sourcePath,
    $destinationPath
)

$sourceFiles = @()

$sourceFiles += Get-ChildItem -Path $sourcePath -Filter "*.xls"

for($i = 0 ; $i -lt $sourceFiles.length ; $i++)
{
    Write-host "File : $($i + 1) of $($sourceFiles.length) : $([math]::Round(($i + 1)/$sourceFiles.length *100,2))% : $(($sourceFiles[$i]).Fullname)"

    try {
        $destinationFile = Get-Item -Path "$($destinationPath)$(($sourceFiles[$i]).Name)" -ErrorAction Stop -WarningAction Stop
        $destinationFilePath = $destinationFile.FullName

        if($destinationFile.LastWriteTime -eq ($sourceFiles[$i]).LastWriteTime)
        {
            #Write-Host "Destination: $((Get-FileHash -Path $destinationFile.FullName).Hash) SourceHash: $((Get-FileHash -Path $sourceFiles[$i].FullName).Hash)"
            if((Get-FileHash -Path $destinationFile.FullName).Hash -eq (Get-FileHash -Path $sourceFiles[$i].FullName).Hash)
            {
                #Write-Host "Skipping: $($destinationFilePath)"
                $destinationFilePath | Out-File -FilePath $logFilesNotEditedPath -Append -Encoding utf8
                continue
            }

            #Start-Sleep 1

            #Write-Host "Replacing: $($destinationFilePath)"

            Remove-Item -Path $destinationFile.FullName
            Copy-Item -Path $sourceFiles[$i].FullName -Destination $destinationFilePath

            Remove-Item -Path $sourceFiles[$i].FullName
        }
        else 
        {
            #Write-Host "Skipping: $($destinationFilePath)"
            $destinationFilePath | Out-File -FilePath $logDatesNotMatchingPath -Append -Encoding utf8
            continue
        }
    }
    catch {
        "$($destinationPath)$(($sourceFiles[$i]).Name)" | Out-File -FilePath $logFileError -Append -Encoding utf8
        Write-Warning "Error: $($destinationPath)$(($sourceFiles[$i]).Name)"
    }
}