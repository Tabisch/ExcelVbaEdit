Param (
    # Param1 help description
    [Parameter(ValueFromPipeline=$true)]
    [string[]]
    $items
)

begin {
    Get-ChildItem -Filter ".\splits\split*.txt" | Remove-Item -Force

    If(!(test-path -PathType container ".\splits\"))
    {
        New-Item -ItemType Directory -Path ".\splits\"
    }

    Start-Sleep 5
    
    $parts = (Get-CimInstance -Class Win32_processor).NumberOfCores

    $dateien = @()
}

process {
    foreach($item in $items)
    {
        $dateien += Get-ChildItem -Path $item -File -Filter "*.xls"
    }
}

end {
    for($k = 0 ; $k -lt $dateien.length ; $k++)
    {
        Write-host "File : $($k + 1) of $($dateien.length) : $([math]::Round(($k + 1)/$dateien.length *100,2))% : $(($dateien[$k]).Fullname)"
        ($dateien[$k]).FullName | Out-File -FilePath ".\splits\split$($k % $parts).txt" -Append -Encoding utf8
    }
}