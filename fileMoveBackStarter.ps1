If(!(test-path -PathType container ".\log\"))
{
      New-Item -ItemType Directory -Path ".\log\"
}

$logDatesNotMatchingPath = ".\log\filesNotMatchingDates.txt"
$logFilesNotEditedPath = ".\log\filesNotEdited.txt"
$logFileError = ".\log\fileError.txt"

Remove-Item -Path $logDatesNotMatchingPath -ErrorAction SilentlyContinue
Remove-Item -Path $logFilesNotEditedPath -ErrorAction SilentlyContinue
Remove-Item -Path $logFileError -ErrorAction SilentlyContinue

.\fileMoveBack.ps1 -sourcePath "*Path of edited Files*" -destinationPath "*Path of Original Files*"