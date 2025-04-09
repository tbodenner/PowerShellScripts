# each of these paths will be checked for old files
$FilePaths = @(
    'C:\inetpub\wwwroot\StreemWebRoot\StreemFiles',
    'C:\inetpub\wwwroot\StreemWebRoot\StreemInFiles',
    'C:\inetpub\wwwroot\StreemWebRoot\StreemOutFiles'
)
# start a transcript
Start-Transcript -Path 'C:\Temp\Remove-ExcessFaxFiles.txt' | Out-Null
# we will add every file to this array
$AllFiles = @()
# get the files from each path
Write-Host 'Getting all files...'
foreach ($Path in $FilePaths) {
    $AllFiles += Get-ChildItem -Path $Path
}
Write-Host "File Count: $($AllFiles.Count.ToString("N0"))"
# set our target date 45 days before today
$TargetDate = (Get-Date).AddDays(-45)
Write-Host "Target Date: $($TargetDate)"
# count our files removed
$RemovedFileCount = 0
$FileSizeCount = 0
# check each file's last write date, and delete the file if older than our date
Write-Host 'Removing files...'
foreach ($File in $AllFiles) {
    if ($File.LastWriteTime -lt $TargetDate) {
        Remove-Item -Path $File.FullName -Force -ErrorAction SilentlyContinue
        $FileSizeCount += $File.Length
        $RemovedFileCount += 1
    }
}
Write-Host "Removed Files:"
Write-Host "  Count: $($RemovedFileCount.ToString("N0"))"
# calculate the number of gigabytes
$FileSizeGig = $FileSizeCount / [Math]::Pow(1024,3)
Write-Host "   Size: $($FileSizeGig.ToString("0.00")) GB"
Write-Host 'Done.'
# stop our transcript
Stop-Transcript | Out-Null
