# each of these paths will be checked for old files
$FilePaths = @(
    'C:\inetpub\wwwroot\StreemWebRoot\StreemFiles',
    'C:\inetpub\wwwroot\StreemWebRoot\StreemInFiles',
    'C:\inetpub\wwwroot\StreemWebRoot\StreemOutFiles'
)
# start a transcript
Start-Transcript -Path 'C:\Temp\Remove-ExcessFaxFiles.txt' -Force | Out-Null
# we will add every file to this list
$AllFiles = [System.Collections.Generic.List[System.IO.FileSystemInfo]]::new()
# get the files from each path
Write-Host 'Getting all files...'
foreach ($Path in $FilePaths) {
    # get the child items
    $Items = Get-ChildItem -Path $Path
    # add the items to the list
    foreach ($Item in $Items) {
        $AllFiles.Add($Item)
    }
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
        # only delete files, skipping folders
        if (Test-Path -Path $File.FullName -PathType Leaf) {
            Remove-Item -Path $File.FullName -Force -ErrorAction SilentlyContinue
            $FileSizeCount += $File.Length
            $RemovedFileCount += 1
        }
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
