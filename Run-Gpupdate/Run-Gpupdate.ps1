#Requires -RunAsAdministrator

# script to run on remote computer
$ScriptBlock = {
    # our temp folder
    $TempFolder = 'C:\Temp'

    # check if temp folder exists
    if ((Test-Path -Path $TempFolder) -eq $false) {
        # if not found, create it
        New-Item -ItemType Directory -Path $TempFolder
    }

    # timestamp file name
    $TimeStampFile = Join-Path -Path $TempFolder -ChildPath 'gpupdate-timestamp.txt'

    # timestamp read from file
    $TimeStamp = $null

    # our boolean for running gpupdate
    $RunGpupdate = $false

    # check if we have a timestamp file
    if ((Test-Path -Path $TimeStampFile) -eq $true) {
        # read the timestamp from the file
        $FileData = Get-Content -Path $TimeStampFile
        # convert the timestamp to a datetime object
        $TimeStamp = [DateTime]::Parse($FileData.Trim())
        # get a datetime 1 day before now
        $OneDayAgo = (Get-Date).AddDays(-1)
        # check if our timestamp is before our one day datetime
        if ($TimeStamp -lt $OneDayAgo) {
            # if true, run gpupdate
            $RunGpupdate = $true
        }    
    }
    else {
        # no timestamp file, run gpupdate
        $RunGpupdate = $true
    }

    # run gpupdate if our boolean is true
    if ($RunGpupdate -eq $true) {
        # run gpupdate
        gpupdate.exe /force /wait:120 | Out-Null

        # write a timestamp to a file in our temp folder
        Get-Date | Out-File -FilePath $TimeStampFile

        # get logged in user
        $User = (Get-CimInstance -ClassName Win32_ComputerSystem).Username

        # if no user is logged in
        if ($null -eq $User) {
            Write-Host "$($env:COMPUTERNAME): Restarting"
            # restart the computer
            shutdown.exe /f /r /t 60 | Out-Null
        }
        else {
            # user logged in
            Write-Host "$($env:COMPUTERNAME): $($User)"
        }
    }
    else {
        # gpupdate was not run
        Write-Host "$($env:COMPUTERNAME): Skipped"
    }
}

# computer list
$Computers = Get-Content '.\ComputerList.txt'

# run the script block for each computer as a job
foreach ($Computer in $Computers) {
    Invoke-Command -ComputerName $Computer -ScriptBlock $ScriptBlock -AsJob | Out-Null
}

# store our jobs
$AllJobs = 'JOBS'
# continue checking for new jobs until none are found
while ($Null -ne $AllJobs) {
	# get all the current jobs
	$AllJobs = Get-Job
	# get each job's status
	foreach ($Job in $AllJobs) {
        # get the computer name from the job
		$Computer = $Job.Location
		# if the job failed, add the computer to our error list and remove the job
        switch ($Job.State) {
            'Failed' {
                Write-Host "$($Computer): Failed"
			    Remove-Job -Job $Job -ErrorAction SilentlyContinue
		    }
		    {$_ -in ('Completed','Stopped')} {
                # get the job data
                $CommandResult = Receive-Job -Job $Job
                # if we have data
                if ($null -ne $CommandResult) {
                    # write the result
                    Write-Host $CommandResult
                }
                # remove the job
                Remove-Job -Job $Job -ErrorAction SilentlyContinue
            }
            {$_ -in ('Blocked', 'Suspended', 'Disconnected')} {
                # stop the job
                Stop-Job -Job $Job -ErrorAction SilentlyContinue
            }
            default { continue }
        }
	}
}

# done
Write-Host "Done."
