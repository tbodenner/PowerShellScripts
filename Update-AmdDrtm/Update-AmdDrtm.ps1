#Requires -RunAsAdministrator

$InstallDriverScriptBlock = {
    function Get-IntFromVersion {
        param ([string]$Version)
        # change the string into an int
        $VersionInt = [int]$Version.Replace('.', '')
        # return the int
        $VersionInt
    }

    function Get-DriverVersion {
        # get all the paths for the amddrtm drivers installed on this computer
        $Drivers = (Resolve-Path -Path 'C:\Windows\System32\DriverStore\FileRepository\amddrtm.inf*\amddrtm.inf').Path
        # our driver version to return
        $OutputDriver = -1
        # for each driver path
        foreach ($Driver in $Drivers) {
            # get the version of the driver
            $DriverVersion = Get-IntFromVersion -Version (Get-WindowsDriver -Online -Driver $Driver).Version
            # if the driver version is greater than our output driver, update our output driver
            if ($DriverVersion -gt $OutputDriver) { $OutputDriver = $DriverVersion }
        }
        # return our output driver
        $OutputDriver
    }

    function Remove-DriverFiles {
        param (
            [string]$Folder,
            [string]$File
        )
        # delete the files and folders
        Remove-Item -Path $Folder -Recurse -Force
        Remove-Item -Path $File -Force
    }

    # try to install our driver
    try {
        # file and folder names
        $TempFolder = 'Temp'
        $UpdateFolder = 'AMD-DRTM'
        $UpdateArchiveFile = 'amddrtm.zip'
        $UpdateFile = 'amddrtm.inf'

        # paths
        $TempPath = Join-Path -Path $env:SystemDrive -ChildPath $TempFolder         # C:\Temp
        $UpdateFolderPath = Join-Path -Path $TempPath -ChildPath $UpdateFolder      # C:\Temp\AMD-DRTM
        $UpdateArchiveFilePath = Join-Path $TempPath -ChildPath $UpdateArchiveFile  # C:\Temp\amddrtm.zip
        $UpdateFilePath = Join-Path -Path $UpdateFolderPath -ChildPath $UpdateFile  # C:\Temp\AMD-DRTM\amddrtm.inf

        # our target version
        $TargetVersion = Get-IntFromVersion -Version '1.0.18.4'

        # current driver version
        $DriverVersion = Get-IntFromVersion -Version (Get-DriverVersion)

        # check for our error value
        if ($DriverVersion -le -1) {
            # remove the driver files
            Remove-DriverFiles -Folder $UpdateFolderPath -File $UpdateArchiveFilePath
            # the driver was not found on this computer
            return 'Driver Not Found'
        }

        # if less than target 1.0.18.4
        if ($DriverVersion -lt $TargetVersion) {
            # decompress the driver
            Expand-Archive -Path $UpdateArchiveFilePath -DestinationPath $UpdateFolderPath -Force
            # install driver
            Start-Process -FilePath 'pnputil.exe' -ArgumentList "/add-driver $($UpdateFilePath) /install" -NoNewWindow -Wait | Out-Null
            # remove the driver files
            Remove-DriverFiles -Folder $UpdateFolderPath -File $UpdateArchiveFilePath
            # check if new driver has been installed
            if((Get-IntFromVersion -Version (Get-DriverVersion)) -ge $TargetVersion) {
                # driver updated
                return 'Updated'
            }
            else {
                # update failed
                return 'Fail'
            }
        }
        else {
            # driver already updated
            return 'Good'
        }
    }
    catch {
        # error
        return 'Error'
    }
}

function Update-ComputerArray {
    param (
        [string]$Name,
        [string[]]$Array
    )
    # create a new array excluding the name
    $NewArray = $Array | Where-Object { $_ -ne $Name }
    # return the new array
    $NewArray
}

function Get-HostFromDns {
    param ([string]$Ip)
    # get our dns data from our ip
    $DnsResult = (Resolve-DnsName -Name $Ip -ErrorAction SilentlyContinue)
    # if out result is null, return null
    if ($null -eq $DnsResult) { return $null }
    # if our result has no host, return null
    if ($null -eq $DnsResult.NameHost) { return $null }
    # get the computer name for this ip and return it
    return $DnsResult.NameHost.Split('.')[0]
}

# file for our list of computers
$ComputerFile = '.\ComputerList.txt'

# if the computer list file is missing, create it
if ((Test-Path -Path $ComputerFile) -eq $false) {
    New-Item -ItemType File $ComputerFile | Out-Null
}

# get list of computers
$Computers = Get-Content -Path $ComputerFile

# file to copy to the computer that contains our drivers
$UpdateArchiveFile = 'amddrtm.zip'

# get AD computers
$ADComputers = (Get-ADComputer -Filter 'Name -like "PRE-LT*"' | Where-Object { $_.Enabled -eq $true }).Name

# clone our array if it has any items in it
if ($Computers.Length -gt 0) {
    $OutputComputers = $Computers.Clone()
}
else {
    # otherwise, fill our computer array
    $Computers = $ADComputers.Clone()
    # and create an empty array for our output
    $OutputComputers = @()
}

# foreach computer
foreach ($Computer in $Computers) {
    # skip any null or empty computers
    if (($null -eq $Computer) -or ($Computer -eq '')) { continue }
    # the current computer we are working on
    Write-Host "$($Computer): " -NoNewline
    # check if the computer is not in AD
    if ($Computer -notin $ADComputers) {
        # remove the good result from our output array
        $OutputComputers = Update-ComputerArray -Name $Computer -Array $OutputComputers
        # write our output
        Write-Host 'Not in AD' -ForegroundColor Red
        # move to the next computer
        continue
    }
    # ping the computer
    if ((Test-Connection -TargetName $Computer -Ping -Count 1 -TimeoutSeconds 1 -Quiet) -eq $True) {
        # try to get dns data
        try {
            # get our computer's name from it's dns ip address
            $IpAddress = (Resolve-DnsName -Name $Computer -ErrorAction SilentlyContinue).IPAddress
            # check if our ip address is null
            if ($null -eq $IpAddress) {
                Write-Host 'IP Not Found' -ForegroundColor Red
                # move to the next computer
                continue
            }
            # check if we got multiple ips
            if ($IpAddress.Count -gt 1) {
                # check each ip for a computer name
                foreach ($Ip in $IpAddress) {
                    # get the computer name for this ip
                    $ComputerDns = Get-HostFromDns -Ip $Ip
                    # no computer name was returned for the ip
                    if ($null -eq $ComputerDns) {
                        # this ip did not return a name
                        continue
                    }
                    # if the computer name matches our computer, stop the loop
                    if ($ComputerDns.ToLower() -eq $Computer.ToLower()) { break }
                }
            }
            else {
                # otherwise, get the computer name for the ip
                $ComputerDns =  Get-HostFromDns -Ip $IpAddress
            }
        }
        catch  {
            # write our error
            Write-Host 'DNS Error' -ForegroundColor Red
            Write-Host ($_ | Out-String)
            # move to the next computer
            continue
        }

        # try to install the driver on the remote computer
        try {
            # check if our dns name is null
            if ($null -eq $ComputerDns) {
                # write the result and move to the next computer
                Write-Host 'DNS Not Found' -ForegroundColor Red
                continue
            }
            # check if the name from dns matches our name
            if ($ComputerDns.ToLower() -eq $Computer.ToLower()) {
                # try to copy the driver to the computer
                try {
                    # copy the file to the computer
                    Copy-Item -Path ".\$($UpdateArchiveFile)" -Destination "\\$($Computer)\c$\Temp\" -Force -ErrorAction SilentlyContinue
                }
                catch {
                    # if the copy failed, move to the next computer
                    Write-Host 'Copy Error' -ForegroundColor Red
                    continue
                }
                # change our default settings for our remote session used by invoke-command
                $PssOptions = New-PSSessionOption -MaxConnectionRetryCount 0 -OpenTimeout 30000 -OperationTimeout 30000
                # invoke command options
                $Parameters = @{
                    ComputerName	= $Computer
                    ScriptBlock		= $InstallDriverScriptBlock
                    ErrorAction		= "SilentlyContinue"
                    SessionOption	= $PssOptions
                }
                # invoke command to run the install script block
                $InvokeResult = Invoke-Command @Parameters
                # check if our result was null
                if ($null -eq $InvokeResult) {
                    # our command returned no results, move onto the next computer
                    Write-Host 'Command Result Null' -ForegroundColor Red
                    continue
                }
                # check our result
                if ($InvokeResult.ToLower() -in @('good','updated','driver not found')) {
                    # remove the good result from our output array
                    $OutputComputers = Update-ComputerArray -Name $Computer -Array $OutputComputers
                }
                # write our output
                Write-Host $InvokeResult -ForegroundColor Green
            }
            else {
                Write-Host 'DNS Mismatch' -ForegroundColor Red
            }
        }
        catch {
            Write-Host 'Command/Copy Error' -ForegroundColor Red
            Write-Host ($_ | Out-String)
        }
}
    else {
        Write-Host 'Offline'
    }
}

# try to write our array to file
try {
    # write our update computer list
    $OutputComputers | Out-File -FilePath $ComputerFile
    # file was written
    Write-Host "Wrote computer list '$($ComputerFile)'"
}
catch {
    # any errors
    Write-Host "Unable to write file '$($ComputerFile)'"
}
