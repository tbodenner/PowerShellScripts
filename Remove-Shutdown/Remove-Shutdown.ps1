#Requires -RunAsAdministrator

# our custom line format for script output
function Format-Line {
	# parameters
	param (
	[Parameter(Mandatory)][string][string]$Text,
	[string]$Computer
	)
	# check if we were given a computer name
	if ($Computer -ne "") {
		# format the line with the our text and computer
		return "[--|$($Computer)| $($Text)"
	}
	else {
		# otherwise, format the line with the our text
		return "[--|$($Text)|"
	}
}

# write a colored line to the host
function Write-ColorLine {
	# parameters
	param (
		[Parameter(Mandatory)][string]$Text,
		[Parameter(Mandatory)][string]$Color
	)
	# check if our color is in the color enum
	if ([ConsoleColor]::IsDefined([ConsoleColor], $Color)) {	
		# color exists, write our colored line
		Write-Host $Text -ForegroundColor $Color
	}
	else {
		# color doesn't exist, write our line without color
		Write-Host $Text
	}
}

# test if a computer can be pinged
function Find-Computer {
	# parameters
	param (
		[Parameter(Mandatory)][string]$ComputerName
	)
	# ping the computer and save the details
	try {
		$ComputerDetails = Test-Connection -TargetName $ComputerName -Count 1 -TimeoutSeconds 3
	}
	# catch the ping exception
	catch [System.Net.NetworkInformation.PingException] {
		return @(0, $NotFound)
	}
	# catch all other errors
	catch {
		# write out the exception
		Write-Host ($_.Exception | Select-Object -Property *)
	}
	# get the pinged computer's ip
	$Ip = $ComputerDetails.Address
	# get the ping result
	$Latency = $ComputerDetails.Latency
	# get the status of the ping
	$Status = $ComputerDetails.Status
	# check if our status is null
	if ($Null -eq $Status) {
		# if null, set our status
		$Status = $NoStatus
	}
	# check if the ping timed out
	if (($Latency -eq 0) -or ($Status -eq $NoStatus)) {
		# the ping timed out, so return the result of our ping
		return @($Latency, $Status)
	}
	# set our error value for our dns name
	$DnsName = $Null
	# check if we got an ip
	if ($Null -ne $Ip) {
		# get our dns data
		try {
			$DnsData = Resolve-DnsName -Name $Ip
		}
		# catch the dns error
		catch [System.ComponentModel.Win32Exception] {
			# get our error code
			$ECode = $_.Exception.NativeErrorCode
			# return our error
			if ($ECode -eq 9003) {
				return @(0, 'DnsNotFound')
			}
		}
		# catch all other errors
		catch {
			# write out the exception
			Write-Host ($_.Exception | Select-Object -Property *)
		}
		# check if we got any data
		if ($Null -eq $DnsData) {
			return @($Latency, 'GenericDnsError')
		}
		# get our host name from the data
		$NameHost = $DnsData.NameHost
		# check if we got a hostname
		if ($Null -eq $NameHost) {
			return @($Latency, 'NoHostName')
		}
		# split the host name and return it
		$DnsName = $NameHost.Split('.')
	}
	# check if our resolved name is null
	if ($Null -eq $DnsName) {
		# if true, write a message
		Write-ColorLine -Text (Format-Line -Text "Unable to resolve computer name from IP address" -Computer $ComputerName) -Color Red
	}
	else {
		# otherwise, check if our computer name matches the dns name
		if ($DnsName[0] -ne $ComputerName) {
			Write-ColorLine -Text (Format-Line -Text "DNS mismatch (DNS: $($DnsName[0]), CN: $($ComputerName))" -Computer $ComputerName) -Color Red
			# return the dns error
			return @($Latency, $DnsMismatch)
		}
	}
	# return the result of our ping
	return @($Latency, $Status)
}

# prepare a file to be written to
function Initialize-ComputerListFile {
	# parameters
	param (
		[Parameter(Mandatory)][string]$FileName
	)
	# check if our computer error list file exists
	if ((Test-Path -Path $FileName) -eq $False) {
		# if not found, then create it
		New-Item -Path $FileName -ItemType "file" | Out-Null
	}
	else {
		# if found, then clear it
		Clear-Content -Path $FileName
	}
}

# stop on errors
$ErrorActionPreference = "Stop"

# if our status is empty, use this status
$NoStatus = 'NoStatus'
# if our ping fails to find a target, use this status
$NotFound = 'NotFound'
# if we have a dns mismatch, use this status
$DnsMismatch = 'DnsMismatch'

# log folder
$LogFolder = '.\Logs'

# check if a log folder exists
if ((Test-Path -Path $LogFolder) -eq $False) {
	# if not, create it
	New-Item -Path $LogFolder -ItemType Directory | Out-Null
}

# an array to collect the computer names with an errror
$ErrorArray = @()
$ErrorComputerFile = "$($LogFolder)\ErrorComputers.txt"
Initialize-ComputerListFile -FileName $ErrorComputerFile
# an array to collect the computer names that have the software already installed
$GoodArray = @()
$GoodComputerFile = "$($LogFolder)\GoodComputers.txt"
Initialize-ComputerListFile -FileName $GoodComputerFile
# an array to collect the computer names that need the software installed
$NeedArray = @()
$NeedComputerFile = "$($LogFolder)\NeedComputers.txt"
Initialize-ComputerListFile -FileName $NeedComputerFile

# file to store a list of computers to check. if not found, file is created and filled from an AD search
$ComputersFile = '.\ComputerList.txt'
# list of computers to check
$ComputerList = $Null

# check if the computer list file exists
if ((Test-Path -Path $ComputersFile) -eq $False) {
	# get the list of computers from a text file
	$SearchBase = 'DC=v18,DC=med,DC=va,DC=gov'
	$Filter = "Name -like 'PRE-MA*'"
	$ComputerList = Get-AdComputer -Filter $Filter -SearchBase $SearchBase | ForEach-Object { $_.Name }
	Out-File -FilePath $ComputersFile -InputObject $ComputerList
}
else {
	# get the computers from out computer file
	$ComputerList = Get-Content -Path $ComputersFile
	# check if our list has any items
	if ($ComputerList.Count -le 0) {
		# if the list is empty, then exit
		Write-ColorLine -Text "Computer list is empty. Exiting." -Color Magenta
		return
	}
}

# check if our computer list is null
if ($Null -eq $ComputerList) {
	# if the list is null, then exit
	Write-ColorLine -Text "Computer list is null. Exiting." -Color Red
	return
}

# write our starting status
$Plural = ""
if ($ComputerList.Count -eq 1) {
	$Plural = "computer"
}
else {
	$Plural = "computers"
}
$StartString = "`nChanging shutdown options on $($ComputerList.Count) $($Plural)`n"
Write-ColorLine -Text $StartString -Color Green

# clear the error list so we can write only our errors
$Error.Clear()

# count variables
$TotalComputers = $ComputerList.Count
$ComputerCount = 0

# percent complete
$PComplete = 0.0

# flush our dns
Clear-DnsClientCache

# change our default settings for our remote session used by invoke-command
$PssOptions = New-PSSessionOption -MaxConnectionRetryCount 0 -OpenTimeout 30000 -OperationTimeout 30000 -CancelTimeout 10000 -IdleTimeout 60000

# define the script we will run on each computer
$ShutdownScriptBlock = {
	# try to change the registry values
	try {
		# set our value to be changed. disable: 0x1, enable: 0x0
		$value = 0x1

		# change the registry value for shutdown
		Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\PolicyManager\default\Start\HideShutDown" -Name "value" -Value $value
		Write-Host "[--|$($env:computername)| Updated shutdown value"
		# change the registry value for sleep
		Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\PolicyManager\default\Start\HideSleep" -Name "value" -Value $value
		Write-Host "[--|$($env:computername)| Updated sleep value"

		# get shutdown value
		$ShutdownValue = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\PolicyManager\default\Start\HideShutDown" -Name "value").value
		# get sleep value
		$SleepValue = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\PolicyManager\default\Start\HideSleep" -Name "value").value

		# check if both values have been set
		if (($ShutdownValue -eq $value) -and ($SleepValue -eq $value)) {
			Write-Host "[--|$($env:computername)| Shutdown and sleep values verified ($($ShutdownValue), $($SleepValue))"
			# return true
			$True
		}
		else {
			Write-Host "[--|$($env:computername)| Failed to set shutdown and sleep values ($($ShutdownValue), $($SleepValue))"
			# return false
			$False	
		}
	}
	catch {
		# write a message if we get an error
		Write-Host "[--|$($env:computername)| Error while trying to update registry"
		# return false
		$False
	}
}

# loop through list of computers
foreach ($Computer in $ComputerList) {
	try {
		# update our progress
		$PComplete = ($ComputerCount / $TotalComputers) * 100
		$Status = "$ComputerCount/$TotalComputers Complete"
		$Activity = Format-Line -Text "Progress   "
		Write-Progress -Activity $Activity -Status $Status -PercentComplete $PComplete
		Write-ColorLine -Text (Format-Line -Text "Trying To Connect" -Computer $Computer) -Color Yellow
		# set our parameters for our invoke command
		$Parameters = @{
			ComputerName	= $Computer
			ScriptBlock		= $ShutdownScriptBlock
			ErrorAction		= "SilentlyContinue"
			SessionOption	= $PssOptions
		}
		
		# get our ping data
		$PingData = Find-Computer $Computer
		# get the result (boolean)
		$PingLatency = $PingData[0]
		# get the ping status (string)
		$PingStatus = $PingData[1]
		# check if we can ping the computer
		if ($PingLatency -gt 0) {
			# check if we did not have a dns mismatch
			if ($PingStatus -ne $DnsMismatch) {
				Write-ColorLine -Text (Format-Line -Text "Ping ($($PingStatus))" -Computer $Computer) -Color Green
				# run the script on the target computer if we can ping the computer
				$ShutdownUpdated = Invoke-Command @Parameters
				# check if we got anything back from the invoke command
				if ($Null -eq $ShutdownUpdated) {
					Write-ColorLine -Text (Format-Line -Text 'Result was NULL' -Computer $Computer) -Color Magenta
					# update our array
					$ErrorArray += $Computer
					Add-Content -Path $ErrorComputerFile -Value $Computer
				}
				else {
					# check if the registry update was made
					if ($ShutdownUpdated) {
						Write-ColorLine -Text (Format-Line -Text 'Updated registry settings' -Computer $Computer) -Color Blue
						# update our array
						$GoodArray += $Computer
						Add-Content -Path $GoodComputerFile -Value $Computer
					}
					else {
						Write-ColorLine -Text (Format-Line -Text 'Failed to update registry settings' -Computer $Computer) -Color Cyan
						# update our array
						$NeedArray += $Computer
						Add-Content -Path $NeedComputerFile -Value $Computer
					}
				}
			}
			else {
				# otherwise, write an error
				Write-Error -Message "DNS mismatch error" -Category ConnectionError -ErrorAction SilentlyContinue
				# update our array
				$ErrorArray += $Computer
				Add-Content -Path $ErrorComputerFile -Value $Computer
			}
		}
		else {
			Write-ColorLine -Text (Format-Line -Text "Ping ($($PingStatus))" -Computer $Computer) -Color Red
			# otherwise, write an error
			Write-Error -Message "Unable to ping $($Computer)" -Category ConnectionError -ErrorAction SilentlyContinue
			# update our array
			$ErrorArray += $Computer
			Add-Content -Path $ErrorComputerFile -Value $Computer
		}
		# increment our count
		$ComputerCount += 1
	}
	catch {
		Write-ColorLine -Text (Format-Line -Text 'Caught Error' -Computer $Computer) -Color Red
		Write-Host "$($_)`n"
		Write-Host ($_ | Select-Object -Property *)
		# add the computer to our error array if an error was caught
		$ErrorArray += $Computer
		Add-Content -Path $ErrorComputerFile -Value $Computer
	}
	finally {
		# write an empty line to seperate the output between computers
		Write-Host
	}
}

# create our counts array to output to the console and a file
$CountsArray = @(
	'Results:'
	"Total: $($TotalComputers)"
	" Good: $($GoodArray.Count)"
	" Need: $($NeedArray.Count)"
	"Error: $($ErrorArray.Count)"
)

# write our counts
Write-Host $CountsArray[0]
Write-ColorLine -Text $CountsArray[1] -Color Yellow
Write-ColorLine -Text $CountsArray[2] -Color Blue
Write-ColorLine -Text $CountsArray[3] -Color Cyan
Write-ColorLine -Text $CountsArray[4] -Color Red

# write our output files
Write-Host 'Writing Output Files'
# output computer names that had the software installed
Out-File -FilePath $GoodComputerFile -InputObject $GoodArray
# output computer names that need the software installed
Out-File -FilePath $NeedComputerFile -InputObject $NeedArray
# output computer names that had an error during the script
Out-File -FilePath $ErrorComputerFile -InputObject $ErrorArray

# write our error computers to our computerlist for the next run
Write-Host "Computer error list written to computer list"
Out-File -FilePath $ComputersFile -InputObject $ErrorArray

# output our errors encountered during the script
Write-Host "Errors Written to Errors.txt"
Out-File -FilePath "$($LogFolder)\Errors.txt" -InputObject $Error

# write log file
$DateString = Get-Date -Format "MM.dd.yyyy-HH.mm.ss"
$ShutdownLogFile = "$($LogFolder)\$($DateString)-Shutdown.log"
Out-File -FilePath $ShutdownLogFile -InputObject $CountsArray
Write-Host "Results written to '$($ShutdownLogFile)'`n"
