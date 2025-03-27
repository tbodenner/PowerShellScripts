#Requires -RunAsAdministrator

# properties

# the AD object we are going to search
$SearchBase = 'OU=Prescott (PRE),OU=VISN18,DC=v18,DC=med,DC=va,DC=gov'

# test if a computer can be pinged and check dns names
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
    if ($Null -eq $Ip) {
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

# get all our computers from our domain
function Get-Computers {
    # create our empty arrays
    $ADArray = @()
    # get our AD data
    $ADComputers = Get-ADComputer -Filter * -SearchBase $SearchBase | Select-Object Name
    # create a clean array from our AD data
    foreach ($Item in $ADComputers) {
        $ADArray += $Item.Name
    }
    # return our filled array
    return $ADArray
}

# change our default settings for our remote session used by invoke-command
$PssOptions = New-PSSessionOption -MaxConnectionRetryCount 0 -OpenTimeout 30000 -OperationTimeout 30000

# our commands to run on the remote computer
$ScriptBlock = {
    # get computer name
    $ComputerName = $env:COMPUTERNAME
    # get mac and ip addresses
    $NetworkData = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration | Where-Object IPAddress -ne $Null | Select-Object MACAddress,IPAddress,Description
    # get out user profile paths and last login date and time
    $ProfileData = Get-CimInstance -ClassName Win32_UserProfile | Where-Object Special -eq $False | Select-Object -Property LocalPath,LastUseTime
    # return our data using a hashtable
    New-Object -TypeName PSCustomObject -Property @{ ComputerName=$env:COMPUTERNAME; NetworkData=$NetworkData; ProfileData=$ProfileData }
}

# flush our dns
Clear-DnsClientCache

# loop through list of computers
foreach ($Computer in Get-Computers) {
    try {
        # get the last error in the error variable
        $LastError = $Error[0]
        
        Write-Host "Trying To Connect [$($Computer)]" -ForegroundColor Yellow
        # set our parameters for our invoke command
        $Parameters = @{
            ComputerName    = $Computer
            ScriptNlock        = $ScriptBlock
            ErrorAction        = "SilentlyContinue"
            SessionOption    = $PssOptions
        }
        # get our ping data
        $PingData = Computer-IsFound $Computer
        # get the result (boolean)
        $PingLatency = $PingData[0]
        # get the ping status (string)
        $PingStatus = $PingData[1]
        # check if we can ping the computer
        if ($PingLatency -gt 0) {
            # check if we did not have a dns mismatch
            if ($PingStatus -ne $DnsMismatch) {
                Write-Host "Ping ($($PingStatus))" -ForegroundColor Green
                # run the script on the target computer if we can ping the computer
                $InvokeResult = Invoke-Command @Parameters
                # check if our return object is null
                if ($Null -ne $InvokeResult) {
                    # TODO: parse return object and add to our output
                }
            }
            else {
                # otherwise, write an error
                Write-Error -Message "DNS mismatch error" -Category ConnectionError -ErrorAction SilentlyContinue
            }
        }
        else {
            Write-Host "Ping ($($PingStatus))" -ForegroundColor Red
            # otherwise, write an error
            Write-Error -Message "Unable to ping $($Computer)" -Category ConnectionError -ErrorAction SilentlyContinue
        }
        # determine if the last computer was a success or error
        if ($LastError -eq $Error[0]) {
            # if the script finished without adding an error, add the computer to our success array
            $SuccessArray += $Computer
        }
        else {
            Write-Host "Error: $($Computer)`n" -ForegroundColor Yellow
            # if the script added an error, add the computer to our success array
            $ErrorArray += $Computer
        }
    }
    catch {
        Write-Output "Caught Error"
        Write-Output "$($_)`n"
        Write-Host ($_ | Select-Object -Property *)
        # add the computer to our error array if an error was caught
        $ErrorArray += $Computer
    }
}