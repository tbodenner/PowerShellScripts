param (
        [switch]$ShowAll,
        [switch]$Offline
    )

function Get-UserCert {
    # get our AD user data
    $UserAdData = Get-ADUser -Identity $env:username -Properties *
    # get full name
    $FullName = "$($UserAdData.GivenName) $($UserAdData.SN)"
    # get certs
    $AllUserCerts = Get-ChildItem -Path "Cert:\CurrentUser\My"
    # our cert thumbprint
    $ReturnCert = $Null
    # loop through our certs looking for our user cert
    foreach ($Cert in $AllUserCerts) {
        # check if the cert suject contains our full name
        if ($Cert.Subject -like "*$($FullName.ToUpper())*") {
            # check if this cert is using for loggin in
            if ($Cert.EnhancedKeyUsageList.FriendlyName -contains 'Smart Card Logon') {
                # set our return cert
                $ReturnCert = $Cert
            }
        }
    }
    # return our last found cert
    $ReturnCert
}

function Get-RecordIds {
    # get a cert for our request
    $UserCert = Get-UserCert
    # build our query string (query string was found on carf leaf site by clicking json button on carf list page)
    $Query = '{"terms":[{"id":"categoryID","operator":"=","match":"form_7d7fc"},'
    $Query += '{"id":"stepID","operator":"!=","match":"resolved"},'
    $Query += '{"id":"stepID","operator":"=","match":"35"},'
    $Query += '{"id":"deleted","operator":"=","match":0,"gate":"AND"}],'
    $Query += '"joins":["service","status","initiatorName"],'
    $Query += '"sort":{},"getData":["53","57","59"]}'
    # the uri for our leaf api
    $LeafURI = "https://leaf.va.gov/VISN18/649/IT_Requests/api/form/query/?q=$($Query)"
    $LeafURI += '&x-filterData=recordID' #,title,service,stepTitle,lastStatus,lastName,firstName'
    # get our leaf json data
    $LeafContent = (Invoke-WebRequest -Certificate $UserCert -URI $LeafURI).Content
    # write our file
    #$LeafContent | Out-File '.\id.json'
    # get our record ids from our request
    $LeafRecordIds = ($LeafContent | ConvertFrom-Json).psobject.properties.name
    # return our records
    $LeafRecordIds
}

# check if our data folder exists, if not found, create it
function Test-DataFolder {
    # name of our folder
    $LeafDataFolder = '.\LeafData'
    # check if the folder exists
    if ((Test-Path -Path $LeafDataFolder) -eq $False) {
        # if not, create it
        New-Item -Path $LeafDataFolder -ItemType Directory
    }
}

# clean html from our string and split it on double spaces
function Get-CleanString {
    param (
        [Parameter(Mandatory=$False)][string]$InputString
    )
    # replace any html in our string with double spaces
    $NoHtmlString = $InputString -creplace "<[^>]*>", "  "
    # split our string on double spaces
    $StringArray = $NoHtmlString -split "\s{2}"
    # create our output array
    $OutputArray = @()
    # loop through our string array
    foreach ($Str in $StringArray) {
        # trim our string before checking it
        $Str = $Str.Trim()
        # if the string is not null or empty
        if (($Null -ne $Str) -and ($Str -ne "")) {
            # add the string to our output
            $OutputArray += $Str
        }
    }
    # return our string areray
    $OutputArray
}

# remove extra characters from the keys and output a comma seperated string
function Get-CleanPrimaryMenu {
    param (
        [Parameter(Mandatory=$False)][string]$InputString
    )
    # clean up our string
    $StringArray = Get-CleanString -InputString $InputString
    # check if we have any items in our array
    if ($StringArray.Count -ge 1) {
        # output the first item
        $StringArray[0]
    }
    else {
        # otherwise, return null
        $Null
    }
}

# remove extra characters from the keys and output a comma seperated string
function Get-CleanSecondaryMenuOrKeys {
    param (
        [Parameter(Mandatory=$False)][string]$InputString
    )
    # clean up our string
    $StringArray = Get-CleanString -InputString $InputString
    # our array we are going to build
    $OutputArray = @()
    # loop through our split array
    foreach ($Str in $StringArray) {
        # trim both ends of our string of any whitespace
        $Str = $Str.Trim()
        # if the string is null or empty, skip it
        if (($Null -eq $Str) -or ($Str -eq "")) { continue }
        # ignore any strings with a lowercase character
        if ($Str -match "(?-i)[a-z]") { continue }
        # look for any strings with all uppercase characters
        if ($Str -match "(?-i)[A-Z0-9\s]+") {
        #if ($Str -match "\[(.*?)\]") {
            # the second match has our string without the square brackets, add it to our array
            #Write-Host "MATCHES: $($matches | Out-String)"
            $OutputArray += $matches[0]
        }
    }
    # output our array as a comma seperated string
    $OutputArray -join ", "
}

# remove extra characters from the keys and output a comma seperated string
function Get-CleanKeys {
    param (
        [Parameter(Mandatory=$False)][string]$InputString
    )
    # clean up our string
    $StringArray = Get-CleanString -InputString $InputString
    # our array we are going to build
    $OutputArray = @()
    # loop through our split array
    foreach ($Str in $StringArray) {
        # trim both ends of our string of any whitespace
        $Str = $Str.Trim()
        # if the string is null or empty, skip it
        if (($Null -eq $Str) -or ($Str -eq "")) { continue }
        # add our string to our array
        $OutputArray += $Str
    }
    # output our array as a comma seperated string
    $OutputArray -join ", "
}

# get our requestor's full name from AD
function Get-RequestorFullName {
    param (
        [Parameter(Mandatory=$False)][string]$InputName
    )
    if ($Null -eq $InputName) {
        $Null
    }
    else {
        (Get-AdUser $InputName -Properties *).Name
    }
}

# try to get an AD name using our Vista name
function Get-NameFromVistaName {
    param (
        [Parameter(Mandatory=$False)][string]$VistaName
    )
    # replace commas with spaces
    $CleanName = $VistaName.Trim() -replace ",", " "
    # split the string by spaces
    $NameArray = $CleanName -split " "
    # create a new array to store a cleaned up array
    $CleanNameArray = @()
    # remove any empty lines from the name array
    foreach ($Item in $NameArray) {
        if (($Null -ne $Item) -and ($Item -ne "")) {
            $CleanNameArray += $Item
        }
    }
    # we need two parts to filter on
    if ($CleanNameArray.Count -ge 2)
    {
        # return the name from AD using our name parts as a filter
        (Get-AdUser -Filter "Name -like '*$($CleanNameArray[0])*' -and Name -like '*$($CleanNameArray[1])*'" -Properties *).Name
    }
    else {
        # otherwise, return null
        $Null
    }
}

# get the leaf id from our record filenames
function Get-OfflineLeafRecordIds {
    # the folder our files are saved in
    $LeafFolder = '.\LeafData'
    # get a list of all our files
    $LeafFiles = Get-ChildItem -Path $LeafFolder
    # the array we will return
    $FileNameArray = @()
    # get the id from each file
    foreach ($File in $LeafFiles) {
        # add the short file name to our array
        $FileNameArray += $File.Name
    }
    # return our array
    $FileNameArray
}

# check if our data folder exists
Test-DataFolder

# determine if we are using online or offline records
if ($Offline -eq $True) {
    # get our records from our saved records
    $LeafRecordIds = Get-OfflineLeafRecordIds
}
else {
    # get our record ids from out online source
    $LeafRecordIds = Get-RecordIds
}

# get the leaf json data for each record id or file
foreach ($LeafId in $LeafRecordIds) {
    # determine if we are using online or offline records
    if ($Offline -eq $True) {
        # create our file path from our file name
        $LeafFile = Join-Path -Path '.\LeafData\' -ChildPath $LeafId
        # get the json data from our file
        $LeafJsonData = Get-Content -Path $LeafFile | ConvertFrom-Json
    }
    else {
        # leaf file name
        $LeafFile = ".\LeafData\leaf$($LeafId).json"

        # get a cert for our request
        $UserCert = Get-UserCert
        # the uri for our leaf api (was found by viewing page source on multiple a carf pages)
        $LeafURI = "https://leaf.va.gov/VISN18/649/IT_Requests/api/form/$($LeafId)/data/tree"
        # get our leaf json data
        $LeafContent = (Invoke-WebRequest -Certificate $UserCert -URI $LeafURI).Content
        # write our leaf file
        $LeafContent | Out-File $LeafFile

        # get the json data from our request
        $LeafJsonData = $LeafContent | ConvertFrom-Json
    }

    # check if we got any data from the request
    if ($Null -ne $LeafJsonData) {
        # get the values from our leaf json data
        $HashData = [ordered]@{
            LeafID = $LeafId
            Requestor = $LeafJsonData[0].userID
            RequestorFullName = ""
            UserVistaName = $LeafJsonData[0].child."53".value
            UserFullName = ""
            UserTitle = $LeafJsonData[0].child."53".child."54".value
            UserGender = $LeafJsonData[0].child."56".value
            UserDate = $LeafJsonData[0].child."57".value
            UserPhone = $LeafJsonData[0].child."58".value
            UserType = $LeafJsonData[0].child."59".value
            UserTermDate = $LeafJsonData[0].child."59".child."60".value
            Justification = $LeafJsonData[0].child."62".value.Replace("<br />", " ").Replace("  ", " ")
            UserDisableDate = $LeafJsonData[0].child."63".value
            UserSS = $LeafJsonData[0].child."63".child."327".value
            UserBDay = $LeafJsonData[0].child."63".child."327".child."328".value
            DisableAccount = $LeafJsonData[0].child."210".value
            UserService = $LeafJsonData[0].child."330".value
            PrimaryMenuAdd = Get-CleanPrimaryMenu -InputString $LeafJsonData[1].child."65".child."66".value
            #PrimaryMenuAddRaw = $LeafJsonData[1].child."65".child."66".value
            PrimaryMenuRemove = Get-CleanPrimaryMenu -InputString $LeafJsonData[1].child."65".child."67".value
            #PrimaryMenuRemoveRaw = $LeafJsonData[1].child."65".child."67".value
            SecondaryMenuAdd = Get-CleanSecondaryMenuOrKeys -InputString $LeafJsonData[1].child."68".child."69".value
            #SecondaryMenuAddRaw = $LeafJsonData[1].child."68".child."69".value
            SecondaryMenuRemove = Get-CleanSecondaryMenuOrKeys -InputString $LeafJsonData[1].child."68".child."70".value
            #SecondaryMenuRemoveRaw = $LeafJsonData[1].child."68".child."70".value
            KeysAdd = Get-CleanSecondaryMenuOrKeys -InputString $LeafJsonData[1].child."71".child."72".value
            KeysRemove = Get-CleanSecondaryMenuOrKeys -InputString $LeafJsonData[1].child."71".child."73".value
            VistaTest = $LeafJsonData[1].child."74".child."75".value
            CprsTabs = $LeafJsonData[1].child."295".value
        }

        # sometimes the first userID is null, check each spot it can be until not null
        if ($Null -eq $HashData["Requestor"]) {
            if ($Null -ne $LeafJsonData[0].child."210".userID) {
                $HashData["Requestor"] = $LeafJsonData[0].child."210".userID
            }
            elseif ($Null -ne $LeafJsonData[0].child."53".value) {
                $HashData["Requestor"] = $LeafJsonData[0].child."53".value
            }
            elseif ($Null -ne $LeafJsonData[0].child."53".child."54".value) {
                $HashData["Requestor"] = $LeafJsonData[0].child."53".child."54".value
            }
            else {
                # unable to find requestor, write an error
                Write-Host "Error: Unable to find Requestor ID!" -ForegroundColor Red
            }
        }

        # get requestors full name from AD
        if (($Null -ne $HashData["Requestor"]) -and ($HashData["Requestor"] -ne "")) {
            $HashData["RequestorFullName"] = Get-RequestorFullName -InputName $HashData["Requestor"]
        }

        # try to get our AD name from the input/Vista name
        $HashData["UserFullName"] = Get-NameFromVistaName -VistaName $HashData["UserVistaName"]

        # write the begin block
        Write-Host '----- BEGIN -----' -ForegroundColor White
        # write our values to the console
        foreach ($Key in $HashData.Keys) {
            $Val = $HashData[$Key]
            if (($ShowAll -eq $False) -and (($Key -eq "UserSS") -or ($Key -eq "UserBDay"))) {
                Write-Host "$($Key): " -NoNewline -ForegroundColor Green
                Write-Host "--Hidden--"
            }
            else {
                Write-Host "$($Key): " -NoNewline -ForegroundColor Green
                Write-Host $Val
            }
        }

        # check if the requestor is the same as the user
        if ($HashData["UserFullName"] -eq $HashData["RequestorFullName"]) {
            Write-Host 'ERROR: User has requested their own menus and keys!' -ForegroundColor Red
        }

        # write the end block
        Write-Host '-----  END  -----' -ForegroundColor White
        # write an empty line
        Write-Host
    }
}
