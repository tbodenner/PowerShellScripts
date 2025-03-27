#Requires -RunAsAdministrator

$CsvFile = '.\UserGroups.csv'
$UserGroupHashtable = @{}
$MaxGroupCount = 0
$ADGroupName = 'OIT ITOPS SO EUO PAC T1 PRE IT Staff'

# get all users from an AD group
$ADUsers = Get-ADGroupMember -Identity $ADGroupName

# get user groups data from active directory
foreach ($ADUser in $ADUsers) {
    # get the username from the AD object
    $User = $ADUser.SamAccountName
    # skip the line if the user is null or empty
    if (($Null -eq $User) -or ($User -eq "")) {
        continue
    }
    # get all our groups for this user
    $FullGroupName = (Get-AdUser $User -Properties *).MemberOf
    # if the user has no groups, write the username to host and move onto the next user
    if ($Null -eq $FullGroupName) {
        Write-Host $User -ForegroundColor Red
        continue
    }
    # otherwise, write the username to host
    else {
        Write-Host $User -ForegroundColor Green
    }
    # our array to add our groups to
    $GroupArray = @()
    # loop through our groups
    foreach ($Group in $FullGroupName) {
        # split our full group name and only keep the CN
        $GroupName = $Group.Split(',')[0].Replace("CN=","")
        # if we got a group name, add it to our array
        if (($Null -ne $GroupName) -and ($GroupName -ne "")) {
            $GroupArray += $GroupName
        }
    }
    # check if we added any groups to our array
    if ($GroupArray.Count -gt 0) {
        # check if this group count is the largest group count
        if ($GroupArray.Count -gt $MaxGroupCount) {
            # update our new largest group count
            $MaxGroupCount = $GroupArray.Count
        }
    }
    # add our sorted group array to our hashtable for the user
    $UserGroupHashtable[$User.ToLower()] = $GroupArray | Sort-Object
    # write an empty line to seperate the host output as we move to a new user
}

# define our array with a specific size
$CsvLineArray = [string[]]::new($MaxGroupCount)
# loop through our users in order
foreach ($u in ($UserGroupHashtable.Keys | Sort-Object)) {
    # add the username for the header
    $CsvLineArray[0] += "`"$($u)`","
    # starting on line one, write a group for each user on each line
    for ($l = 1; $l -lt $MaxGroupCount; $l++) {
        # our arrays start at zero
        $g = $l - 1
        # check if the array contains the index number
        if ($g -lt $UserGroupHashtable[$u].Count) {
            # if the index exists, add it to our line
            $CsvLineArray[$l] += "`"$($UserGroupHashtable[$u][$g])`","
        }
        else {
            # otherwise, add an empty value
            $CsvLineArray[$l] += "`"`","
        }
    }
}
# write our csv data to a file
Out-File -FilePath $CsvFile -InputObject $CsvLineArray
