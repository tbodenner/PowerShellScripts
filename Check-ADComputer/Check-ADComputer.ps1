# the AD object we are going to search
$SearchBase = 'OU=Prescott (PRE),OU=VISN18,DC=v18,DC=med,DC=va,DC=gov'
# get our AD data
$ADComputers = Get-ADComputer -Filter * -SearchBase $SearchBase | Select-Object Name
# get our computer list from our file
$ComputerList = Get-Content -Path .\ComputerList.txt
# create our empty arrays
$ADArray = @()
$NotInAD = @()
$InAD = @()
# create a clean array from our AD data
foreach ($Item in $ADComputers) {
	$ADArray += $Item.Name
}
# check if the computer name in our input array is in our AD array
foreach ($Computer in $ComputerList) {
	# make the name uppercase to match AD data
	$Computer = $Computer.ToUpper()
	# check if the computer name in the AD data
	if ($ADArray -notcontains $Computer) {
		# if not found, update our not in AD array
		$NotInAD += $Computer
	}
	else {
		# if found, update our in AD array
		$InAD += $Computer
	}
}
# sort our data then output our arrys
Out-File -FilePath '.\NotInAD.txt' -InputObject ($NotInAD | Sort-Object)
Out-File -FilePath '.\InAD.txt' -InputObject ($InAD | Sort-Object)
