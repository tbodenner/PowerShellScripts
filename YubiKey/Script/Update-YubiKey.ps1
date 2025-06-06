# yubikey manager cli
$YubiCli = 'C:\Program Files\Yubico\YubiKey Manager CLI\ykman.exe'

if (-not (Test-Path -Path $YubiCli)) {
    Write-Host "YubiKey Manager CLI (ykman.exe) missing."
    return
}

# get our passwords to save to the yubikey
$Zero = Get-Credential -Message "Enter zero account username and password"

# ask user if they want to send an enter key after yubikey password
$SendEnter = Read-Host "Do you want ENTER sent after your YubiKey is activated? (Y/N)"
# the arg to add/remove from the yubikey arguments
$EnterArgString = '--no-enter'
# check user's input for anything that is a Y or YES
if (($Null -ne $SendEnter) -and (($SendEnter.ToUpper() -eq "Y") -or ($SendEnter.ToUpper() -eq "YES"))) {
    $EnterArgString = $Null
    Write-Host "Enter will be sent after key activation." -ForegroundColor Cyan
}
else {
    Write-Host "Enter will NOT be sent after key activation." -ForegroundColor Cyan
}

# check if we got a 0 password from the user
if (($Null -ne $Zero) -and ($Zero.Length -gt 0)) {
    # get our mod hex version of our password
    $ZeroPass = ConvertFrom-SecureString $Zero.Password -AsPlainText
    $ZeroUser = "$($Zero.UserName)`t$(ConvertFrom-SecureString $Zero.Password -AsPlainText)"

    # update our first key with our 0 password
    & $YubiCli otp static $EnterArgString --force --keyboard-layout US 1 -- $ZeroPass
    # add a delay to our second slot
    & $YubiCli otp settings --force --pacing 40 2
    # update our second key with our login and 0 password
    & $YubiCli otp static $EnterArgString --force --keyboard-layout US 2 -- $ZeroUser
    Write-Host "Done."
}
else {
    Write-Host "Credentials were null.`nYubiKey not updated."
}
