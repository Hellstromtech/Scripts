#Script to set pwdlastset to negative and update to todays date and current time.

Import-Module ActiveDirectory

$OU = "OU=Test,OU=ITRUsers,OU=.ITRH,DC=int,DC=test,DC=dk"

$Users = Get-ADUser -Filter * -Properties * -SearchBase $OU
ForEach ($User in $Users) {
    Write-Host "Processing "$User -ForegroundColor Yellow
    $User = Get-ADUser $User -Properties pwdLastSet
    $User.pwdLastSet = 0
    Set-ADUser -Instance $User -Server "domainname.com"
    $User.pwdLastSet = -1
    Set-ADUser -Instance $User -Server "domainname.com"
    Write-Host "Processed "$User -ForegroundColor Green
}
REPADMIN /syncall