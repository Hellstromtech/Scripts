# Simple script to bulk change UPN suffix of users based on the OU they are placed in.
Import-Module ActiveDirectory

$oldSuffix = "INPUT OLD SUFFIX"
$newSuffix = "INPUT NEW SUFFIX"
$ou = "INPUT OU PATH FX OU=DK,OU=Users,OU=Customer,DC=int,DC=domain,DC=tld"
$server = "INPUT SERVER WITH AD TO USE"


Get-ADUser -SearchBase $ou -filter * | ForEach-Object {
    $newUpn = $_.UserPrincipalName.Replace($oldSuffix,$newSuffix)
    $_ | Set-ADUser -server $server -UserPrincipalName $newUpn
}