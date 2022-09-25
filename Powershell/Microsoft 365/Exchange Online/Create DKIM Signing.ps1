#Load Modules
$ErrorOccured = $false
$modules = @("ExchangeOnlineManagement")

foreach ($module in $modules) {

try {import-module $module -ErrorAction Stop}
catch {$ErrorOccured = $true}

if (!$ErrorOccured) {
    write-host "PS module"$module" imported"
} else {
    write-host "Trying to install PS module"$module
    try {install-module $module -ErrorAction Stop}
    catch {write-host "Failed to install PS module"$module", are you running this as administrator?"}
}
}


# Variables

$tenantAdmin = Read-Host "Type admin name"

Connect-ExchangeOnline -UserPrincipalName $tenantAdmin -ShowBanner:$false

$domain = "contoso.com"


# Create DkimSigningConfig
Try {
    Get-DkimSigningConfig -Identity $domain
    Write-Host "test1"
}
Catch {
    New-DkimSigningConfig -DomainName $domain -Enabled $false
    Write-Host "test2"
}


# CNAME Records

$DkimSigningConfig = Get-DkimSigningConfig -Identity $domain
#Get-DkimSigningConfig -Identity $domain | Format-List Selector1CNAME, Selector2CNAME

$Selector1CNAMEWriteOut = "Record: CNAME, Name: selector1._domainkey."+$domain+", Content: "+$DkimSigningConfig.Selector1CNAME
$Selector2CNAMEWriteOut = "Record: CNAME, Name: selector2._domainkey."+$domain+", Content: "+$DkimSigningConfig.Selector2CNAME

Write-host $Selector1CNAMEWriteOut
Write-host $Selector2CNAMEWriteOut

# Set DKIM

Set-DkimSigningConfig -Identity $domain -Enabled $true