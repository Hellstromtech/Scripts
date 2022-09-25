# Patching

START BY REBOOTING THE SERVER!

setup.exe /PrepareAD /IAcceptExchangeServerLicenseTerms
setup.exe /PrepareSchema /IAcceptExchangeServerLicenseTerms
setup.exe /m:upgrade /IAcceptExchangeServerLicenseTerms


# Tests after patch
Go to https://localhost/OWA and send a e-mail to a external recipient, and reply to that email to verify mailflow

Get-ServerComponentState -Identity $env:computername
Test-ServiceHealth $env:computername
Test-ReplicationHealth $env:computername
Test-MAPIConnectivity
Test-OutlookWebservices -MailboxCredential user@domain.com (Get-credential)
Get-ExchangeServer | Get-Message


OWA & ECP not working after KB5004779
https://docs.microsoft.com/en-us/exchange/troubleshoot/administration/cannot-access-owa-or-ecp-if-oauth-expired