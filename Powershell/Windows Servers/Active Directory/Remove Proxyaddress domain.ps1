Get-AdUser  -Filter * -SearchBase "dc=domain,dc=local" –prop proxyaddresses |
    ForEach-Object{
        if($pr =  $_.proxyaddresses  -like '*@domain.local*'){
            Set-ADUser $_ -Remove @{proxyaddresses=$pr}
        }
    } 