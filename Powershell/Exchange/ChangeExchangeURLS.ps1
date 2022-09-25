$URL = "https://mail.domain.com" # REMEMBER TO CHANGE THIS URL BEFORE RUNNING.

Set-ClientAccessServer -Identity "$env:COMPUTERNAME" -AutodiscoverServiceInternalURI "$URL/Autodiscover/Autodiscover.xml"
Set-OwaVirtualDirectory -Identity "$env:COMPUTERNAME\owa (Default Web Site)" -InternalUrl "$URL/owa" -ExternalUrl "$URL/owa"
Set-MapiVirtualDirectory -Identity "$env:COMPUTERNAME\mapi (Default Web Site)" -InternalUrl "$URL/mapi" -ExternalUrl "$URL/mapi"
Set-OabVirtualDirectory -Identity "$env:COMPUTERNAME\oab (Default Web Site)" -InternalUrl "$URL/oab" -ExternalUrl "$URL/oab"
Set-EcpVirtualDirectory -Identity "$env:COMPUTERNAME\ecp (Default Web Site)" -InternalUrl "$URL/ecp" -ExternalUrl "$URL/ecp"
Set-WebServicesVirtualDirectory -Identity "$env:COMPUTERNAME\ews (Default Web Site)" -InternalUrl "$URL/EWS/Exchange.asmx" -ExternalUrl "$URL/EWS/Exchange.asmx"
Set-ActiveSyncVirtualDirectory -Identity "$env:COMPUTERNAME\Microsoft-Server-ActiveSync (Default web site)" -InternalUrl "$URL/Microsoft-Server-ActiveSync" -ExternalUrl "$URL/Microsoft-Server-ActiveSync"
