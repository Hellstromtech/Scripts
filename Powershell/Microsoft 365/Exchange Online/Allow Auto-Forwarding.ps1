<# 

.DESCRIPTION  
Creates a policy and a filter in antispam section of EOP/M365: https://protection.office.com/antispam
Also creates a new group, add anyone who should be able to autoforward externally to this group.

 
.OUTPUTS 
N/A

.NOTES 
Requires a ExchangeOnlineManagement Module, also requires the Azure AD module, for creating the group, both are installed on Citrix.


Made by Ludwig Møller



#>




#Variables#
$OutboundPolicyName = "Allow Outbound Policy"
$OutboundRuleName   = "Allow Auto-Forwarding  - Outbound Policy"
$AllowSendersGroup  = "Allow_External_Forwarding"



#Login
Get-Module exchangeonlinemanagement | Out-Null
Get-Module azuread | Out-Null
$Cred = Get-Credential
Connect-ExchangeOnline -Credential $cred
Connect-AzureAD -Credential $cred



#Creates a new group
New-UnifiedGroup -DisplayName $AllowSendersGroup -Notes "Allows users in this group to auto-forward to external emails."

#Creates the policies.
New-HostedOutboundSpamFilterPolicy -Name $OutboundPolicyName -ActionWhenThresholdReached BlockUserForToday -AutoForwardingMode "on" -AdminDisplayName "Allows auto forwarding for the users in 'applies to'"
New-HostedOutboundSpamFilterRule -Name $OutboundRuleName -HostedOutboundSpamFilterPolicy $OutboundPolicyName -FromMemberOf $AllowSendersGroup
