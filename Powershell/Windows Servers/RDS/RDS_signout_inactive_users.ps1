#Signs out inactive users

$RDConnectionBrokers = Get-RDServer -Role RDS-CONNECTION-BROKER
$PrimaryConnectionBroker = $RDConnectionBrokers.Server | Select-Object -Index 0

$InactiveUsers = Get-RDUserSession -ConnectionBroker $PrimaryConnectionBroker | Where-Object {$_.SessionState -eq "STATE_DISCONNECTED"}

foreach ($InactiveUser in $InactiveUsers){
    Invoke-RDUserLogoff -HostServer $InactiveUser.HostServer -UnifiedSessionID $InactiveUser.UnifiedSessionId -Force
}