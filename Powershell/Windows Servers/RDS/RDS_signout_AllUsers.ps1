# Get Connection Brokers in Remote Desktop Envrionment.
$RDConnectionBrokers = Get-RDServer -Role "RDS-CONNECTION-BROKER"
$PrimaryConnectionBroker = $RDConnectionBrokers.Server | Select-Object -Index 0
# Get active sessions on all Connection Brokers.
$ActiveUsers = Get-RDUserSession -ConnectionBroker $PrimaryConnectionBroker.Server
# Inform users of restart and that they will be logged off.
foreach ($ActiveUser in $ActiveUsers) {
    Send-RDUserMessage -HostServer $ActiveUser.HostServer -UnifiedSessionID $ActiveUser.UnifiedSessionId -MessageTitle "Message from administrator: Server restart" -MessageBody "Please save your work. You will be logged off in 15 minutes."
}
# Wait 15 minutes before signing out users..
Start-Sleep -Seconds 900
# Refresh active sessions on all Connection Brokers.
$ActiveUsers = Get-RDUserSession -ConnectionBroker $PrimaryConnectionBroker
# Sign out users.
foreach ($ActiveUser in $ActiveUsers){
    Invoke-RDUserLogoff -HostServer $ActiveUser.HostServer -UnifiedSessionID $ActiveUser.UnifiedSessionId -Force
}