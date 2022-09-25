#Changes the default sharing policy on all users to a certain access right

$users = Get-Mailbox -Resultsize Unlimited
foreach ($user in $users) {
Write-Host -ForegroundColor green "Setting permission for $($user.alias)..."
Set-MailboxFolderPermission -Identity "$($user.alias):\kalender" -User Default -AccessRights limiteddetails
}