#give 1 user access to all calendars in a tenant.

$userToAdd = "itr_cloudadmin@kielbergadvokater.onmicrosoft.com"
$users = Get-Mailbox | Select -ExpandProperty PrimarySmtpAddress
Foreach ($u in $users)
{
$ExistingPermission = Get-MailboxFolderPermission -Identity $u":\calendar" -User $userToAdd -EA SilentlyContinue
if ($ExistingPermission) {Remove-MailboxFolderPermission -Identity $u":\calendar" -User $userToAdd -Confirm:$False}
}