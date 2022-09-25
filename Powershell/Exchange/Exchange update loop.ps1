#Creates a loop that sets services to automatic when services dont come online for exchange updates.
#Run script while updating exchange server.

While (1 -le 2) { Sleep 1 ; Get-Service | Where { $_.DisplayName -Like ‘Microsoft Exchange*’ } | Set-Service –StartupType ‘Automatic’ }
