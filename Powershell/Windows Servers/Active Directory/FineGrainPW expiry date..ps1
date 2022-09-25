#Gets expiry date if user has fine grain password policy.

Get-ADUser "username" -Properties msDS-UserPasswordExpiryTimeComputed, PasswordLastSet, CannotChangePassword | select Name, @{Name="ExpiryDate";Expression={[datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")}}, PasswordLastSet