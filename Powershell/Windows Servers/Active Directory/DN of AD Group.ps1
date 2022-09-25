#Get's Distinguished name of AD Group

Get-ADGroup groupname -Properties DistinguishedName | Select-Object DistinguishedName 