# Import the ActiveDirectory module
Import-Module ActiveDirectory

# Prompt the user for the name of the user to disable
$userName = Read-Host "Enter the name of the user to disable"

# Search for the user in Active Directory
$user = Get-ADUser -Filter "Name -eq '$userName'"

# Check if the user was found
if ($user) {
    # Disable the user account
    Disable-ADAccount -Identity $user

    # Generate a random password that is at least 20 characters long
    $password = -Join ((65..90) + (97..122) | Get-Random -Count 20 | % {[char]$_})

    # Set the password for the user
    Set-ADAccountPassword -Identity $user -NewPassword (ConvertTo-SecureString $password -AsPlainText -Force)

    # Move the user to the Disabled Users OU
    $disabledUsersOU = "OU=Disabled Users,DC=example,DC=com"
    Move-ADObject -Identity $user -TargetPath $disabledUsersOU

    # Set the description for the user to the current date
    Set-ADUser -Identity $user -Description (Get-Date)

    # Clear the Office and Department attributes for the user
    Set-ADUser -Identity $user -Office "" -Department ""

    # Remove the user from all groups except for Domain Users
    $domainUsersGroup = Get-ADGroup -Filter "Name -eq 'Domain Users'"
    Get-ADPrincipalGroupMembership $user | Where-Object {$_.Name -ne $domainUsersGroup.Name} | ForEach-Object {Remove-ADPrincipalGroupMembership -Identity $user -MemberOf $_}
}
else {
    # The user was not found
    Write-Output "User not found"
}

# Connect to Office 365
$credential = Get-Credential
$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $credential -Authentication Basic -AllowRedirection
Import-PSSession $session

# Convert the mailbox to a shared mailbox
Set-Mailbox -Identity $user.UserPrincipalName -Type Shared

# Remove all licenses from the user
Set-MsolUser -UserPrincipalName $user.UserPrincipalName -UsageLocation "US" -LicenseOptions ""

# Assign the mailbox to the user's manager
$manager = Get-ADUser -Filter "SamAccountName -eq '$($user.Manager)'"
$mailbox = Get-Mailbox -Identity $user.UserPrincipalName
Set-Mailbox -Identity $mailbox -ManagedBy $manager.UserPrincipalName

# Remove the user's access to the mailbox
Set-Mailbox -Identity $mailbox -AccessRights FullAccess -User $manager.UserPrincipalName
Remove-MailboxPermission -Identity $mailbox -User $user.UserPrincipalName -AccessRights FullAccess

# Disconnect from Office 365
Remove-PSSession $session