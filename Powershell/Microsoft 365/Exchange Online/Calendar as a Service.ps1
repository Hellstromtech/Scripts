
.DESCRIPTION  
This scripts gives X (user or group) access to calendars for X users.
Default is all users, the accessrights can be changed under variables, so can the user/group.

 
.OUTPUTS 
It has a log, other than that, it just changes calendar permissions.
It's got a test function, if you remove the #/Hashtag from line 114 (the one that actually changes permissions)
 
.NOTES 
Created by Ludwig Møller
No required modules.
It does require you to create a securepassword for the user, this is done automatically, if the file is not present.


it outputs the file to "C:\office365Scripts\Keys\Credentials.txt"
If the file is not present, one will be made.

 


#Variables
$username                   = "ADMIN BRUGERNAVN"
[string]$Log                = "C:\Office365Scripts\Calendar.log"
[string]$UserToGiveAccess   ="Brugernavn@company.com OR groupname" #If the setting is org-wide, write "Default"
[string]$AccessRight        ="editor"
$Mailboxes                  = Get-Mailbox
$pathoffile                 = "C:\office365Scripts\Keys"
$filename                   = "credentials.txt"


#PasswordCheck
$totalpath                  = "$pathoffile\$filename"
if (!(Test-Path $totalpath)){
    New-Item -ItemType directory $pathoffile
    New-Item -path $pathoffile -Name "Credentials.txt" -ItemType file
    Write-Host "Created new file"
    $password = read-host "Enter the password of the user"
    $secureStringPwd = $password | ConvertTo-SecureString -AsPlainText -Force 
    $secureStringText = $secureStringPwd | ConvertFrom-SecureString 

    Set-Content $totalpath $secureStringText
    }
else
{
  Write-Host
}

Start-Transcript -Path $Log -Force


#login
$password = Get-Content $totalpath | ConvertTo-SecureString
$psCred = New-Object System.Management.Automation.PSCredential -ArgumentList ($username, $password)
$session = New-PSSession –ConnectionUri https://ps.outlook.com/powershell –AllowRedirection –Authentication Basic –Credential $psCred –ConfigurationName Microsoft.Exchange
$Import = Import-Pssession $Session -AllowClobber
Import-Module MSOnline -verbose


#################################################################################################################
##############################################START OF ACTUAL SCRIPT#############################################
#################################################################################################################




#################################################################################################################
#################################################################################################################
#################################################################################################################
<# If it's for all users, just use this:
#################################################################################################################
#################################################################################################################
#################################################################################################################


foreach($mbx in Get-Mailbox -ResultSize Unlimited | where-object {$_.displayname -notmatch "discovery"})
                        {
                            $languageCalendar = (Get-MailboxFolderStatistics -Identity $mbx.userprincipalname -FolderScope Calendar | Select-Object -first 1).name
                            Set-MailboxFolderPermission -Identity ($mbx.UserPrincipalName+":\$LanguageCalendar") -User Default -AccessRights $accessright
                            Get-MailboxFolderPermission -Identity ($mbx.UserPrincipalName+":\$LanguageCalendar") -User Default | Select-Object identity,accessrights
                        }


#>

#################################################################################################################
#################################################################################################################
#################################################################################################################
#################################################################################################################
<# If you want to seperate it some more, and not just set default for everyone, use this section
#################################################################################################################
#################################################################################################################
#################################################################################################################
#################################################################################################################

$needpermissions= @()
foreach ($mbx in $Mailboxes){
    $CalendarIdentity = (($mbx.SamAccountName) + ":\" + (Get-MailboxFolderStatistics -Identity $mbx.Identity -FolderScope Calendar | Select-Object -First 1).Name)
    $CalendarPermissions = Get-MailboxFolderPermission -Identity $CalendarIdentity
    $calendarpermissions = $calendarpermissions.user
    $calendarpermissions = $calendarpermissions.displayname
    if($CalendarPermissions -contains "Airplanorama"){
    echo "$CalendarIdentity is OK"}else{
    echo "$mbx needs permissions" 
    $needpermissions += $mbx.UserPrincipalName
    }

}


foreach ($mbx in $needpermissions)
{
    $CalendarIdentity = (($mbx) + ":\" + (Get-MailboxFolderStatistics -Identity $mbx -FolderScope Calendar | Select-Object -First 1).Name)
    $CalendarPermissions = Get-MailboxFolderPermission -Identity $CalendarIdentity
    $calendarpermissions
 

    [bool]$PermissionExist = $false
    foreach ($cp in $CalendarPermissions)
    {
        if ("$($cp.User)" -ieq "$($UserToGiveAccess)")
        {
            if ("$($cp.AccessRights)" -ine $AccessRight)
            {
                # If $UserToGiveAccess already has permission then change it
                Write-Host "Editing permission on '$($mbx)' calendar."
                Set-MailboxFolderPermission -Identity $CalendarIdentity -User $UserToGiveAccess -AccessRights $AccessRight -Confirm:$false -WarningAction SilentlyContinue
            }
            else
            {
                Write-Host "Permission is OK for '$($mbx)' calendar."
            }

 

            # Uncomment to remove permission
            #Write-Host "Removing permission on '$($mbx.DisplayName)' calendar."
            #Remove-MailboxFolderPermission -Identity $CalendarIdentity -User $UserToGiveAccess -Confirm:$false #-WhatIf
            
            $PermissionExist = $true
            break
        }
    }

 

    if (-not $PermissionExist)
    {
        # If $UserToGiveAccess haven't got permission then add it
        Write-Host "Adding permission on '$($mbx)' calendar."
        Add-MailboxFolderPermission -Identity $CalendarIdentity -User $UserToGiveAccess -AccessRights $AccessRight -Confirm:$false -WarningAction SilentlyContinue #-WhatIf
    }
}

#>

#################################################################################################################
#################################################################################################################
#################################################################################################################


try
{
    Stop-Transcript | Out-Null
    [string]::Join("`r`n",(Get-Content $Log)) | Out-File $Log
}
catch
{}
Exit


