# Random signs for your secure random password according to your definition.
function Get-RandomCharacters($length, $characters) {
    $random = 1..$length | ForEach-Object { Get-Random -Maximum $characters.length }
    $private:ofs=""
    return [String]$characters[$random]
}
 
# Scramble the string
function Scramble-String([string]$inputString){     
    $characterArray = $inputString.ToCharArray()   
    $scrambledStringArray = $characterArray | Get-Random -Count $characterArray.Length     
    $outputString = -join $scrambledStringArray
    return $outputString 
}

# Connect to Office 365
Connect-MsolService

# Get tenant initial domain name.
$primaryDomain = (Get-MsolDomain | where IsInitial -eq $true)[0].Name

# Ask for common Display name, something that blend in with all other users.
$DisplayName = Read-Host -Prompt "Input a normal name for a user, for example: Ulrik Andersen."

# Random number add to userUPN
$number = get-random -minimum 1000000 -maximum 9999999
$newUserUPN = "User$number@$primaryDomain"

# Password policy
$password = (Get-RandomCharacters -length 3 -characters 'abcdefghiklmnoprstuvwxyz')
$password += (Get-RandomCharacters -length 3 -characters 'ABCDEFGHKLMNOPRSTUVWXYZ')
$password += (Get-RandomCharacters -length 1 -characters '1234567890')
$password += (Get-RandomCharacters -length 9 -characters '!"$%&/\()=?}][{#*+')
# Scramble the string characters
$password = Scramble-String $password

# TODO: try/catch and error handling.
New-MsolUser -DisplayName $DisplayName -UserPrincipalName $newUserUPN -ForceChangePassword $false -Password $password -PasswordNeverExpires $true
Add-MsolRoleMember -RoleObjectId b1be1c3e-b65d-4f19-8427-f6fa0d97feb9 -RoleMemberEmailAddress $newUserUPN # b1be1c3e-b65d-4f19-8427-f6fa0d97feb9 = Role: Conditional Access Administrator


Write-Host "Remember to document user!" 
