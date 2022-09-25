 $oldverbose = $VerbosePreference
    $VerbosePreference = "continue"

    #Run netsh command to get wireless profiles, and store to a variable
    $command = "netsh.exe wlan show profiles"
    $output = Invoke-Expression $command
    $SSIDsResult = $Output | Select-String -Pattern 'All User Profile'

    #Parse Output by splitting on the colon and trim
    $SSIDs = ($SSIDsResult -split ":",2).Trim()

    #Count the number of elements in the array
    $NumElements = $SSIDs.GetUpperBound(0)
    Write-Verbose "Elements: $NumElements"
    Write-Verbose "$SSIDs"

    #Setup variables
    $TempIndex = 0
    $AllSSIDs = @()

    #Count through finding odd elements only to store to array, only the odd elements are profiles names
    While ($TempIndex -lt $NumElements) {
            $TempIndex = $TempIndex + 1
    
            #Do math to see if the Index is odd
            If ($TempIndex % 2 -eq 1) {
                #Debug output not sure why this does not work 
                Write-Verbose "($SSIDs[$TempIndex])"
        
                #Add the odd elements to the array        
                $AllSSIDs += ,@($SSiDs[$TempIndex])
            }
        }

    Write-Verbose "$AllSSIDs"

    $AllProfiles = @()

    foreach ($SSID in $AllSSIDs)
        {
        $PW='<no password / null value>'
        $ProfileName=''
        $PwSearchResult=''
    
        #Run netsh command to get wireless profiles and password store to a variable
        $str = $SSID
        $command = "netsh.exe wlan show profiles name=""$str"" key=clear"
        Write-Verbose "$command"
        $output = Invoke-Expression $command

        #Use select-String and search for 'SSID'
        $SSIDSearchResult = $output | Select-String -Pattern 'SSID Name'

        #Parse Output by splitting on the colon, grab the last element in the array, trim, and remove quotes
        $ProfileName = ($SSIDSearchResult -split ":",2)[-1].Trim() -replace '"'


        #Retrieve wireless profile password using the same method
        $PwSearchResult = $Output | Select-String -Pattern 'Key Content'
        If ($PwSearchResult -ne $null) { 
            $PW = ($PwSearchResult -split ":",2)[1].Trim() -replace '"'
            }
        Write-Verbose "$ProfileName $PW"

        #Add the SSID name and Password to the array    
        $Allprofiles += ,@($ProfileName,$PW)
        }



    #Format the array for output to the screen and clipboard
    $ProfilesTable = ($AllProfiles | % { $_ -join "`t" }) -join "`n"

    Write-Host ("$ProfilesTable") -ForegroundColor White -BackgroundColor DarkGray
    Write-Output "$ProfilesTable" | Clip
    Write-Host ("`nCopied to clipboard") -ForegroundColor Yellow -BackgroundColor Black
    $VerbosePreference = $oldverbose