#Requires -Version 4.0

<#

 Convert a DHCP lease address to a static IP address assignment on a remote machine.
 Uses PSRemoting to detect IP settings and a scheduled task to make the IP configuration
 change.

.PARAMETER ComputerName
 The remote computer(s) to convert to static.

.PARAMETER Credential
 The credential to use for the IP conversion.

#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact='High')]
param(
    
    [Parameter(Mandatory)]
    [string[]]
    $ComputerName,

    [pscredential]
    $Credential

)

begin {
    
    $CredentialSplat = @{}
    if ( $Credential ) { $CredentialSplat.Credential = $Credential }

    if ( $PSBoundParameters.Keys -notcontains 'ErrorAction' ) { $ErrorActionPreference = 'Stop' }
    
}

process {
    
    foreach ( $ComputerNameItem in $ComputerName ) {
        
        $RemoteComputerIP = Resolve-DnsName -Name $ComputerNameItem -Type A -DnsOnly -ErrorAction SilentlyContinue -Verbose:$false |
            Select-Object -ExpandProperty IPAddress

        if ( -not $RemoteComputerIP ) {

            Write-Error ( 'Could not resolve {0} to an IPv4 address' -f $ComputerNameItem )

            continue

        }

        $Caption     = 'Convert IP address to static?'
        $Description = 'Update computer {0} and convert DHCP IP address {1} to static' -f $ComputerNameItem, $RemoteComputerIP

        if ( $PSCmdlet.ShouldProcess($Description, $null, $Caption) ) {

            $Session = New-PSSession -ComputerName $ComputerNameItem @CredentialSplat

            if ( -not( $RemotePSVersion = Invoke-Command -Session $Session -ScriptBlock { $PSVersionTable } ) -or $RemotePSVersion.PSVersion -lt [version]'4.0' ) {

                Write-Error ( 'PowerShell version on remote host {0} is not supported. Detected version {1}.' -f $ComputerNameItem, $RemotePSVersion.PSVersion ) -ErrorAction Continue

                Remove-PSSession -Session $Session

                continue

            }

            $NetIPAddress = Invoke-Command -Session $Session -ScriptBlock { Get-NetIPAddress -IPAddress $Using:RemoteComputerIP }
            $NetRoute = Invoke-Command -Session $Session -ScriptBlock { Get-NetRoute -DestinationPrefix '0.0.0.0/0' -InterfaceAlias $Using:NetIPAddress.InterfaceAlias }
            $DnsClientServerAddress = Invoke-Command -Session $Session -ScriptBlock { Get-DnsClientServerAddress -InterfaceAlias $Using:NetIPAddress.InterfaceAlias -AddressFamily IPv4 }
                
            if ( $NetIPAddress.PrefixOrigin -eq 'Manual' ) {

                Write-Warning ( 'Static IP address is already defined on {0}' -f $ComputerNameItem )

            }
                
            if ( $NetIPAddress.PrefixOrigin -eq 'Dhcp' -and $null -ne $NetRoute.NextHop ) {


                $UpdateCommand = @(
                    #'Start-Transcript -Path "C:\static_ip.log"'
                    'Remove-NetIPAddress -InterfaceAlias {0} -IPAddress {1} -Confirm:$false' -f $NetIPAddress.InterfaceAlias, $NetIPAddress.IPv4Address
                    'New-NetIPAddress -InterfaceAlias {0} -AddressFamily IPv4 -IPAddress {1} -PrefixLength {2} -DefaultGateway {3}' -f $NetIPAddress.InterfaceAlias, $NetIPAddress.IPAddress, $NetIPAddress.PrefixLength, $NetRoute.NextHop
                    'Set-DnsClientServerAddress -InterfaceAlias {0} -ServerAddresses {1}' -f $NetIPAddress.InterfaceAlias, ( $DnsClientServerAddress.ServerAddresses -join ', ' )
                ) -join '; '

                Write-Verbose $UpdateCommand

                $TaskStart = Get-Date
    
                $TaskSplat = @{
                    Action      = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "-NoProfile -Command `"$($UpdateCommand.Replace('"', '\"') )`""
                    Trigger     = New-ScheduledTaskTrigger -Once -At $TaskStart.AddSeconds(10)
                    Description = 'Task created by PowerShell'
                    Settings    = New-ScheduledTaskSettingsSet -StartWhenAvailable -DisallowDemandStart -DeleteExpiredTaskAfter 30
                    Principal   = New-ScheduledTaskPrincipal -UserID "NT AUTHORITY\SYSTEM" -LogonType ServiceAccount -RunLevel Highest
                }

                $Task = New-ScheduledTask @TaskSplat |
                    %{ $_.Triggers[0].EndBoundary = $TaskStart.ToUniversalTime().AddSeconds(30).ToString('u').Replace(' ', 'T'); $_ }

                if ( Invoke-Command -Session $Session -ScriptBlock { $Using:Task | Register-ScheduledTask -TaskName 'Configure Static IP' } ) {

                    Write-Information ( 'Scheduled task created to convert {0} from DHCP to static IP assignment. Make sure you convert the lease in DHCP to a reservation.' -f $ComputerNameItem ) -InformationAction Continue

                }

            }

            Remove-PSSession -Session $Session

        }
        
    }
    
}
