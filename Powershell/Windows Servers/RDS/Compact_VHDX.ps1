#Script to compact VHDX files.

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true, ValueFromPipeline)]
    [string]
    $VHDXStoreServer,

    [Parameter(Mandatory = $true, ValueFromPipeline)]
    [string]
    $LocalVHDXLocation,

    [Parameter(Mandatory = $false)]
    [string]
    $LocalLogPath = "C:\Logs\RDS-CompactVHDX\"
)

# Get the information needed for logging.
$dateTime = $(get-date -f yyyy-ddMM-HHmm)
$dateTimeSeconds = $(get-date -f yyyy-ddMM-HHmm)

if (!(Test-Path -path $LocalLogPath)) {
    New-Item $LocalLogPath -Type Directory -Force
}

# Get all .vhdx-files.
$AllPathVHDX = Get-ChildItem $LocalVHDXLocation -Recurse -Filter *.vhdx -Depth 1 | Where-Object name -ne UVHD-template.vhdx | Select-Object -ExpandProperty fullname
# Get all folders containing .vhdx-files to use for UserID matching.
$UserIDs = Get-ChildItem -Path $LocalVHDXLocation -Name

# Iterate through all .vhdx-files to Compact them and log, based on UserID matching.
foreach ($PathVHDX in $AllPathVHDX){
    $pathArray = $pathVHDX.Split('\')
    foreach($UserID in $UserIDs) {
        foreach($pathNode in $pathArray) {
            if($UserID -eq $pathNode) {
                $logFile = "$LocalLogPath$dateTime-$pathNode.txt"
                New-Item -Path $logFile -ItemType file -force | OUT-NULL
                Add-Content -Path $logFile "[$dateTimeSeconds] Started compacting vhdx."
                Add-Content -Path $logFile "select vdisk file= $pathVHDX"
                Add-Content -Path $logFile "compact vdisk"
                diskpart.exe /S $logFile >> "$logFile"
                Add-Content -Path $logFile "[$dateTimeSeconds] Finished compacting vhdx."
            }
        }
    }
}

# Clean up so that we do not have logs older than 3 days.
Get-ChildItem –Path $LocalLogPath -Recurse | Where-Object {($_.LastWriteTime -lt (Get-Date).AddDays(-30))} | Remove-Item