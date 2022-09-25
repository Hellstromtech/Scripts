# Path to folders containing .vhdx-files.
$rootPath = "E:\TS_ProfileDisk"
# Path to put log-files.
$logPath = "C:\Temp\Logs\compactUPD\"

# Get date and time.
$dateTime = $(get-date -f yyyy-ddMM-HHmm)
$dateTimeSeconds = $(get-date -f yyyy-ddMM-HHmm)

# Get all .vhdx-files.
$allPathUPD = Get-ChildItem $rootPath -Recurse -Filter *.vhdx -Depth 0 | Where-Object name -ne UVHD-template.vhdx | Select-Object -ExpandProperty fullname

# Iterate through all .vhdx-files to Compact them and log, based on UserID matching.
foreach ($pathUPD in $allPathUPD){
    $pathArray = $pathUPD.Split('\')
    $pathNode = $pathArray[2]
    $logFile = "$logPath$dateTime-$pathNode.txt"
    New-Item -Path $logFile -ItemType file -force | OUT-NULL
    Add-Content -Path $logFile "select vdisk file= $pathVHDX"
    Add-Content -Path $logFile "compact vdisk"
    .\diskpart.exe /S $logFile >> "$logFile"
    Add-Content -Path $logFIle "[$dateTimeSeconds] Finished compacting vhdx."
}

# Clean up so that we do not have logs older than 2 days.
forfiles /p "$logPath" /s /m *.* /c "cmd /c del @path /Q" /d -3