# Opretter IIS-directories på F-drevet
New-Item -ItemType Directory -Path F:\inetpub\logs\LogFiles\W3SVC1
New-Item -ItemType Directory -Path F:\inetpub\logs\LogFiles\W3SVC2

# Sætter IIS-directory logging i IIS administrations modulet.  
Import-Module WebAdministration
Set-ItemProperty "IIS:\Sites\Default Web Site" -name logFile.directory  -value "F:\Inetpub\logs\LogFiles"
Set-ItemProperty "IIS:\Sites\Exchange Back End" -name logFile.directory  -value "F:\Inetpub\logs\LogFiles"