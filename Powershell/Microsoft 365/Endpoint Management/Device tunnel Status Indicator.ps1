#Sets the device tunnel to show in gui

New-Item -Path ‘HKLM:\SOFTWARE\Microsoft\Flyout\VPN’ -Force
New-ItemProperty -Path ‘HKLM:\Software\Microsoft\Flyout\VPN\’ -Name ‘ShowDeviceTunnelInUI’ -PropertyType DWORD -Value 1 -Force