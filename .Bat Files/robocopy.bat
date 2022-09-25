Robocopy E:\Dokumenter L:\Delte /E /R:1 /ZB /MT /COPYALL /V /TEE 

Robocopy \\dc10\Public_All d:\Shared\Public_All /E /R:1 /ZB /MT /COPYALL /V /TEE /UNILOG:C:\Public_All.txt

Robocopy \\dc10\User \\dc01\User /R:10 /E /COPYALL /LEV:2 /LOG:C:\Copyresults.txt

robocopy E:\Dokumenter L:\Delte /IS /ZB /MIR /r:1 /w:1 /MT /copy:DAT /log:"driveletter:\destination\file directory\backup.log"
