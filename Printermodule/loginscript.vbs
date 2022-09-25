'Version 1.3
'För användning på Redekos server hos Atea
'För att detta skript skall fungera så måste filen Val_av_standardskrivare.hta ligga i samma mapp alternativt måste hela sökvägen pekas ut.
'Detta skript samt Val_av_standardskrivare.hta kontrollerar och spar 2st nycklar i registret i sökvägen HKCU\Software\Redeko\
'STD-Skrivare	- anger skrivaren som användaren valt som standardskrivare i HTA dialogen
'VisaSTDDialog	- anger om man vill att HTA dialogen skall visas vid varje inloggning, kan ändras av användaren i HTA dialogen
'				  HTA dialogen visas oxå om ingen matchning görs

'Variabler:

' Variabelnamn: numSecsDelay
' Datatyp: Heltal
'För att ha en delay innan skrivare kontrolleras, sätt numSecsDelay till något värde högre än 0, t.ex. 120 för 120 sekunders delay/sleep innan körningen fortsätter
'Detta skall fungera på senare versioner av Windows eftersom då exekverar logon skript asynkront, alltså inloggningen fortsätter utan paus.
'Alternativt så anges numSecsDelay till 0 och så anges maxPrinterSearchInSecs nedan till t.ex. 60 så försöker skriptet hitta matchning i max 60 sekunder innan det fortsätter.
'Om maxPrinterSearchInSecs är lika med 0 (eller mindre än vad det tar att kontrollera matchning) så görs endast en kontroll om skrivaren laddats in innan det fortsätter.
numSecsDelay = 0 'Standard=0

' Variabelnamn: maxPrinterSearchInSecs
' Datatyp: Heltal
'Ange hur i hur många sekunder skriptet skall försöka hitta matchning med vald skrivare och inladdade skrivare
'Det körs i en loop tills antingen matchning hittats eller att maxPrinterSearchInSecs < körningstid eller om ingen skrivare har valts av användaren tidigare.
maxPrinterSearchInSecs = 60 'Standard=60

'Slut - Variabler

Function GetDefaultFromRegistry
On Error Resume Next
Dim MyKey
	Set WshShell = CreateObject("WScript.Shell")
	MyKey = WshShell.RegRead("HKCU\Software\Redeko\STD-Skrivare")
	If Err <> 0 Then
		'MsgBox Err.Number & vbcr & Err.Description,vbCritical,Title
	MyKey = ""
	Else
		'MsgBox MyKey,vbInformation,Title
	End If  
	GetDefaultFromRegistry = MyKey
End Function

Function GetVisaDialogFromRegistry
On Error Resume Next
Dim MyKey
	Set WshShell = CreateObject("WScript.Shell")
	MyKey = WshShell.RegRead("HKCU\Software\Redeko\VisaSTDDialog")
	If Err <> 0 Then
		'MsgBox Err.Number & vbcr & Err.Description,vbCritical,Title
	MyKey = "1"
	Else
		'MsgBox MyKey,vbInformation,Title
	End If  
	GetVisaDialogFromRegistry = MyKey
End Function

Function GetMatch(strTest)
On Error Resume Next
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
	Set colPrinters = objWMIService.ExecQuery("Select * From Win32_Printer")
	For Each objPrinter in colPrinters
		strPrinter = objPrinter.Name
		If strPrinter <> "" And InStr(strPrinter,"(redirected")>0 Then
			stIndex1 = InStr(strPrinter,"(redirected")
			strDef1 = Left(strPrinter,stIndex1)
			if strTest = strDef1 Then
				GetMatch = strPrinter
				Exit For
			end If
		ElseIf strPrinter <> "" Then 
			if strTest = strPrinter Then
				GetMatch = strPrinter
				Exit For
			end If
		End if
	Next
End Function

'Börja körningen...
If numSecsDelay> 0 then
	WScript.Sleep (numSecsDelay * 1000)
End If

startTime=now
totSecs=0
hasMatch=False

Do 
	'Hämta sparad standarskrivare (om valt)
	strDefault = GetDefaultFromRegistry

	If strDefault = "" Then
		Exit Do
	End If 

	strDefaultOrg = strDefault
	'Hämta om HTA dialogen skall visas oavsett om standardskrivare valts (eller hittas)
	strVisaDialog = GetVisaDialogFromRegistry

	'Kontrollera om det är en skrivare som plockats med i sessionen för användaren som valts som standard
	'Om så är fallet hitta vad den skrivaren heter nu typ: Brother 1234 (redirected 99)
	If strDefault <> "" And InStr(strDefault,"(redirected")>0 Then
		stIndex = InStr(strDefault,"(redirected")
		strDef = Left(strDefault,stIndex)
		'msgbox strDef
		strDefault2 = GetMatch(strDef)
		'msgbox strDefault2
		strDefault = strDefault2
		If strDefault2 <> "" then
			hasMatch=True
		End If 
	ElseIf strDefault <> "" Then 
		strDefault2 = GetMatch(strDefault)
		'msgbox strDefault2
		strDefault = strDefault2
		If strDefault2 <> "" then
			hasMatch=True
		End If 
	Else
		'Bryt loopen
		hasMatch=True
	End If

	totSecs = DateDiff("s",startTime,now)
	If totSecs = 0 Then
		totSecs = 1
	End If 

Loop While hasMatch = False And totSecs < maxPrinterSearchInSecs

'msgbox strDefault
'msgbox totSecs
'msgbox hasMatch

'Ställ in standardskrivare om den valts och hittats
'Vid fel så visas dialogen oavsett vad man valt tidigare
If strDefault <> "" Then
	On Error Resume next
	Set WshNetwork = CreateObject("Wscript.Network")
    WshNetwork.SetDefaultPrinter strDefault
	If Err <> 0 Then
		strVisaDialog = "1"
	End If
End If

'Om fel skall visas avmarkera nedan
'On Error GoTo 0

'Kontrollera om dialog behöver visas
If strDefault = "" Or strVisaDialog = "1" Then
	Set objShell = CreateObject("Wscript.Shell")
	objShell.Run "Val_av_standardskrivare.hta"
Else
	'ingen dialog behövs
End If