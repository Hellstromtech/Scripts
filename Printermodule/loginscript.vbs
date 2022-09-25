'Version 1.3
'F�r anv�ndning p� Redekos server hos Atea
'F�r att detta skript skall fungera s� m�ste filen Val_av_standardskrivare.hta ligga i samma mapp alternativt m�ste hela s�kv�gen pekas ut.
'Detta skript samt Val_av_standardskrivare.hta kontrollerar och spar 2st nycklar i registret i s�kv�gen HKCU\Software\Redeko\
'STD-Skrivare	- anger skrivaren som anv�ndaren valt som standardskrivare i HTA dialogen
'VisaSTDDialog	- anger om man vill att HTA dialogen skall visas vid varje inloggning, kan �ndras av anv�ndaren i HTA dialogen
'				  HTA dialogen visas ox� om ingen matchning g�rs

'Variabler:

' Variabelnamn: numSecsDelay
' Datatyp: Heltal
'F�r att ha en delay innan skrivare kontrolleras, s�tt numSecsDelay till n�got v�rde h�gre �n 0, t.ex. 120 f�r 120 sekunders delay/sleep innan k�rningen forts�tter
'Detta skall fungera p� senare versioner av Windows eftersom d� exekverar logon skript asynkront, allts� inloggningen forts�tter utan paus.
'Alternativt s� anges numSecsDelay till 0 och s� anges maxPrinterSearchInSecs nedan till t.ex. 60 s� f�rs�ker skriptet hitta matchning i max 60 sekunder innan det forts�tter.
'Om maxPrinterSearchInSecs �r lika med 0 (eller mindre �n vad det tar att kontrollera matchning) s� g�rs endast en kontroll om skrivaren laddats in innan det forts�tter.
numSecsDelay = 0 'Standard=0

' Variabelnamn: maxPrinterSearchInSecs
' Datatyp: Heltal
'Ange hur i hur m�nga sekunder skriptet skall f�rs�ka hitta matchning med vald skrivare och inladdade skrivare
'Det k�rs i en loop tills antingen matchning hittats eller att maxPrinterSearchInSecs < k�rningstid eller om ingen skrivare har valts av anv�ndaren tidigare.
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

'B�rja k�rningen...
If numSecsDelay> 0 then
	WScript.Sleep (numSecsDelay * 1000)
End If

startTime=now
totSecs=0
hasMatch=False

Do 
	'H�mta sparad standarskrivare (om valt)
	strDefault = GetDefaultFromRegistry

	If strDefault = "" Then
		Exit Do
	End If 

	strDefaultOrg = strDefault
	'H�mta om HTA dialogen skall visas oavsett om standardskrivare valts (eller hittas)
	strVisaDialog = GetVisaDialogFromRegistry

	'Kontrollera om det �r en skrivare som plockats med i sessionen f�r anv�ndaren som valts som standard
	'Om s� �r fallet hitta vad den skrivaren heter nu typ: Brother 1234 (redirected 99)
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

'St�ll in standardskrivare om den valts och hittats
'Vid fel s� visas dialogen oavsett vad man valt tidigare
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

'Kontrollera om dialog beh�ver visas
If strDefault = "" Or strVisaDialog = "1" Then
	Set objShell = CreateObject("Wscript.Shell")
	objShell.Run "Val_av_standardskrivare.hta"
Else
	'ingen dialog beh�vs
End If