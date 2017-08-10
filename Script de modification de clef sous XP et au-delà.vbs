'--------------------------------------------------------------------
' Script de modification de clef sous XP et au-delà
' JCB © 2004
'--------------------------------------------------------------------
On Error Resume Next

Set args  = Wscript.Arguments
Set shell = WScript.CreateObject("WScript.Shell")

nbargs=args.count
if nbargs<1 then Syntaxe
PK=Ucase(Replace(args(0),"-",""))
If PK="/?" Then Syntaxe

Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
Set objOS=objWMI.ExecQuery("Select * from Win32_OperatingSystem",,48)
For each item in objOS
	Build=cint(item.BuildNumber)
	If build<2600 Then
		wscript.echo "Ce script ne fonctionne qu'à partir de Windows XP !"
		wscript.quit
		End If
	Next

'	Shell.RegDelete "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\WPAEvents\OOBETimer" 

Set objWPA=objWMI.ExecQuery("Select * from win32_WindowsProductActivation",,48)
for each item in objWPA
	result = item.SetProductKey(PK)
	If err=0 Then
		WScript.Echo "Le ""ProductKey"" a été modifié avec succès"
	Else
		WScript.Echo "Erreur : " & Err.Description &  " (0x" & Hex(Err.Number) & ")"
		End If
	Next
wscript.quit	
'--------------------------------------------------------------------
Sub Syntaxe
msg="Script de modification du ""ProductKey"" sous Windows XP (et au delà)" & VBCRLF & VBCRLF
msg=msg & "Syntaxe :" & VBCRLF
msg=msg & "  SetPK <productkey>" & VBCRLF
msg=msg & "Paramètre :" & VBCRLF
msg=msg & "  <productkey> : " & VBCRLF
msg=msg & "     la clef de produit figurant sur le coffret du CD" & VBCRLF
msg=msg & "     de la forme : ABCDE-FGHIJ-KLMNO-PRSTU-WYQZX" & VBCRLF 
wscript.echo msg
wscript.quit
End Sub