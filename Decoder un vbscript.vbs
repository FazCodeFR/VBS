'=======================================================================================================================================================
'Auteur : Brughes
'Note : vous pouvez écouter ou télécharger ma musique gratuitement sur : http://soundcloud.com/cyberflaneur ou http://www.jamendo.com/fr/artist/Brughes
'Description : Décode les fichiers vbe encodés par screnc.exe en utilisant des boîtes de dialogue pour les fichiers d'entrée (vbe) et de sortie (vbs).
'Inspiré de decovbe.vbs de Jean-Luc Antoine à retrouver sur www.interclasse.com/scripts/decovbe.php
'L'algorythme fonctionne également pour les fichiers .je .vbe .asp .hta .htm .html ...
'=======================================================================================================================================================

Option explicit

Dim strComputer, fd, fs, pe, ext, fic, Contenu, DebutCode, FinCode, objDialog, Srce, WshShell, fso, Dest
Const ForReading = 1
Const TagInit = "#@~^"    '#@~^awQAAA==
Const TagFin = "==^#~@"   '& chr(0)

strComputer = "."

'Quitter si le script est déjà lancé
If AppPrevInstance() = True Then TerminateApp()

'Ouvrir une boîte de dialogue, pointant sur le répertoire courant, pour lire le script à encoder
Set objDialog = CreateObject("UserAccounts.CommonDialog")
objDialog.Filter = "VBScript (*.vbe)|*.vbe"
objDialog.FilterIndex = 1
objDialog.Flags = 0
Set WshShell = WScript.CreateObject("WScript.Shell")
objDialog.InitialDir = WshShell.CurrentDirectory
If objDialog.ShowOpen Then
        Srce = objDialog.FileName
Else
        TerminateApp()
End If

If Srce = "" Then TerminateApp()

'Convertir Srce en nom court
Set fso = CreateObject("Scripting.FileSystemObject")
If Right(Srce, 1) = "\" Then Srce = Left(Srce, Len(Srce) - 1)
Srce = fso.GetAbsolutePathName(Srce)
If fso.FileExists(Srce) Then
	Srce = fso.GetFile(Srce).ShortPath
Else
        TerminateApp()
End If

'Ouvrir une boîte de dialogue, pointant sur le répertoire courant, pour enregistrer le fichier vbscript encodé
Set objDialog = CreateObject("SAFRCFileDlg.FileSave")
Set WshShell = WScript.CreateObject("WScript.Shell")
objDialog.FileName = WshShell.CurrentDirectory & "\*.vbs"
Set WshShell = Nothing
objDialog.FileType = "VBScript (*.vbs)"
If objDialog.OpenFileSaveDlg() Then
        Dest = objDialog.FileName
        If Right(Dest, 1) = "\" Then Dest = Left(Dest, Len(Dest) - 1)
        Dest = fso.GetAbsolutePathName(Dest)
Else
        TerminateApp()
End If

If Dest = "" Then TerminateApp()

'Décodage
Set fs = fso.GetFile(Srce)
pe = InstrRev(fs.Name,".")
If pe > 0 Then ext = LCase(Mid(fs.Name, pe))
If ext <> ".vbe" Then TerminateApp()

If Srce <> "" Then
	If fso.FileExists(Srce) Then
		Set fic = fso.OpenTextFile(Srce, ForReading, False)
		Contenu = fic.readAll
		fic.close
		Set fic = Nothing
		Do
			FinCode = 0
			DebutCode = Instr(Contenu, TagInit)
			If DebutCode > 0 Then
				If (Instr(DebutCode, Contenu, "==") - DebutCode) = 10 Then 'If "==" follows the tag
					FinCode = Instr(DebutCode, Contenu, TagFin)
					If FinCode > 0 Then
						Contenu = Left(Contenu, DebutCode - 1) & Decode(Mid(Contenu, DebutCode + 12, FinCode - DebutCode - 12 - 6)) & Mid(Contenu, FinCode + 6)
					End If
				End If
			End If
		Loop Until FinCode = 0
		Set fd = fso.CreateTextFile(Dest, True, False)
		fd.Write Contenu
		fd.Close
	Else
		WScript.Echo Srce & " not found"
	End If
Else
	TerminateApp()
End If

TerminateApp()


'==================================================================================================
'Fonctions et procédures
'==================================================================================================

Function AppPrevInstance()
'Vérifier si un script portant le même nom que le présent script est déjà lancé
        Dim objWMIService, colScript, objScript, RunningScriptName, Counter
        Counter = 0
        Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
        Set colScript = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'Wscript.exe' OR Name = 'Cscript.exe'")
        For Each objScript In colScript
                RunningScriptName = Mid(objScript.CommandLine, InstrRev(objScript.CommandLine, "\", -1, 1) + 1, Len(objScript.CommandLine) - InstrRev(objScript.CommandLine, "\", -1, 1) - 2)
                If WScript.ScriptName = RunningScriptName Then Counter = Counter + 1
		Wscript.Sleep 100
        Next
        If  Counter > 1 Then
                AppPrevInstance = True
        Else
                AppPrevInstance = False
        End If
        Set colScript = Nothing
        Set objWMIService = Nothing
End Function

Sub TerminateApp()
'Effacer les objets en mémoire et quitter
	Set fso = Nothing
        WScript.Quit
End Sub

Function Decode(Chaine)
	Dim se,i,c,j,index,ChaineTemp
	Dim tDecode(127)
	Const Combinaison="1231232332321323132311233213233211323231311231321323112331123132"

	Set se=WSCript.CreateObject("Scripting.Encoder")
	For i=9 to 127
		tDecode(i)="JLA"
	Next
	For i=9 to 127
		ChaineTemp=Mid(se.EncodeScriptFile(".vbs",string(3,i),0,""),13,3)
		For j=1 to 3
			c=Asc(Mid(ChaineTemp,j,1))
			tDecode(c)=Left(tDecode(c),j-1) & chr(i) & Mid(tDecode(c),j+1)
		Next
	Next
	'Next line we correct a bug, otherwise a ")" could be decoded to a ">"
	tDecode(42)=Left(tDecode(42),1) & ")" & Right(tDecode(42),1)
	Set se=Nothing

	Chaine=Replace(Replace(Chaine,"@&",chr(10)),"@#",chr(13))
	Chaine=Replace(Replace(Chaine,"@*",">"),"@!","<")
	Chaine=Replace(Chaine,"@$","@")
	index=-1
	For i=1 to Len(Chaine)
		c=asc(Mid(Chaine,i,1))
		If c<128 Then index=index+1
		If (c=9) or ((c>31) and (c<128)) Then
			If (c<>60) and (c<>62) and (c<>64) Then
				Chaine=Left(Chaine,i-1) & Mid(tDecode(c),Mid(Combinaison,(index mod 64)+1,1),1) & Mid(Chaine,i+1)
			End If
		End If
	Next
	Decode=Chaine
End Function
