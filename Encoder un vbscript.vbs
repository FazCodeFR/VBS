'=======================================================================================================================================================
'Auteur : Brughes
'Note : vous pouvez écouter ou télécharger ma musique gratuitement sur : http://soundcloud.com/cyberflaneur ou http://www.jamendo.com/fr/artist/Brughes
'Description : Encode les fichiers vbs en utilisant des boîtes de dialogue pour les fichiers d'entrée (vbs) et de sortie (vbe).
'=======================================================================================================================================================

Option Explicit

Dim WshShell, fso, objDialog, Srce, Dest, strComputer, Ret, fs, pe, ext, data, se, dataenc, fd
Const ForReading = 1

strComputer = "."

'Quitter si le script est déjà lancé
If AppPrevInstance() = True Then TerminateApp()

'Ouvrir une boîte de dialogue, pointant sur le répertoire courant, pour lire le script à encoder
Set objDialog = CreateObject("UserAccounts.CommonDialog")
objDialog.Filter = "VBScript (*.vbs)|*.vbs"
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
objDialog.FileName = WshShell.CurrentDirectory & "\*.vbe"
Set WshShell = Nothing
objDialog.FileType = "VBScript (*.vbe)"
If objDialog.OpenFileSaveDlg() Then
        Dest = objDialog.FileName
        If Right(Dest, 1) = "\" Then Dest = Left(Dest, Len(Dest) - 1)
	Set fso = CreateObject("Scripting.FileSystemObject")
        Dest = fso.GetAbsolutePathName(Dest)
Else
        TerminateApp()
End If

If Dest = "" Then TerminateApp()

'Encodage
Set fs = fso.GetFile(Srce)
pe = InstrRev(fs.Name,".")
If pe <> 0 Then ext = LCase(Mid(fs.Name,pe))
If ext <> ".vbs" Then TerminateApp()

Set fs = fso.OpenTextFile(Srce, ForReading, False)
data = fs.ReadAll
WScript.Sleep 100
fs.Close
WScript.Sleep 100
Set se = WScript.CreateObject("Scripting.Encoder")
WScript.Sleep 100
dataenc = se.EncodeScriptFile(ext, data, 0, "")
WScript.Sleep 100

Set fd = fso.CreateTextFile(Dest, True, False)
WScript.Sleep 100
fd.Write dataenc
WScript.Sleep 100
fd.Close
WScript.Sleep 100

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
        Set objDialog = Nothing
	Set WshShell = Nothing
        Set fs = Nothing
        Set se = Nothing
        Set fd = Nothing
        WScript.Quit
End Sub
