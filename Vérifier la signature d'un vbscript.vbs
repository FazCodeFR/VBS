'====================================================================================================================================================================================
'Auteur : Brughes.
'Note : vous pouvez écouter/télécharger ma musique en open source : http://soundcloud.com/cyberflaneur ou http://www.jamendo.com/fr/artist/Brughes
'Description : vérifie la signature d'un script en ouvrant une boîte de dialogue pointant sur le script.
'====================================================================================================================================================================================

Option Explicit

Dim WshShell, objDialog, strFile, strComputer
strComputer = "."

'Quitter si le script est déjà lancé
If AppPrevInstance() = True Then TerminateApp()

'Ouvrir une boîte de dialogue pointant sur le script à signer
Set objDialog = CreateObject("UserAccounts.CommonDialog")
objDialog.Filter = "Vbscript|*.vbs;*.vbe"
objDialog.FilterIndex = 1
objDialog.Flags = 0
Set WshShell = WScript.CreateObject("WScript.Shell")
objDialog.InitialDir = WshShell.CurrentDirectory 'WshShell.SpecialFolders("Desktop")
If objDialog.ShowOpen Then
        strFile = objDialog.FileName
Else
        TerminateApp()
End If

VerifyIfScriptIsTrustable strFile

'Effacer les objets en mémoire et quitter
TerminateApp()


'====================================================================================================================================================================================
'Fonctions et procédures
'====================================================================================================================================================================================

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
        Set WshShell = Nothing
        Set objDialog = Nothing
        WScript.Quit
End Sub

Sub VerifyIfScriptIsTrustable(ByVal strScriptFile)
'vérifie la signature du script strScriptFile
        Dim objSigner, blnShowGUI, blnIsSigned

        'True pour afficher l'interface utilisateur pour les scripts non signés ou non fiables
        blnShowGUI = False 'True
        
        Set objSigner = WScript.CreateObject("Scripting.Signer")
        blnIsSigned = objSigner.VerifyFile(strScriptFile, blnShowGUI)

        If blnIsSigned = True Then
                MsgBox "    Ce vbscript est fiable.             ", 64, "Signature d'un vbscript"
        Else
                MsgBox "    Ce vbscript n'est pas fiable.             ", 64, "Signature d'un vbscript"
        End If

        Set objSigner = Nothing
End Sub

