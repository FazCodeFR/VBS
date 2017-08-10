'Vider la corbeille

Option Explicit

Dim fso, Shell, Folder, FolderItem, WsShell, strComputer

strComputer = "."

'Quitter si le script est déjà lancé
If AppPrevInstance() = True Then TerminateApp()

Set fso = CreateObject("Scripting.FileSystemObject")
Set Shell = CreateObject("Shell.Application")
Set WsShell = CreateObject("Wscript.Shell")

On Error Resume Next
Err.Clear

'Création d'un objet Corbeille
Set Folder = Shell.NameSpace(10)

'Si la Corbeille est disponible
If Err.Number = 0 Then

	'Tester si la corbeille est pleine
	If Folder.Items.Count <> 0 Then

		'Vider la corbeille
		For Each FolderItem In Folder.Items
			If fso.FileExists(FolderItem.Path) Then
				fso.DeleteFile FolderItem.Path
			ElseIf Fso.FolderExists(FolderItem.Path) Then
				fso.DeleteFolder FolderItem.Path
			End If
			Wscript.Sleep 100
		Next

		'Rafraichir l'icône de la corbeille
		WsShell.Run "Rundll32.exe shell32.dll,SHUpdateRecycleBinIcon", 0, True

	End If

	Set Folder = Nothing
Else
	Err.Clear
End If

On Error Goto 0

'Supprimer les objets en mémoire et quitter
TerminateApp()


'==================================================================================================
'Fonctions et procédures
'==================================================================================================

Sub TerminateApp()
'Effacer les objets en mémoire et quitter
        Set WsShell = Nothing
        Set Shell = Nothing
        Set fso = Nothing

        WScript.Quit
End Sub

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


