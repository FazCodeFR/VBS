'==========================================================================================================================================
'Supprime tous les points de restauration et en crée un nouveau sans modifier la configuration.
'Fonctionne sur un système à un seul disque dur ou à plusieurs disques durs.
'ATTENTION : ce script efface TOUS les points de restauration.
'L'activation ou la désactivation de la restauration pour chaque lecteur est identifiée et non modifiée.
'
'Auteur : Brughes
'Vous pouvez écouter/télécharger ma musique en open source : http://soundcloud.com/cyberflaneur ou http://www.jamendo.com/fr/artist/Brughes
'==========================================================================================================================================
Option Explicit

Dim WshShell, objWMIService, objItem, oSRP, colItems, fso, dc, d, Drive(), RestoreState(), RestorePointCount, errResults, Counter, i

Const cRestorePointType = 0
Const cEventType = 100
Const strComputer = "."

'Quitter si le script est déjà lancé
If AppPrevInstance() = True Then TerminateApp()

Set WshShell = WScript.CreateObject("WScript.Shell")

'Si la stratégie de restauration est désactivée, quitter.
On Error Resume Next
If WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore\DisableSR") = 1 Then
	If Err.Number = 0 Then
		Set WshShell = Nothing
		On Error Goto 0
		WScript.Quit
	Else
		Set WshShell = Nothing
		Err.Clear
		On Error Goto 0
		WScript.Quit
	End If
End If

'La stratégie de restauration est activée pour le lecteur principal car la clé DisableSR a été filtrée.
'On peut donc traiter les points de restauration existants.

Set fso = CreateObject("Scripting.FileSystemObject")
Set dc = fso.Drives
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\default")
Set objItem = objWMIService.Get("SystemRestore")

'Parcours des disques durs autres que la partition principale.
Counter = 0
For Each d in dc
	If d.DriveType = 2 And d <> WshShell.ExpandEnvironmentStrings("%homedrive%") Then
			Counter = Counter + 1
			ReDim Preserve RestoreState(Counter)
			ReDim Preserve Drive(Counter)
			
			'Mémoriser la lettre du lecteur secondaire.
			Drive(Counter) = d
			
			'Tester si la partition du lecteur secondaire est activée en l'activant.
			errResults = objItem.Enable(d & "\")
			
			'Si on a tenté d'activer une restauration déjà active pour le lecteur secondaire.
                     	  	If errResults = 1056 Then
				'Mémoriser l'activation de la restauration du lecteur secondaire.
				RestoreState(Counter) = True
			End If

			'Si on a tenté d'activer une restauration inactive pour le lecteur secondaire.
			If errResults = 0 Then
				'Mémoriser l'inactivation de la restauration du lecteur secondaire.
				RestoreState(Counter) = False
			End If
	End If
	Wscript.Sleep 100
Next

'Arrête le service de restauration et supprime tous les points de restauration.
errResults = objItem.Disable("")

'Attente de la suppression des points de restauration.
Set colItems = objWMIService.ExecQuery("Select * from SystemRestore")
Do While colItems.Count > 0
	WScript.Sleep 100
Loop

'Faire une pause pour permettre la finalisation de la suppression des points de restauration.
WScript.Sleep 5000

If Counter = 0 Then
	'Si il n'y a pas de partition secondaire, active uniquement la restauration de la partition principale.
	errResults = objItem.Enable(WshShell.ExpandEnvironmentStrings("%homedrive%") & "\")
Else
	'Active la restauration de la partition principale en ajoutant l'option True pour permettre de désactiver la restauration des partitions secondaires.
	errResults = objItem.Enable(WshShell.ExpandEnvironmentStrings("%homedrive%") & "\", True)

	'Rétabli l'état de la restauration des lecteurs secondaires pour lesquels la restauration était désactivée.
	For i = 1 To Counter
		If RestoreState(i) = False Then
			errResults = objItem.Disable(Drive(i) & "\")
                        'RestorePointCount = 0
			'Attente de la création d'un point de vérification système.
			Set colItems = objWMIService.ExecQuery("Select * from SystemRestore")
			'Do While colItems.Count = RestorePointCount
			'	WScript.Sleep 100
			'Loop			
		End If
		WScript.Sleep 100		
	Next
	Set colItems = Nothing
End If

'Efface les tableaux.
Erase Drive
Erase RestoreState

'Crée un point de restauration nommé "Point de Restauration généré par script".
Set oSRP = GetObject("winmgmts:\\" & strComputer & "\root\default:SystemRestore")
Set colItems = objWMIService.ExecQuery("Select * from SystemRestore")
RestorePointCount = colItems.Count
errResults = oSRP.CreateRestorePoint("Point de restauration généré par script", cRestorePointType, cEventType)

'Attente de la création du "Point de restauration généré par script".
Set colItems = objWMIService.ExecQuery("Select * from SystemRestore")
Do While colItems.Count < RestorePointCount + 1
	WScript.Sleep 100
Loop

'Faire une pause pour permettre la finalisation du point de restauration nommé "Point de Restauration généré par script" avant de quitter.
WScript.Sleep 5000

'Effacer les objets en mémoire et quitter
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
        Set objItem = Nothing
        Set colItems = Nothing
        Set objWMIService = Nothing
        Set dc = Nothing
        Set fso = Nothing
        Set oSRP = Nothing
        Set WshShell = Nothing
        WScript.Quit
End Sub

