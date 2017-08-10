'====================================================================================================================================================================================
'Efface les fichiers stockés sur le disque par Internet Explorer
'Auteur : Brughes
'
'Vous pouvez écouter/télécharger ma musique en open source : http://soundcloud.com/cyberflaneur ou http://www.jamendo.com/fr/artist/Brughes
'====================================================================================================================================================================================

Option Explicit

Dim WsShell, Param, strComputer

strComputer = "."

'Quitter si le script est déjà lancé.
If AppPrevInstance() = True Then TerminateApp()

Set WsShell = CreateObject("Wscript.Shell")

Param = 0

'Effacer l'historique
Param = Param + 1

'Effacer les Cookies
Param = Param + 2

'Effacer les fichiers Internet temporaires
Param = Param + 8

'Effacer les données des formulaires
'Param = Param + 16

'Effacer les mots de passe
'Param = Param + 32

'Effacer l'historique de navigation y compris l'historique des compléments
Param = Param + 193

'Effacer complètement l'historique de navigation
Param = Param + 255

'Effacer le cheminement
'Param = Param + 2048

'Tout effacer y compris les fichiers et les paramètres des compléments
'Param = Param + 4351

'Préserver les favoris
'Param = Param + 8192

'Effacer les fichiers téléchargés (downloaded Files)
Param = Param + 16384

'Tout effacer
'Param = Param + 22783

On Error Resume Next

WsShell.Run "Rundll32.exe InetCpl.cpl,ClearMyTracksByProcess " & Param, 0, True

On Error Goto 0

'Supprimer les objets en mémoire et quitter
TerminateApp()


'====================================================================================================================================================================================
'Fonctions et procédures.
'====================================================================================================================================================================================

Sub TerminateApp()
'Supprime les objets en mémoire et quitte
        Set WsShell = Nothing
	WScript.Quit
End Sub

Function AppPrevInstance()
'Vérifie si un script portant le même nom que le présent script est déjà lancé
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

        'Efface les objets en mémoire
        Set colScript = Nothing
        Set objWMIService = Nothing
End Function

