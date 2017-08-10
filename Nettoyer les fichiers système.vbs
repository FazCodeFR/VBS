'============================================================================================
'Nettoie récursivement les fichiers système (valable pour XP).
'Les fichiers en cours d'utilisation ne sont pas effacés.
'Dans ce script, seuls l'effacement des répertoires C:\Temp et C:\Documents and Settings\Utilisateur\Local Settings\temp est activé.
'Vous pouvez réactiver les lignes de code désactivées.
'Les autres répertoires sont donnés à titre indicatif.
'ATTENTION : vérifier toujours le chemin du répertoire et son contenu avant de décider d'effacer.
'
'Auteur : Brughes
'Vous pouvez écouter/télécharger ma musique en open source : http://soundcloud.com/cyberflaneur ou http://www.jamendo.com/fr/artist/Brughes
'============================================================================================

Option Explicit

Dim WshShell, fso, Return, strComputer

strComputer = "."

'Quitter si le script est déjà lancé.
If AppPrevInstance() = True Then TerminateApp()

Set WshShell = CreateObject("Wscript.Shell")
Set fso = WScript.CreateObject("Scripting.FileSystemObject")

'Supprimer les fichiers temporaires système
Return = DelTree(WshShell.ExpandEnvironmentStrings("%TEMP%"))

'Supprimer les fichiers temporaires de Windows
Return = DelTree(WshShell.ExpandEnvironmentStrings("%systemroot%") & "\Temp")

'Supprimer les fichiers temporaires de User Application Data
Return = DelTree(WshShell.ExpandEnvironmentStrings("%USERPROFILE%") & "\Local Settings\Application Data\Temp")

'Supprimer les fichiers et dossiers de Temp
Return = DelTree(WshShell.ExpandEnvironmentStrings("%homedrive%") & "\Temp")

'Supprimer les fichiers de C:\WINDOWS\PCHealth\HelpCtr\Temp
Return = DelTree(WshShell.ExpandEnvironmentStrings("%systemroot%") & "\PCHealth\HelpCtr\Temp")

'Supprimer les fichiers récents
Return = DelFileWithExt(WshShell.SpecialFolders("Recent"), "lnk")

'Supprimer les installations téléchargées User (effacées avec IE)
Return = DelTree(WshShell.ExpandEnvironmentStrings("%systemroot%") & "\Downloaded Installations")

'Supprimer les installations téléchargées de All User
Return = DelTree(WshShell.ExpandEnvironmentStrings("%ALLUSERSPROFILE%") & "\Application Data\Downloaded Installations")

'Supprimer l'historique des applications
Return = DelTree(WshShell.ExpandEnvironmentStrings("%USERPROFILE%") & "\Local Settings\Application Data\ApplicationHistory")

'Supprimer les fichiers AntiPhishing de IE
Return = DelTree(WshShell.ExpandEnvironmentStrings("%USERPROFILE%") & "\Local Settings\Temporary Internet Files\AntiPhishing")

'Supprimer le cache Feeds
Return = DelTree(WshShell.ExpandEnvironmentStrings("%USERPROFILE%") & "\Local Settings\Application Data\Microsoft\Feeds Cache")

'Supprimer les fichiers .log de C:\WINDOWS
Return = DelFileWithExt(WshShell.ExpandEnvironmentStrings("%systemroot%"), "log")

'Supprimer les fichiers .log de C:\WINDOWS\PCHealth\HelpCtr\Logs
Return = DelFileWithExt(WshShell.ExpandEnvironmentStrings("%systemroot%") & "\PCHealth\HelpCtr", "log")

'Supprimer les fichiers .log de C:\WINDOWS\Logs
Return = DelFileWithExt(WshShell.ExpandEnvironmentStrings("%systemroot%") & "\Logs", "log")

'Supprimer les fichiers .log de C:\WINDOWS\Debug
Return = DelFileWithExt(WshShell.ExpandEnvironmentStrings("%systemroot%") & "\Debug", "log")

'Supprimer les fichiers .log de C:\WINDOWS\Debug\UserMode
Return = DelFileWithExt(WshShell.ExpandEnvironmentStrings("%systemroot%") & "\Debug\UserMode", "log")

'Supprimer les fichiers .log de C:\WINDOWS\Debug\WPD
Return = DelFileWithExt(WshShell.ExpandEnvironmentStrings("%systemroot%") & "\Debug\WPD", "log")

'Supprimer les fichiers .log de C:\WINDOWS\Debug\Setup
Return = DelFileWithExt(WshShell.ExpandEnvironmentStrings("%systemroot%") & "\Debug\Setup", "log")

'Supprimer les fichiers .log de wbem. GetSpecialFolderPath(37) donne le chemin de System32
Return = DelTree(GetSpecialFolderPath(37) & "\wbem\Logs")

'Supprimer les installations Office
ClearMSOfficeInstallationFiles()

'Supprimer les rapports de ZoneAlarm
Return = DelTree(WshShell.ExpandEnvironmentStrings("%systemroot%") & "\Internet Logs")

'Supprimer l'historique des sites de Flash Player
'Racine = WshShell.ExpandEnvironmentStrings("%USERPROFILE%") & "\Application Data\Macromedia\Flash Player\#SharedObjects\9MG28Y2U"
DeleteFlashPlayerHistoryPath

'Supprimer le cache de Flash Player
'Racine = WshShell.ExpandEnvironmentStrings("%APPDATA%") & "\Adobe\Flash Player\AssetCache\39JV6SUD"
DeleteFlashPlayerCachePath

'Supprimer l'historique de Flash Player
Return = DelTree(WshShell.ExpandEnvironmentStrings("%APPDATA%") & "\Macromedia\Flash Player\macromedia.com\support\flashplayer\sys")

'Supprimer le cache de Mozilla FireFox
'Racine = WshShell.ExpandEnvironmentStrings("%USERPROFILE%") & "\Local Settings\Application Data\Mozilla\Firefox\Profiles\42cezepf.default\Cache"
DeleteMozillaFireFoxCachePath

'Supprimer le cache download.sqlite des téléchargements de Mozilla FireFox
DeleteMozillaFireFoxDownloadCachePath()

'Supprimer les rapports d'erreurs de Mozilla FireFox
Return = DelTree(WshShell.ExpandEnvironmentStrings("%APPDATA%") & "\Mozilla\Firefox\Crash Reports\pending")

'Supprimer les rapports de  Spybot - Search & Destroy
Return = DelTree(WshShell.ExpandEnvironmentStrings("%ALLUSERSPROFILE%") & "\Application Data\Spybot - Search & Destroy\Logs")

'Supprimer les anciens fichiers de uTorrent
Return = DelFileWithExt(WshShell.ExpandEnvironmentStrings("%APPDATA%") & "\utorrent", "old")

'Supprimer le cache des images de uTorrent
Return = DelTree(WshShell.ExpandEnvironmentStrings("%APPDATA%") & "\utorrent\dlimagecache")

'Supprimer le cache de Google Updater
Return = DelTree(WshShell.ExpandEnvironmentStrings("%ALLUSERSPROFILE%") & "\Application Data\Google Updater\cache")

'Supprimer les objets en mémoire et quitter
TerminateApp()



'=======================================================================================================
'Fonctions et procédures.
'=======================================================================================================

Sub TerminateApp()
'Supprime les objets en mémoire et quitte
        Set fso = Nothing
        Set WshShell = Nothing
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

Function GetSpecialFolderPath(ByVal FolderNamespace)
'Renvoie le chemin du SpecialFolder désigné par FolderNamespace
        Dim objShell, objFolder, objFolderItem
        Set objShell = CreateObject("Shell.Application")
        Set objFolder = objShell.Namespace(FolderNamespace)
        Set objFolderItem = objFolder.Self
        GetSpecialFolderPath = objFolderItem.Path
        Set objShell = Nothing
        Set objFolder = Nothing
        Set objFolderItem = Nothing
End Function

Function DelTree(ByVal StrPath)
'Efface récursivement tous les fichiers et tous les sous-répertoires contenus dans le répertoire strPath
	Dim TheFiles, AFile, Afolder, SubFolder

	On Error Resume Next
	Err.Clear

	If fso.FolderExists(strPath) Then
		Set Afolder = fso.GetFolder(strPath)
		Set TheFiles = Afolder.Files

		For Each AFile In TheFiles
			AFile.Delete
			Wscript.Sleep 100
		Next

		For Each SubFolder In Afolder.SubFolders
			DelTree(SubFolder)
			SubFolder.Delete
			Wscript.Sleep 100
		Next

		Set TheFiles = Nothing
		Set Afolder = Nothing
	End If

	DelTree = Err.Number

	Err.Clear
	On Error Goto 0
End Function

Function DelFile(ByVal FileIncludingPath)
'Efface le fichier FileIncludingPath (chemin_complet_du_fichier\nom_du_fichier.extension) même si il est en lecture seule
	Const DeleteReadOnly = True

	On Error Resume Next
	Err.Clear

	If fso.FileExists(FileIncludingPath) Then
		fso.DeleteFile FileIncludingPath, DeleteReadOnly
	End If

	DelFile = Err.Number

	Err.Clear
	On Error Goto 0
End Function

Function DelFileWithExt(ByVal FilePath, ByVal strExtention)
'Efface tous les fichiers ayant l'extension strExtention contenus dans le répertoire FilePath y compris les fichiers en lecture seule
	Const DeleteReadOnly = True

	On Error Resume Next
	Err.Clear

	If fso.FileExists(FilePath & "\*." & strExtention) Then
		fso.DeleteFile FilePath & "\*." & strExtention, DeleteReadOnly
	End If

	DelFileWithExt = Err.Number

	Err.Clear
	On Error Goto 0
End Function

Sub ClearMSOfficeInstallationFiles()
'Supprimer les installations Office
	Const DriveTypeFixed = 2
	Dim dc, d

	Set dc = fso.Drives

	On Error Resume Next
	Err.Clear

	For Each d In dc
		If d.DriveType = DriveTypeFixed Then
			If fso.FolderExists(d & "\msdownld.tmp") Then DelTree(d & "\msdownld.tmp")
		End If
		Wscript.Sleep 100
	Next

	On Error Goto 0
	Err.Clear
	
	Set dc = Nothing
End Sub

Sub DeleteMozillaFireFoxCachePath()
'Trouver le chemin du cache de Mozilla FireFox et effacer le cache
	Dim Afolder, SubFolder, DefaultSubDir, Racine, MozillaFireFoxCachePath

	Racine = WshShell.ExpandEnvironmentStrings("%USERPROFILE%") & "\Local Settings\Application Data\Mozilla\Firefox\Profiles"
	
	If fso.FolderExists(Racine) Then
		Set Afolder = fso.GetFolder(Racine)
		For Each SubFolder In Afolder.SubFolders
			If InstrRev(SubFolder, ".default", -1, 1) <> 0 Then
				DefaultSubDir = Mid(SubFolder, InstrRev(SubFolder, "\", -1, 1) + 1 , InstrRev(SubFolder, ".default", -1, 1) - (InstrRev(SubFolder, "\", -1, 1) + 1))
			End If
			Wscript.Sleep 100
		Next
		Set Afolder = Nothing
		MozillaFireFoxCachePath = Racine & "\" & DefaultSubDir & ".default\Cache"
		If fso.FolderExists(MozillaFireFoxCachePath) Then DelTree(MozillaFireFoxCachePath)
	End If
End Sub

Sub DeleteMozillaFireFoxDownloadCachePath()
'Trouver le chemin du cache de Mozilla FireFox et effacer le cache des téléchargements de Mozilla FireFox
	Dim Afolder, SubFolder, DefaultSubDir, Racine, MozillaFireFoxDownloadCachePath

	Racine = WshShell.ExpandEnvironmentStrings("%USERPROFILE%") & "\Local Settings\Application Data\Mozilla\Firefox\Profiles"

	If fso.FolderExists(Racine) Then
		Set Afolder = fso.GetFolder(Racine)
		For Each SubFolder In Afolder.SubFolders
			If InstrRev(SubFolder, ".default", -1, 1) <> 0 Then
				DefaultSubDir = Mid(SubFolder, InstrRev(SubFolder, "\", -1, 1) + 1 , InstrRev(SubFolder, ".default", -1, 1) - (InstrRev(SubFolder, "\", -1, 1) + 1))
			End If
			Wscript.Sleep 100
		Next
		Set Afolder = Nothing
		MozillaFireFoxDownloadCachePath = Racine & "\" & DefaultSubDir & ".default\downloads.sqlite"
		If fso.FileExists(MozillaFireFoxDownloadCachePath) Then DelFile(MozillaFireFoxDownloadCachePath)
	End If
End Sub

Sub DeleteFlashPlayerHistoryPath()
'Trouver le chemin de l'historique des sites de Flash Player et effacer l'historique
	Dim Afolder, SubFolder, DefaultSubDir, Racine, FlashPlayerHistoryPath

	Racine = WshShell.ExpandEnvironmentStrings("%USERPROFILE%") & "\Application Data\Macromedia\Flash Player\#SharedObjects"

	If fso.FolderExists(Racine) Then
		Set Afolder = fso.GetFolder(Racine)
		For Each SubFolder In Afolder.SubFolders
			DefaultSubDir = Mid(SubFolder, InstrRev(SubFolder, "\", -1, 1) + 1 , Len(SubFolder) - Len(Racine))
			Wscript.Sleep 100
		Next
		Set Afolder = Nothing
		FlashPlayerHistoryPath = Racine & "\" & DefaultSubDir
		If fso.FolderExists(FlashPlayerHistoryPath) Then DelTree(FlashPlayerHistoryPath)
	End If
End Sub

Sub DeleteFlashPlayerCachePath()
'Trouver le chemin du cache de Flash Player et effacer le cache
	Dim Afolder, SubFolder, DefaultSubDir, Racine, FlashPlayerCachePath
	
        Racine = WshShell.ExpandEnvironmentStrings("%APPDATA%") & "\Adobe\Flash Player\AssetCache"
	
        If Fso.FolderExists(Racine) Then
		Set Afolder = fso.GetFolder(Racine)
		For Each SubFolder In Afolder.SubFolders
			DefaultSubDir = Mid(SubFolder, InstrRev(SubFolder, "\", -1, 1) + 1, Len(SubFolder) - Len(Racine))
			Wscript.Sleep 100
		Next
		Set Afolder = Nothing
		FlashPlayerCachePath = Racine & "\" & DefaultSubDir
		If fso.FolderExists(FlashPlayerCachePath) Then DelTree(FlashPlayerCachePath)
	End If
End Sub

