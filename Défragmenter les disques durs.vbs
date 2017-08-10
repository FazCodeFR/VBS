'=============================================================================================
'Description :
'Défragmente en priorité basse tous les disques durs sans que rien ne soit visible et sans intervention de l'utilisateur.
'Une défragmentation poussée est obtenue par itérations.
'Le code calcul, par différentiation, le temps écoulé entre deux défragmentations successives.
'Ceci rend le code indépendant de la vitesse d'exécution du PC et des performances des disques durs.
'Optimise également le Prefetch.
'Ceci permet de défragmenter régulièrement en lançant ce script par le planificateur des tâches.
'
'Commentaires :
'Il faut qu'il y ait au moins 15% d'espace libre pour réaliser la défragmentation dans de bonnes conditions.
'Cet espace est utilisé comme zone de transit pour les fichiers déplacés.
'En deça de ce seuil, la défragmentation risque d'être longue et incomplète.
'Les lecteurs à mémoire Flash ne doivent jamais être défragmentés car leur nombre d'écriture est limité.
'
'Microsoft conseille de défragmenter un disque dur si le taux de fragmentation est supérieur à 10%.
'Un taux de fragmentation de moins de 10 % est en principe pas gênant. Cependant, les performances des disques sont notablement moindres !
'Depuis Windows Vista, la fragmentation atteint rarement un seuil gênant car un utilitaire inclus défragmente automatiquement, par défaut, tous les mercredis à 1 heure du matin.
'Sous XP cet fonctionnalité n'est pas implémentée. Il est donc intéressant de connaître le taux de fragmentation des disques et de défragmenter régulièrement.
'
'Auteur : Brughes
'Vous pouvez écouter/télécharger ma musique en open source : http://soundcloud.com/cyberflaneur ou http://www.jamendo.com/fr/artist/Brughes
'=============================================================================================

Option Explicit

Const DriveTypeFixed = 2
Dim WshShell, fso, d, dc, Return, X, I, Time0, Time1, ElapsedSecond, ElapsedSecond1, ElapsedSecond2, Counter

Set WshShell = WScript.CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
Set dc = fso.Drives

'Défragmente les disques durs
X = 0
For Each d in dc
	If d.DriveType = DriveTypeFixed Then
		X = X + 1 : Counter = 0 : ElapsedSecond1 = 0 : ElapsedSecond2 = 0

		'Boucle d'itérations jusqu'à ce que le temps de défragmentation des disques soit inférieur à 10 secondes avec une limite de 8 itérations
		Do While Counter < 8
			Counter = Counter + 1

			'Mémorise le temps initial converti en secondes
			Time0 = Now : Time0 = Second(Time0) + Minute(Time0) * 60 + Hour(Time0) * 3600

			'Lance Defrag de manière invisible
			RunDefrag d, "-f"

			'Si une erreur est survenue quitter la boucle pour passer à la défragmentation du disque suivant
			If Return <> 0 Then Exit Do

			'Mémorise le temps final converti en secondes
			Time1 = Now : Time1 = Second(Time1) + Minute(Time1) * 60 + Hour(Time1) * 3600

			'Calcul du temp écoulé en secondes
			If Time1 >= Time0 Then
				ElapsedSecond = Time1 - Time0
			Else
				'Détecter le passage à minuit
				ElapsedSecond = 24 * 3600 - Time0 + Time1
			End If

			'Calcul de la différence de temps des deux dernières défragmentations
			If ElapsedSecond1 <> 0 Then
				ElapsedSecond2 = ElapsedSecond 'temps écoulé le plus récent
			Else
				ElapsedSecond1 = ElapsedSecond 'temps écoulé le plus ancien
			End If

			If ElapsedSecond2 <> 0 And ElapsedSecond1 <> 0 Then
				ElapsedSecond = ElapsedSecond1 - ElapsedSecond2  'temps écoulé le plus ancien - temps écoulé le plus récent
				ElapsedSecond1 = ElapsedSecond2
				ElapsedSecond2 = 0
			End If

			If Abs(ElapsedSecond) < 10 Then Exit Do

			Wscript.Sleep 100
		Loop
	End If
	Wscript.Sleep 100
Next

'Défragmente le boot si ce n'est pas réalisé par la défragmentation classique
DefragBoot

'Force la mise à jour de Layout.ini utilisé par le prefetch et la défragmentation du boot
UpDateLayout

'Supprimer les objets en mémoire et quitter
Set dc = Nothing
Set fso = Nothing
Set WshShell = Nothing

WScript.Quit


Function IsProcessRunning(ByVal strProcessName)
'Détecter si strProcessName est en cours d'exécution

	Const strComputer = "."
	Dim objWMIService, objProcess, colProcess

	Set objWMIService = GetObject("winmgmts:" &"{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colProcess = objWMIService.ExecQuery ("Select * from Win32_Process")

	IsProcessRunning = False
	For Each objProcess In colProcess
		If objProcess.Name = strProcessName Then IsProcessRunning = True
		Wscript.Sleep 100
	Next

	Set objWMIService = Nothing
	Set colProcess = Nothing

End Function

Function WaitUntilProcessEnds(ByVal strProcessName)
'Boucle d'attente si strProcessName est en cours d'exécution

	Do While IsProcessRunning(strProcessName) = True
		Wscript.Sleep 100
	Loop

	WaitUntilProcessEnds = True

End Function

Sub RunDefrag(ByVal strDrive, ByVal strOption)
'Crée un process Defrag en priorité basse sur l'ordinateur local

	Dim objWMIService, objConfig, objStartup, objProcess, intProcessID, strCommand, Return
	Const SW_HIDE = 0
	Const strComputer = "."
	Const IDLE_PROCESS = 64

	Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set objStartup = objWMIService.Get("Win32_ProcessStartup")

	'Crée la chaîne de commande dfrgntfs.exe
	strCommand = "dfrgntfs.exe -Embedding"

	'Configure dfrgntfs.exe en priorité basse avec une fenêtre cachée
	Set objConfig = objStartup.SpawnInstance_
	objConfig.ShowWindow = SW_HIDE
	objConfig.PriorityClass = IDLE_PROCESS

	'Crée le process dfrgntfs.exe
	Set objProcess = objWMIService.Get("Win32_Process")
	Return = objProcess.Create(strCommand, Null, objConfig, intProcessID)

	'Attendre que le process dfrgntfs.exe soit lancé
	IsProcessRunning("dfrgntfs.exe")

	'Crée la chaîne de commande defrag.exe pour le lecteur strDrive avec le paramètre strOption
	strCommand = "defrag.exe " & strDrive & " " & Trim(strOption)

	'Configure defrag.exe en priorité basse avec une fenêtre cachée
	Set objConfig = objStartup.SpawnInstance_
	objConfig.ShowWindow = SW_HIDE
	objConfig.PriorityClass = IDLE_PROCESS

	'Crée le process defrag.exe
	Set objProcess = objWMIService.Get("Win32_Process")
	Return = objProcess.Create(strCommand, Null, objConfig, intProcessID)

	'Attendre que le process defrag.exe soit lancé
	IsProcessRunning("defrag.exe")

	'Attendre que le process defrag.exe soit arrêté
	WaitUntilProcessEnds("defrag.exe")

	Set objWMIService = Nothing
	Set objStartup = Nothing
	Set objConfig = Nothing
	Set objProcess = Nothing

End Sub

Sub DefragBoot()
'Une défragmentation normale défragmente le boot si la valeur Enable = Y pour clé HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Dfrg\BootOptimizeFunction
'Ceci force donc la défragmentation du boot si la valeur Enable = N pour clé HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Dfrg\BootOptimizeFunction

	On Error Resume Next
	Err.Clear

	'Vérifie si la défragmentation automatique du boot est désactivée
	If WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Dfrg\BootOptimizeFunction\Enable") <> "Y" Then
		If Err.Number = 0 Then

			'Détecter si defrag.exe est en cours d'exécution avant de poursuivre
			WaitUntilProcessEnds("defrag.exe")

			'Défragmente le Boot
			RunDefrag WshShell.ExpandEnvironmentStrings("%HOMEDRIVE%"), "-b"

		End If
	Else
		Err.Clear
	End If

	On Error Goto 0

End Sub

Sub UpDateLayout()
'Lance le processus d'optimisation du disque et également les processus qui sont lancés lorsque la machine est inactive si la sous clé "Enable" de la clé HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Dfrg\BootOptimizeFunction a la valeur Y.
'Il ne s'agit pas d'une défragmentation complète mais d'une optimisation de la zone de boot pour accélérer le temps de démarrage et le temps d'accès au disque.
'Ceci force la mise à jour de Layout.ini utilisé par le prefetch à condition que Layout.ini ait été créé et la défragmentation immédiate du boot.

	On Error Resume Next
	Err.Clear

	'Vérifie si Layout a été créé
	If WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Dfrg\BootOptimizeFunction\OptimizeComplete") = "Yes" Then
		If Err.Number = 0 Then

			'Vérifie si le processus d'optimisation est désactivé. Il est inutile de le lancer si il est déjà activé car Windows le lancera automatiquement.
			If WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Dfrg\BootOptimizeFunction\Enable") <> "Y" Then
				If Err.Number = 0 Then

					'Détecter si Defrag.exe est en cours d'exécution avant de poursuivre
					WaitUntilProcessEnds("defrag.exe")

					'La défragmentation du boot s'effectue tous les 3 jours et nécessite un redémarrage. Ceci force la mise à jour de Layout.ini utilisé par le prefetch. Sinon, il faut attendre 3 jours et redémarrer la machine.
					Return = WshShell.Run("Rundll32.exe advapi32.dll,ProcessIdleTasks", 0, True)

				Else
					Err.Clear
				End If
			End If
		Else
			Err.Clear
		End If
	End If

	On Error Goto 0

End Sub
