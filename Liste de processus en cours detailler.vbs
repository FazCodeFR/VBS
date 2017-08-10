'Les éléments à démarrage automatique + ListProcessCmdLine.vbs © Hackoo © 2011
Set fso = CreateObject("Scripting.FileSystemObject")
Set Ws = CreateObject("WScript.Shell")
Set ProcessEnv = Ws.Environment("Process")
NomMachine = ProcessEnv("COMPUTERNAME") 
NomUtilisateur = ProcessEnv("USERNAME") 
NomFichierLog="Liste_Processus.txt"
temp = Ws.ExpandEnvironmentStrings("%temp%")
PathNomFichierLog = temp & "\" & NomFichierLog
Set OutPut = fso.CreateTextFile(temp & "\" & NomFichierLog,2)
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" _ 
& strComputer & "\root\cimv2") 
Set colProcesses = objWMIService.ExecQuery ("Select * from Win32_Process")
count=0 
'Dim StartTime : StartTime = Timer
Call Infosys
OutPut.WriteLine "[COLOR=""Red""]" & String(14,"*")& "Liste des Processus en cours d'exécution le " & date & " à " & time & " sur Le PC "& NomMachine &" connecté en tant que " & NomUtilisateur & String(14,"*")& vbNewline & String(80,"*") & "[/COLOR]"
For Each objProcess in colProcesses
ProcessName = objProcess.Name
ProcessID = objProcess.ProcessID
CommandLine = objProcess.CommandLine    
count=count+1
Texte = "Numéro PID = "& objProcess.ProcessID & VbNewLine & "Nom du Processus = " & objProcess.Name & VbNewLine &"Ligne de Commande = "& objProcess.CommandLine &_
VbNewLine & String(100,"*")
OutPut.WriteLine Texte
Next
 
OutPut.WriteLine  "Il y a "& Count &" Processus en cours d'exécution le " & date & " à " & time & " sur Le PC "& NomMachine &" connecté en tant que " & NomUtilisateur & vbNewline
Call StartupCommand
'OutPut.WriteLine FormatNumber(Timer - StartTime, 0) & " seconds."
Function StartupCommand()
strComputer = "."
resultat=""
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colStartupCommands = objWMIService.ExecQuery ("Select * from Win32_StartupCommand")
 
For Each objStartupCommand in colStartupCommands
resultat=resultat & "Nom: " & objStartupCommand.Name & vbNewline
resultat=resultat & "Description: " & objStartupCommand.Description & vbNewline
resultat=resultat & "Emplacement: " & objStartupCommand.Location & vbNewline
resultat=resultat & "Commande: " & objStartupCommand.Command & vbNewline
resultat=resultat & "Utilisateur: " & objStartupCommand.User & vbNewline
resultat=resultat & String(100,"*") & vbNewline 
Next
OutPut.WriteLine "[COLOR=""Red""]" & String(50,"*") &" Les éléments à démarrage automatique "& String(40,"*") &"[/COLOR]"
OutPut.WriteLine resultat & "[/quote]"
end Function
 
Explorer(PathNomFichierLog)
 
Function Explorer(File)
    Set ws=CreateObject("wscript.shell")
    ws.run "Notepad "& File,1,True
end Function
 
Function InfoSys
strComputer = "."
strMessage=""
Set objWMIService = GetObject("winmgmts:"  & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colSettings = objWMIService.ExecQuery  ("Select * from Win32_ComputerSystem")
Set colSettings2 = objWMIService.ExecQuery ("Select * from Win32_BIOS")
Set colSettings3 = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem")
For Each objBIOS in colSettings2 
      strMessage=strMessage & "[quote]BIOS " & objBIOS.Version & vbNewline & vbNewline
Next
For Each objComputer in colSettings 
      strMessage=strMessage & "Nom de l'ordinateur : " & objComputer.Name & vbNewline & "Fabriquant: " & objComputer.Manufacturer & vbNewline & "Modèle : " & objComputer.Model & vbNewline & vbNewline
 
Next
For Each objOperatingSystem in colSettings3
      strMessage=strMessage &  objOperatingSystem.Name & vbNewline
      strMessage=strMessage &  "Version " & objOperatingSystem.Version & vbNewline
      strMessage=strMessage &  "Service Pack " & objOperatingSystem.ServicePackMajorVersion & "." & objOperatingSystem.ServicePackMinorVersion &vbNewline
      strMessage=strMessage &  "Dossier de Windows: " & objOperatingSystem.WindowsDirectory &vbNewline
Next
OutPut.WriteLine strMessage & ""
end Function