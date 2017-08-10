Set oShell = CreateObject("wscript.Shell")
Set env = oShell.environment("Process")
strComputer = env.Item("Computername")
Const HKEY_LOCAL_MACHINE = &H80000002
Const UnInstPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_
".\root\default:StdRegProv")
 
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem",,48)
 
report = report & "******************************************" & vbCrLf
report = report & " - Information  de l'ordinateur " & strComputer & " - " & vbCrLf
report = report & "******************************************" & vbCrLf & vbCrLf
 
 
report = report & vbCrLf & "******************************************" & vbCrLf
report = report & "Informations sur Windows" & vbCrLf & "******************************************" & vbCrLf
 
For Each objItem in colItems
    report = report &  "- Nom du poste: " & strComputer  & vbCrLf
    report = report &  "- Description de l'ordinateur: " & objItem.Description & vbCrLf
    report = report &  "- Utilisateur possédant la licence Windows: " & objItem.RegisteredUser & vbCrLf
    report = report &  "- Organisation possédant la licence Windows: " & objItem.Organization & vbCrLf
    report = report & "******************************************" & vbCrLf
    report = report &  "- Nom du système d'exploitation: " & objItem.Caption & vbCrLf
    If (objItem.OSProductSuite <> "")Then
        report = report &  "- Système d'exploitation de la suite " & objItem.OSProductSuite & vbCrLf
    End If
    report = report &  "- Version: " & objItem.Version & vbCrLf
    report = report &  "- Date de son installation: " & objItem.InstallDate & vbCrLf
    report = report &  "- Numéro de série de " & objItem.Caption & ": " & objItem.SerialNumber & vbCrLf
    report = report & vbCrLf
    report = report & "******************************************" & vbCrLf
    report = report & "Détails techniques sur Windows"& vbCrlf
    report = report & "******************************************" & vbCrLf
    report = report &  "- Numéro du dernier Service Pack majeur installé: "
    report = report & objItem.ServicePackMajorVersion & vbCrLf
 
    If (objItem.MaxNumberOfProcesses="-1") Then
        report = report &  "- Maximum de processus pouvant être ouvert: Aucune limite fixée" & vbCrLf
    Else
        report = report &  "- Maximum de processus pouvant être ouvert: " & objItem.MaxNumberOfProcesses & vbCrLf
    End If
Next
 
Set colSettings = objWMIService.ExecQuery _
("Select * from Win32_ComputerSystem")
report = report & "******************************************" & vbCrLf
report = report & "Mémoire vive (RAM) et processeur" & vbCrLf & "******************************************" & vbCrLf
For Each objComputer in colSettings
'report = report & objComputer.Name & vbcrlf
    report = report & "- Vous avez actuellement " & objComputer.TotalPhysicalMemory /1024\1024+1 & " Mo de mémoire vive(RAM) au total." & vbcrlf
Next
 
Set colSettings = objWMIService.ExecQuery _
("Select * from Win32_Processor")
For Each objProcessor in colSettings
 
    report = report & "- Type de processeur: "
    If objProcessor.Architecture = 0 Then
        report = report & "x86" & vbCrLf
    ElseIf objProcessor.Architecture = 1 Then
        report = report & "MIPS" & vbCrLf
    ElseIf objProcessor.Architecture = 2 Then
        report = report & "Alpha" & vbCrLf
    ElseIf objProcessor.Architecture = 3 Then
        report = report & "PowerPC" & vbCrLf
    ElseIf objProcessor.Architecture = 6 Then
        report = report & "ia64" & vbCrLf
    Else
        report = report & "inconnu" & vbCrLf
    End If
 
    report = report & "- Nom du processeur: " & objProcessor.Name & vbCrLf
    report = report & "- Description du processeur: " & objProcessor.Description & vbCrLf
    report = report & "- Vitesse actuelle du processeur: " & objProcessor.CurrentClockSpeed & " Mhz" & vbCrLf
    report = report & "- Vitesse maximale du processeur: " & objProcessor.MaxClockSpeed & " Mhz" & vbCrLf
 
    report = report & vbCrLf
Next
 
report = report & "******************************************" & vbCrLf
report = report & "Disque(s) dur(s) et autres lecteurs actuellement " & vbCrLf
report = report & "en usage" & vbCrLf & "******************************************" & vbCrLf
 
Dim oFSO
Set oFSO = WScript.CreateObject("Scripting.FileSystemObject")
 
Dim oDesLecteurs
Set oDesLecteurs = oFSO.Drives
 
Dim oUnLecteur
Dim strLectType
 
For Each oUnLecteur in oDesLecteurs
    If oUnLecteur.IsReady Then
        Select Case oUnLecteur.DriveType
        Case 0: strLectType = "Inconnu"
        Case 1: strLectType = "Amovible (Disquette, clé USB, etc.)"
        Case 2: strLectType = "Fixe (Disque dur, etc.)"
        Case 3: strLectType = "Réseau"
        Case 4: strLectType = "CD-Rom"
 
        End Select
 
        report = report & "- Lettre du lecteur: " & oUnLecteur.DriveLetter & vbCrLf
        report = report & "- Numéro de série: " & oUnLecteur.SerialNumber & vbCrLf
'             'report = report & "- Type de lecteur: " & oUnLecteur.strLectType & vbCrLf
        If (oUnLecteur.FileSystem <> "") Then
            report = report & "- Système de fichier utilisé: " & oUnLecteur.FileSystem & vbCrLf
        End If
 
        Set objWMIService = GetObject("winmgmts:")
        Set objLogicalDisk = objWMIService.Get("Win32_LogicalDisk.DeviceID='" & oUnLecteur.DriveLetter & ":'")
 
    End If
    report = report & vbCrLf
Next
 
srComputer = "."
Set objWMIService = GetObject("winmgmts:" & "!\\" & srComputer & "\root\cimv2")
Set colAdapters = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled = True") 
For Each objAdapter in colAdapters
    IPdebut = LBound(objAdapter.IPAddress)
    IPfin = UBound(objAdapter.IPAddress)
    If (objAdapter.IPAddress(IPdebut) <> "") then
        For i = IPdebut To IPfin
            msg =  msg  & "utilise l'adresse IP " & objAdapter.IPAddress(i) & vbCrLf
        Next     
    End If
Next
 
Set fso = CreateObject("Scripting.FileSystemObject")
'Détermine si le fichier texte existe déjà ou s'il doit le créer
If Not fso.FileExists("inventaire" & strComputer & ".txt") Then
    set ts = fso.CreateTextFile("inventaire_" & strComputer & ".txt", True)
Else
    set ts = fso.OpenTextFile("inventaire_" & strComputer & ".txt", 2, True)
End If
'Wscript.Echo msg
ts.write report
ts.write software
ts.write msg
Set ws = CreateObject("wscript.shell")
ws.run "notepad inventaire_" & strComputer & ".txt"