Set dtmConvertedDate = CreateObject("WbemScripting.SWbemDateTime")
strComputer = "."
strResultat=""
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colOperatingSystems = objWMIService.ExecQuery _
    ("Select * from Win32_OperatingSystem")
For Each objOperatingSystem in colOperatingSystems
    strResultat=strResultat &  objOperatingSystem.Caption & vbCR
    strResultat=strResultat & "Version: " & objOperatingSystem.Version & vbCR
    strResultat=strResultat & "Numéro de série: " & objOperatingSystem.SerialNumber & vbCR
    strResultat=strResultat & "Langue: " & objOperatingSystem.OSLanguage & vbCR
    strResultat=strResultat &  "Code pays: " & objOperatingSystem.CountryCode & vbCR
    strResultat=strResultat & "Utilisateur enregistré: " & objOperatingSystem.RegisteredUser & vbCR
    strResultat=strResultat & "Organisation: " & objOperatingSystem.Organization & vbCR
    dtmConvertedDate.Value = objOperatingSystem.InstallDate 
    dtmInstallDate = dtmConvertedDate.GetVarDate
    strResultat=strResultat & "Date installation: " & dtmInstallDate & vbCR
Next
MsgBox strResultat,64,"OSLang"