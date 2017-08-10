'=======================================================================================================================================================
'Auteur : Brughes
'Note : vous pouvez écouter ou télécharger ma musique gratuitement sur : http://soundcloud.com/cyberflaneur ou http://www.jamendo.com/fr/artist/Brughes
'Description : incorporer un fichier .exe ou .com zippé ou non dans un vbscript en utilisant des boîtes de dialogue pour les fichiers d'entrée (exe et com) et de sortie (vbs).
'Le fichier vbscript créé contient également le code pour exécuter puis effacer le fichier .exe, .cpl ou .com encapsulé.
'=======================================================================================================================================================

Option Explicit

Dim WshShell, fso, objDialog, strFile, strComputer, Ret
strComputer = "."
Const ForWriting = 2
Const AdTypeBinary = 1

'Quitter si le script est déjà lancé
If AppPrevInstance() = True Then TerminateApp()
	
'Vérifier si l'objet ADODB.Stream est disponible
If Not IsRegistered("ADODB.Stream") Then
        MsgBox "ADODB n'est pas installé sur votre système." & vbcrlf & "Installez la dernière version de Microsoft Data Access Components.        ", vbOKOnly & vbInformation, "Incorporer un exécutable"
	WScript.Quit
End If

'Ouvrir une boîte de dialogue, pointant sur le bureau, pour incorporer l'exécutable
Set objDialog = CreateObject("UserAccounts.CommonDialog")
objDialog.Filter = "Programmes|*.exe;*.cpl;*.com"
objDialog.FilterIndex = 1
objDialog.Flags = 0
Set WshShell = WScript.CreateObject("WScript.Shell")
objDialog.InitialDir = WshShell.CurrentDirectory 'WshShell.SpecialFolders(0)
Set WshShell = Nothing
If objDialog.ShowOpen Then
        strFile = objDialog.FileName
Else
        TerminateApp()
End If

'Convertir strFile en nom court
Set fso = CreateObject("Scripting.FileSystemObject")
If Right(strFile, 1) = "\" Then strFile = Left(strFile, Len(strFile) - 1)
strFile = fso.GetAbsolutePathName(strFile)
If fso.FileExists(strFile) Then
	strFile = fso.GetFile(strFile).ShortPath
Else
        TerminateApp()
End If

Ret = MsgBox ("Voulez-vous zipper l'éxécutable ?", vbYesNoCancel Or vbSystemModal Or vbQuestion Or vbDefaultButton2, "Incorporer un exe dans un vbs")

If Ret = vbNo Then
        PutExeInVBS(strFile)
ElseIf Ret = vbYes Then
        PutZippedExeInVBS(strFile)
ElseIf Ret = VbAbort Then
        TerminateApp()
End If

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
        Set fso = Nothing
        Set objDialog = Nothing
        WScript.Quit 1
End Sub

Function IsRegistered(strObjectName)
'Renvoie True si l'objet peut être créé
        Dim obj
        On Error Resume Next
	Set obj = Nothing
	Set obj = CreateObject(strObjectName)
	If obj Is Nothing Then
		IsRegistered = False
	Else
		IsRegistered = True
		Set obj = Nothing
	End If
	Err.Clear
        On Error Goto 0
End Function

Sub Zip(InputFile, ZippedFile)
'Zip le fichier InputFile en une archive zip ZippedFile
	Dim oZip, fso, Shell
	Set Shell = CreateObject("Shell.Application")
        Set fso = CreateObject("Scripting.fileSystemObject")

        DelFile ZippedFile
        
	fso.CreateTextFile(ZippedFile, True).WriteLine "PK" & Chr(5) & Chr(6) & String(18, 0)
	Set oZip = Shell.NameSpace(ZippedFile)
	
	oZip.CopyHere InputFile

	Do Until oZip.Items.Count = 1
		Wscript.Sleep 100
	Loop
	
	Set Shell = Nothing
	Set oZip = Nothing
	Set fso = Nothing
End Sub

Function DelFile(ByVal FileIncludingPath)
'Efface le fichier FileIncludingPath (chemin_complet_du_fichier\nom_du_fichier.extension) même si il est en lecture seule
        Dim fso
	Const DeleteReadOnly = True
	Set fso = CreateObject("Scripting.fileSystemObject")

	On Error Resume Next
	Err.Clear

	If fso.FileExists(FileIncludingPath) Then
		fso.DeleteFile FileIncludingPath, DeleteReadOnly
		Do While fso.FileExists(FileIncludingPath) = True
		      Wscript.Sleep 100
		Loop
	End If

	DelFile = Err.Number
	Err.Clear
	On Error Goto 0
	
	Set fso = Nothing
End Function

Sub PutExeInVBS(ByVal strFile)
        Dim oStream, BinStream, nHexByte, i, FileOut, ScriptFileName
        
        'Lire strFile
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Type = AdTypeBinary
        oStream.Open
        oStream.LoadFromFile strFile
        BinStream = oStream.Read
        oStream.Close

        'Ouvrir une boîte de dialogue, pointant sur le bureau, pour enregistrer le fichier vbscript qui encapsulera l'exécutable
        Set objDialog = Nothing
        Set objDialog = CreateObject("SAFRCFileDlg.FileSave")
        Set WshShell = WScript.CreateObject("WScript.Shell")
        objDialog.FileName = WshShell.CurrentDirectory & "\*.vbs" 'objDialog.FileName = CreateObject("WScript.Shell").SpecialFolders(0) & "\*.vbs" ou 'objDialog.FileName = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\*.vbs"
        Set WshShell = Nothing
        objDialog.FileType = "VBScript (*.vbs)"
        If objDialog.OpenFileSaveDlg() Then
                ScriptFileName = objDialog.FileName
        Else
                TerminateApp()
        End If

        'Créer le fichier vbscript qui encapsulera l'exécutable
        Set FileOut = fso.OpenTextFile(ScriptFileName, ForWriting, True)

        'Ecrire le code de décodage dans le vbscript qui encapsulera l'exécutable
        FileOut.WriteLine "'Code qui incorpore un exécutable qui sera enregistré, exécuté et ensuite effacé"
        FileOut.WriteLine ""
        FileOut.WriteLine "Option Explicit"
        FileOut.WriteLine ""
        FileOut.WriteLine "Dim WsShell, fso, sLine, oFile, oRs, oStream, adFileName, strComputer"
        FileOut.WriteLine "strComputer = " & """."""
        FileOut.WriteLine ""
        FileOut.WriteLine "'Quitter si le script est déjà lancé"
        FileOut.WriteLine "If AppPrevInstance() = True Then WScript.Quit"
        FileOut.WriteLine ""
        FileOut.WriteLine "'Vérifie si l'objet ADODB.Stream est disponible"
        FileOut.WriteLine "If Not IsRegistered(" & """ADODB.Stream""" & ") Then"
        FileOut.WriteLine "                MsgBox " & """ADODB n'est pas installé sur votre système.""" & " & vbcrlf & " & """Installez la dernière version de Microsoft Data Access Components.        """ & ", vbOKOnly & vbInformation, " & """Incorporer un exécutable"""
        FileOut.WriteLine "                WScript.Quit"
        FileOut.WriteLine "End If"
        FileOut.WriteLine ""
        FileOut.WriteLine "'Déclaration des objets"
        FileOut.WriteLine "Set WsShell = CreateObject(" & """Wscript.Shell""" & ")"
        FileOut.WriteLine "Set fso = CreateObject(" & """Scripting.FileSystemObject""" & ")"
        FileOut.WriteLine "Set oFile = fso.OpenTextFile(WScript.ScriptFullName)"
        FileOut.WriteLine "Set oRs = CreateObject(" & """ADODB.RecordSet""" & ")"
        FileOut.WriteLine "Set oStream = CreateObject(" & """ADODB.Stream""" & ")"
        FileOut.WriteLine ""
        FileOut.WriteLine "'Déclaration des constantes"
        FileOut.WriteLine "Const adTypeBinary = 1"
        FileOut.WriteLine "Const adSaveCreateOverWrite = 2"
        FileOut.WriteLine "Const AdVarBinary = 204"
        FileOut.WriteLine "Const DeleteReadOnly = True"
        FileOut.WriteLine ""
        FileOut.WriteLine "'Nom du fichier exécutable qui sera enregistré"
        FileOut.WriteLine "adFileName = WsShell.ExpandEnvironmentStrings(" & """%TEMP%""" & ") & " & """" & "\" & fso.GetFile(strFile).Name & """"
        FileOut.WriteLine ""
        FileOut.WriteLine "oStream.Type = adTypeBinary"
        FileOut.WriteLine "oStream.Open"
        FileOut.WriteLine ""
        FileOut.WriteLine "oRs.Fields.Append " & """Data""" & ", adVarBinary, 32"
        FileOut.WriteLine "oRs.Open"
        FileOut.WriteLine "oRs.AddNew"
        FileOut.WriteLine ""
        FileOut.WriteLine "'Décode le fichier exécutable incorporé"
        FileOut.WriteLine "Do While Not oFile.AtEndOfStream"
        FileOut.WriteLine "        sLine = oFile.ReadLine"
        FileOut.WriteLine "        If Left(sLine, 3) = " & """'# """ & " Then"
        FileOut.WriteLine "                oRs(" & """Data""" & ") = Right(sLine, Len(sLine) - 3)"
        FileOut.WriteLine "                oRs.Update"
        FileOut.WriteLine "                oStream.Write oRs(" & """Data""" & ")"
        FileOut.WriteLine "        End If"
        FileOut.WriteLine "        Wscript.Sleep 0"
        FileOut.WriteLine "Loop"
        FileOut.WriteLine ""
        FileOut.WriteLine "'Enregistre le fichier exécutable incorporé sous le nom strFile"
        FileOut.WriteLine "oStream.SaveToFile adFileName, adSaveCreateOverWrite"
        FileOut.WriteLine ""
        FileOut.WriteLine "'Exécute l'exécutable puis l'efface"
        FileOut.WriteLine "If fso.FileExists(adFileName) Then"
        FileOut.WriteLine "        If fso.GetFile(adFileName).Size > 0 Then"
        FileOut.WriteLine "                WsShell.Run adFileName, 0, True"
        FileOut.WriteLine "                fso.DeleteFile adFileName, DeleteReadOnly"
        FileOut.WriteLine "        End If"
        FileOut.WriteLine "End If"
        FileOut.WriteLine ""
        FileOut.WriteLine "'Efface les objets en mémoire et quitte"
        FileOut.WriteLine "Set WsShell = Nothing"
        FileOut.WriteLine "Set fso = Nothing"
        FileOut.WriteLine "Set oFile = Nothing"
        FileOut.WriteLine "Set oRs = Nothing"
        FileOut.WriteLine "Set oStream = Nothing"
        FileOut.WriteLine "WScript.Quit"
        FileOut.WriteLine ""
        FileOut.WriteLine ""
        FileOut.WriteLine ""
        FileOut.WriteLine "'---------------------------------------------------------------------"
        FileOut.WriteLine "'Fonctions et procédures"
        FileOut.WriteLine "'---------------------------------------------------------------------"
        FileOut.WriteLine ""
        FileOut.WriteLine "Function AppPrevInstance()"
        FileOut.WriteLine "'Vérifie si un script portant le même nom que le présent script est déjà lancé"
        FileOut.WriteLine "        Dim objWMIService, colScript, objScript, RunningScriptName, Counter"
        FileOut.WriteLine "        Counter = 0"
        FileOut.WriteLine "        Set objWMIService = GetObject(" & """winmgmts:""" & " & " & """{impersonationLevel=impersonate}!\\""" & " & strComputer & " & """\root\cimv2""" & ")"
        FileOut.WriteLine "        Set colScript = ObjWMIService.ExecQuery(" & """SELECT * FROM Win32_Process WHERE Name = 'Wscript.exe' OR Name = 'Cscript.exe'""" & ")"
        FileOut.WriteLine "        For Each objScript In colScript"
        FileOut.WriteLine "                RunningScriptName = Mid(objScript.CommandLine, InstrRev(objScript.CommandLine, " & """\""" & ", -1, 1) + 1, Len(objScript.CommandLine) - InstrRev(objScript.CommandLine, " & """\""" & ", -1, 1) - 2)"
        FileOut.WriteLine "                If WScript.ScriptName = RunningScriptName Then Counter = Counter + 1"
        FileOut.WriteLine "                Wscript.Sleep 100"
        FileOut.WriteLine "        Next"
        FileOut.WriteLine "        If  Counter > 1 Then"
        FileOut.WriteLine "                AppPrevInstance = True"
        FileOut.WriteLine "        Else"
        FileOut.WriteLine "                AppPrevInstance = False"
        FileOut.WriteLine "        End If"
        FileOut.WriteLine "        Set ColScript = Nothing"
        FileOut.WriteLine "        Set ObjWMIService = Nothing"
        FileOut.WriteLine "End Function"
        FileOut.WriteLine ""
        FileOut.WriteLine "Function IsRegistered(strObjectName)"
        FileOut.WriteLine "'Renvoie True si l'objet peut être créé"
        FileOut.WriteLine "        Dim obj"
        FileOut.WriteLine "        On Error Resume Next"
        FileOut.WriteLine "        Set obj = Nothing"
        FileOut.WriteLine "        Set obj = CreateObject(strObjectName)"
        FileOut.WriteLine "        If obj Is Nothing Then"
        FileOut.WriteLine "                IsRegistered = False"
        FileOut.WriteLine "        Else"
        FileOut.WriteLine "                IsRegistered = True"
        FileOut.WriteLine "                Set obj = Nothing"
        FileOut.WriteLine "        End If"
        FileOut.WriteLine "        Err.Clear"
        FileOut.WriteLine "        On Error Goto 0"
        FileOut.WriteLine "End Function"
        FileOut.WriteLine ""
        FileOut.Write "'Fichier exécutable incorporé"

        'Encoder et écrire l'exécutable dans le vbscript qui encapsulera l'exécutable
        For i = 0 To Lenb(BinStream) - 1
                If i Mod 32 = 0 Then FileOut.Write vbcrlf & "'# "
                nHexByte = Right("0" & Hex(Ascb(Midb(BinStream, i + 1, 1))), 2)
                FileOut.Write nHexByte
                Wscript.Sleep 0
        Next

        'Rajouter un saut de ligne à la fin du vbscript qui encapsulera l'exécutable
        FileOut.WriteLine vbcrlf & ""

        'Fermer le fichier vbscript
        FileOut.Close

        Set FileOut = Nothing
        Set oStream = Nothing
End Sub

Sub PutZippedExeInVBS(ByVal strFile)
        Dim oStream, BinStream, nHexByte, i, FileOut, ScriptFileName, ZipFile
        
        'Zipper l'exécutable dans Program.zip
        Set WshShell = CreateObject("WScript.Shell")
        ZipFile = WshShell.ExpandEnvironmentStrings("%TEMP%") & "\Program.zip"
        Zip strFile, ZipFile
        Set WshShell = Nothing

        'Lire Program.zip
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Type = AdTypeBinary
        oStream.Open
        oStream.LoadFromFile ZipFile
        BinStream = oStream.Read
        oStream.Close

        'Ouvrir une boîte de dialogue, pointant sur le bureau, pour enregistrer le fichier vbscript qui encapsulera l'exécutable zippé
        Set objDialog = Nothing
        Set objDialog = CreateObject("SAFRCFileDlg.FileSave")
        Set WshShell = WScript.CreateObject("WScript.Shell")
        objDialog.FileName = WshShell.CurrentDirectory & "\*.vbs" 'objDialog.FileName = CreateObject("WScript.Shell").SpecialFolders(0) & "\*.vbs"
        Set WshShell = Nothing
        objDialog.FileType = "VBScript (*.vbs)"
        If objDialog.OpenFileSaveDlg() Then
        	ScriptFileName = objDialog.FileName
        Else
                TerminateApp()
        End If

        'Créer le fichier vbscript qui encapsulera l'exécutable zippé
        Set FileOut = fso.OpenTextFile(ScriptFileName, ForWriting, True)

        'Ecrire le code de décodage dans le vbscript qui encapsulera l'exécutable zippé
        FileOut.WriteLine "'Code qui incorpore un exécutable zippé qui sera enregistré, exécuté et ensuite effacé"
        FileOut.WriteLine ""
        FileOut.WriteLine "Option Explicit"
        FileOut.WriteLine ""
        FileOut.WriteLine "Dim WsShell, fso, sLine, oFile, oRs, oStream, adFileName, strComputer"
        FileOut.WriteLine "strComputer = " & """."""
        FileOut.WriteLine ""
        FileOut.WriteLine "'Quitter si le script est déjà lancé"
        FileOut.WriteLine "If AppPrevInstance() = True Then WScript.Quit"
        FileOut.WriteLine ""
        FileOut.WriteLine "'Vérifie si l'objet ADODB.Stream est disponible"
        FileOut.WriteLine "If Not IsRegistered(" & """ADODB.Stream""" & ") Then"
        FileOut.WriteLine "                MsgBox " & """ADODB n'est pas installé sur votre système.""" & " & vbcrlf & " & """Installez la dernière version de Microsoft Data Access Components.        """ & ", vbOKOnly & vbInformation, " & """Incorporer un exécutable"""
        FileOut.WriteLine "                WScript.Quit"
        FileOut.WriteLine "End If"
        FileOut.WriteLine ""
        FileOut.WriteLine "'Déclaration des objets"
        FileOut.WriteLine "Set WsShell = CreateObject(" & """Wscript.Shell""" & ")"
        FileOut.WriteLine "Set fso = CreateObject(" & """Scripting.FileSystemObject""" & ")"
        FileOut.WriteLine "Set oFile = fso.OpenTextFile(WScript.ScriptFullName)"
        FileOut.WriteLine "Set oRs = CreateObject(" & """ADODB.RecordSet""" & ")"
        FileOut.WriteLine "Set oStream = CreateObject(" & """ADODB.Stream""" & ")"
        FileOut.WriteLine ""
        FileOut.WriteLine "'Déclaration des constantes"
        FileOut.WriteLine "Const adTypeBinary = 1"
        FileOut.WriteLine "Const adSaveCreateOverWrite = 2"
        FileOut.WriteLine "Const AdVarBinary = 204"
        FileOut.WriteLine "Const DeleteReadOnly = True"
        FileOut.WriteLine ""
        FileOut.WriteLine "'Nom du fichier Zip qui sera enregistré"
        FileOut.WriteLine "adFileName = WsShell.ExpandEnvironmentStrings(" & """%TEMP%""" & ") & " & """\Program.zip"""
        FileOut.WriteLine ""
        FileOut.WriteLine "oStream.Type = adTypeBinary"
        FileOut.WriteLine "oStream.Open"
        FileOut.WriteLine ""
        FileOut.WriteLine "oRs.Fields.Append " & """Data""" & ", adVarBinary, 32"
        FileOut.WriteLine "oRs.Open"
        FileOut.WriteLine "oRs.AddNew"
        FileOut.WriteLine ""
        FileOut.WriteLine "'Décode le fichier zip incorporé"
        FileOut.WriteLine "Do While Not oFile.AtEndOfStream"
        FileOut.WriteLine "        sLine = oFile.ReadLine"
        FileOut.WriteLine "        If Left(sLine, 3) = " & """'# """ & " Then"
        FileOut.WriteLine "                oRs(" & """Data""" & ") = Right(sLine, Len(sLine) - 3)"
        FileOut.WriteLine "                oRs.Update"
        FileOut.WriteLine "                oStream.Write oRs(" & """Data""" & ")"
        FileOut.WriteLine "        End If"
        FileOut.WriteLine "        Wscript.Sleep 0"
        FileOut.WriteLine "Loop"
        FileOut.WriteLine ""
        FileOut.WriteLine "'Enregistre le fichier zip incorporé sous le nom Program.zip"
        FileOut.WriteLine "oStream.SaveToFile adFileName, adSaveCreateOverWrite"
        FileOut.WriteLine ""
        FileOut.WriteLine "'Extrait l'exécutable du zip"
        FileOut.WriteLine "Unzip adFileName, WsShell.ExpandEnvironmentStrings(" & """%TEMP%""" & ") & " & """\""" & ", " & """" & fso.GetFile(strFile).Name & """"
        FileOut.WriteLine ""
        FileOut.WriteLine "'Efface Program.zip"
        FileOut.WriteLine "adFileName = WsShell.ExpandEnvironmentStrings(" & """%TEMP%""" & ") & " & """\Program.zip"""
        FileOut.WriteLine "If fso.FileExists(adFileName) Then"
        FileOut.WriteLine "        If fso.GetFile(adFileName).Size > 0 Then"
        FileOut.WriteLine "                'Efface Program.zip"
        FileOut.WriteLine "                fso.DeleteFile adFileName, DeleteReadOnly"
        FileOut.WriteLine "        End If"
        FileOut.WriteLine "End If"
        FileOut.WriteLine ""
        FileOut.WriteLine "'Efface le répertoire temporaire " & """Répertoire temporaire 1 pour Program.zip"""
        FileOut.WriteLine "adFileName = WsShell.ExpandEnvironmentStrings(" & """%TEMP%""" & ") & " & """\Répertoire temporaire 1 pour Program.zip"""
        FileOut.WriteLine "If fso.FolderExists(adFileName) Then"
        FileOut.WriteLine "        'Efface l'exécutable du " & """Répertoire temporaire 1 pour Program.zip"""
        FileOut.WriteLine "        If fso.FileExists(adFileName & " & """\""" & " & " & """" & fso.GetFile(strFile).Name & """" & ") Then fso.DeleteFile adFileName & " & """\""" & " & " & """" & fso.GetFile(strFile).Name & """" & ", DeleteReadOnly"
        FileOut.WriteLine "        'Efface " & """Répertoire temporaire 1 pour Program.zip"""
        FileOut.WriteLine "        fso.DeleteFolder adFileName"
        FileOut.WriteLine "End If"
        FileOut.WriteLine ""
        FileOut.WriteLine "'Exécute l'exécutable puis l'efface"
        FileOut.WriteLine "adFileName = WsShell.ExpandEnvironmentStrings(" & """%TEMP%""" & ") & " & """\""" & " & " & """" & fso.GetFile(strFile).Name & """"
        FileOut.WriteLine "If fso.FileExists(adFileName) Then"
        FileOut.WriteLine "        If fso.GetFile(adFileName).Size > 0 Then"
        FileOut.WriteLine "                WsShell.Run adFileName, 0, True"
        FileOut.WriteLine "                fso.DeleteFile adFileName, DeleteReadOnly"
        FileOut.WriteLine "        End If"
        FileOut.WriteLine "End If"
        FileOut.WriteLine ""
        FileOut.WriteLine "'Efface les objets en mémoire et quitte"
        FileOut.WriteLine "Set WsShell = Nothing"
        FileOut.WriteLine "Set fso = Nothing"
        FileOut.WriteLine "Set oFile = Nothing"
        FileOut.WriteLine "Set oRs = Nothing"
        FileOut.WriteLine "Set oStream = Nothing"
        FileOut.WriteLine "WScript.Quit"
        FileOut.WriteLine ""
        FileOut.WriteLine "Function AppPrevInstance()"
        FileOut.WriteLine "'Vérifie si un script portant le même nom que le présent script est déjà lancé"
        FileOut.WriteLine "        Dim objWMIService, colScript, objScript, RunningScriptName, Counter"
        FileOut.WriteLine "        Counter = 0"
        FileOut.WriteLine "        Set objWMIService = GetObject(" & """winmgmts:""" & " & " & """{impersonationLevel=impersonate}!\\""" & " & strComputer & " & """\root\cimv2""" & ")"
        FileOut.WriteLine "        Set ColScript = objWMIService.ExecQuery(" & """SELECT * FROM Win32_Process WHERE Name = 'Wscript.exe' OR Name = 'Cscript.exe'""" & ")"
        FileOut.WriteLine "        For Each objScript In colScript"
        FileOut.WriteLine "                RunningScriptName = Mid(objScript.CommandLine, InstrRev(objScript.CommandLine, " & """\""" & ", -1, 1) + 1, Len(objScript.CommandLine) - InstrRev(objScript.CommandLine, " & """\""" & ", -1, 1) - 2)"
        FileOut.WriteLine "                If WScript.ScriptName = RunningScriptName Then Counter = Counter + 1"
        FileOut.WriteLine "                Wscript.Sleep 100"
        FileOut.WriteLine "        Next"
        FileOut.WriteLine "        If  Counter > 1 Then"
        FileOut.WriteLine "                AppPrevInstance = True"
        FileOut.WriteLine "        Else"
        FileOut.WriteLine "                AppPrevInstance = False"
        FileOut.WriteLine "        End If"
        FileOut.WriteLine "        Set colScript = Nothing"
        FileOut.WriteLine "        Set objWMIService = Nothing"
        FileOut.WriteLine "End Function"
        FileOut.WriteLine ""
        FileOut.WriteLine "Function IsRegistered(strObjectName)"
        FileOut.WriteLine "'Renvoie True si l'objet peut être créé"
        FileOut.WriteLine "        Dim obj"
        FileOut.WriteLine "        On Error Resume Next"
        FileOut.WriteLine "        Set obj = Nothing"
        FileOut.WriteLine "        Set obj = CreateObject(strObjectName)"
        FileOut.WriteLine "        If obj Is Nothing Then"
        FileOut.WriteLine "                IsRegistered = False"
        FileOut.WriteLine "        Else"
        FileOut.WriteLine "                IsRegistered = True"
        FileOut.WriteLine "                Set obj = Nothing"
        FileOut.WriteLine "        End If"
        FileOut.WriteLine "        Err.Clear"
        FileOut.WriteLine "        On Error Goto 0"
        FileOut.WriteLine "End Function"
        FileOut.WriteLine ""
        FileOut.WriteLine "Sub Unzip(ZippedFile, UnZippedFile, ExistingUnZippedFile)"
        FileOut.WriteLine "'Dezippe le fichier ZippedFile vers l'emplacement UnZippedFile et efface le fichier ExistingUnZippedFile"
        FileOut.WriteLine "        Dim FilesInZip, Shell"
        FileOut.WriteLine "        DelFile UnZippedFile & ExistingUnZippedFile"
        FileOut.WriteLine "        Set Shell = CreateObject(" & """Shell.Application""" & ")"
        FileOut.WriteLine "        Set FilesInZip = Shell.NameSpace(ZippedFile).Items"
        FileOut.WriteLine "        Shell.NameSpace(UnZippedFile).CopyHere FilesInZip,(4 + 8 + 16 + 512 + 1024)"
        FileOut.WriteLine "        Do"
        FileOut.WriteLine "                If fso.FileExists(UnZippedFile & ExistingUnZippedFile) Then"
        FileOut.WriteLine "                        If fso.GetFile(UnZippedFile & ExistingUnZippedFile).Size > 0 Then Exit Do"
        FileOut.WriteLine "                End If"
        FileOut.WriteLine "                Wscript.Sleep 1000"
        FileOut.WriteLine "        Loop"
        FileOut.WriteLine "        Set FilesInZip = Nothing"
        FileOut.WriteLine "        Set Shell = Nothing"
        FileOut.WriteLine "End Sub"
        FileOut.WriteLine ""
        FileOut.WriteLine "Function DelFile(ByVal FileIncludingPath)"
        FileOut.WriteLine "'Efface le fichier FileIncludingPath (chemin_complet_du_fichier\nom_du_fichier.extension) même si il est en lecture seule"
        FileOut.WriteLine "     Dim fso"
        FileOut.WriteLine "	Const DeleteReadOnly = True"
        FileOut.WriteLine "	Set fso = CreateObject(" & """Scripting.fileSystemObject""" & ")"
        FileOut.WriteLine "	On Error Resume Next"
        FileOut.WriteLine "	Err.Clear"
        FileOut.WriteLine "	If fso.FileExists(FileIncludingPath) Then"
        FileOut.WriteLine "		fso.DeleteFile FileIncludingPath, DeleteReadOnly"
        FileOut.WriteLine "		Do While fso.FileExists(FileIncludingPath) = True"
        FileOut.WriteLine "		      Wscript.Sleep 100"
        FileOut.WriteLine "		Loop"
        FileOut.WriteLine "	End If"
        FileOut.WriteLine "	DelFile = Err.Number"
        FileOut.WriteLine "	Err.Clear"
        FileOut.WriteLine "	On Error Goto 0"	
        FileOut.WriteLine "	Set fso = Nothing"
        FileOut.WriteLine "End Function"
        FileOut.WriteLine ""
        FileOut.Write "'Fichier exécutable zippé incorporé"

        'Encoder et écrire l'exécutable zippé dans le vbscript qui encapsulera l'exécutable
        For i = 0 To Lenb(BinStream) - 1
                If i Mod 32 = 0 Then FileOut.Write vbcrlf & "'# "
                nHexByte = Right("0" & Hex(Ascb(Midb(BinStream, i + 1, 1))), 2)
                FileOut.Write NHexByte
                Wscript.Sleep 0
        Next

        'Rajouter un saut de ligne à la fin du vbscript qui encapsulera l'exécutable
        FileOut.WriteLine vbcrlf & ""

        'Fermer le fichier vbscript
        FileOut.Close

        'Effacer Program.zip
        DelFile ZipFile

        Set FileOut = Nothing
        Set oStream = Nothing
End Sub

