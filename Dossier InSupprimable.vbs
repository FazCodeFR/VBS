'*** Choix du créé ou supprimer  ***
choix = InputBox ("Voulez vous créer un dossier ou le supprimer ?" & vbCr & vbCr & vbCr & "1 = Créer" &vbCr & "2 = Supprimer","Créer ou Supprimer un dossier ?","1") 

'*** Choix du répertoire ***
Const RETURNONLYFSDIRS = &H1 
  
Set oShell = CreateObject("Shell.Application") 
Set oFolder = oShell.BrowseForFolder(&H0&, "Choisir un répertoire", RETURNONLYFSDIRS, "c:\") 
If oFolder is Nothing Then  
	MsgBox "Aucun dossier choissi !",vbCritical 
Else 
  Set oFolderItem = oFolder.Self
  nomdudossier = oFolderItem.path 
End If 
Set oFolderItem = Nothing 
Set oFolder = Nothing 
Set oShell = Nothing

If choix = "1" Then
'Choix créer

'*** Exercution code cmd ***
Set WS = CreateObject("WScript.Shell")
Command = "cmd /C md " & nomdudossier & "\con\"
Result = Ws.Run(Command,0,True)

If Result = "0" Then 
MsgBox "La création ou la suppression c'est exercuté avec succès.", vbInformation, "SUCCES"
Else
MsgBox "Erreur fatal lors de création ou la suppression du dossier.", vbError, "ERREUR"
End If 

ElseIf choix = "2" Then
'Choix supprimer

'*** Exercution code cmd ***
Set WS = CreateObject("WScript.Shell")
Command = "cmd /C rd " & nomdudossier & "\"
Result = Ws.Run(Command,0,True)

If Result = "0" Then 
MsgBox "La création ou la suppression c'est exercuté avec succès.", vbInformation, "SUCCES"
Else
MsgBox "Erreur fatal lors de création ou la suppression du dossier.", vbError, "ERREUR"
End If 

Else
'Autre choix 
MsgBox "Demande incomprise, l'application va quitter"
WScript.Quit ()
End If 