Option Explicit
Dim File,MyRootFolder,RootFolder
MyRootFolder = Browse4Folder
Call Scan4File(MyRootFolder)
MsgBox "Script fini !"  & vbcr & "Tout les fichiers du dossier sont passer !" & vbcr & "Merci a hackoofr",VbInformation,"Opération fini !"
'**************************************************************************
Function Browse4Folder()
    Dim objShell,objFolder,Message
    Message = "Please select a folder in order to scan into it and its subfolders to rename files"
    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.BrowseForFolder(0,Message,0,0)
    If objFolder Is Nothing Then
        Wscript.Quit
    End If
    Browse4Folder = objFolder.self.path
End Function
'****************************************************************************
Function Scan4File(Folder)
    Dim fso,objFolder,arrSubfolders,FileName
    Dim Tab,SubFolder,NewFileName,aFile,partfic
    Dim oShell,iRet
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set objFolder = fso.GetFolder(Folder)
    Set arrSubfolders = objFolder.SubFolders
    For Each FileName in objFolder.Files
        Tab = Split(FileName,".")
        NewFileName = Tab(0) & " " & Tab(1) & "." & Tab(UBound(Tab))
        Set aFile = fso.GetFile(FileName)
        partfic = split(FileName,".")
        Set oShell = WScript.CreateObject("WScript.Shell")
        iRet = oShell.Popup ("Le fichier : " & vbcr & FileName & vbcr & vbcr & " Va devenir : " & vbcr & NewFileName,8,"Oui ou Non ?",vbOKCancel+32)
         If iRet = 1 or -1 Then
           MsgBox "Ok ou attend"
           aFile.Move NewFileName
           For Each SubFolder in arrSubfolders
            'Call Scan4File(SubFolder) 'appel récursive pour scanner dans les sous-dossiers
         End If
         If iRet = 2 Then 
           oShell.Popup "Fichier Annuler",4," /!\ ERREUR /!\",48
           Next
         End If
Next
End Function
'**************************************************************************