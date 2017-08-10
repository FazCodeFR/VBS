Dim FileEntrer, FileSorti, oFSO 
Set scrEnc = CreateObject("Scripting.Encoder")
Set scrFSO = CreateObject("Scripting.FileSystemObject")

    FileEntrer = Inputbox ("Le nom du fichier vbs a encoder","By ABOAT")
    If scrFSO.FileExists(FileEntrer & ".bat") Then 
    MsgBox "Le script existe bien !",Vbinformation,"Coder par ABOAT !"
    FileSorti = Inputbox ("Le nom du futur fichier","By ABOAT")
    myfile = scrFSO.OpenTextFile(FileEntrer & ".bat").ReadAll
    If scrFSO.FileExists("" & FileSorti & ".vbe") Then scrFSO.DeleteFile "" & FileSorti & ".vbe", True
    myFileEncode=scrENC.EncodeScriptFile(".bat", myfile, 0, "")
 
    Set ts = scrFSO.CreateTextFile("" & FileSorti & ".vbe", True, False)
    ts.Write myFileEncode 
    ts.Close
    MsgBox "Operation fini ! Le fichier vbs d'origine: " & FileEntrer & " et devenue : " & FileSorti,Vbinformation,"Encodeur de VBS"
    
    Else
      MsgBox "Le fichier n'existe pas ! ",vbCritical
    End If