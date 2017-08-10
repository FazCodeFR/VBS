Option Explicit
'************************************************************************************** 
' GetFileDlg() And GetFileDlgBar() by omen999 - may 2014 - http://omen999.developpez.com
' Universal Browse for files function  
' compatibility : all versions windows and IE - supports start folder, filters and title
' note : the global size of the parameters cannot exceed 191 chars for GetFileDlg 
'**************************************************************************************
Dim Title,sIniDir,sFilter,sTitle,InFile,WshShell,fso,URL,ProtocoleHTTP,LogFile,Data,i
Title = "Check URL © Hackoo 2014"
Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
sIniDir = WshShell.CurrentDirectory & "\url2check.txt"
sFilter = "Fichiers Textes (*.txt)|*.txt|Tous les fichiers (*.*)|*.*|"
sTitle = Title & "   Parcourir le fichier texte"
'sIniDir must be conformed to the javascript syntax
InFile = GetFileDlg(Replace(sIniDir,"\","\\"),sFilter,sTitle)
LogFile = Left(Wscript.ScriptFullName,InstrRev(Wscript.ScriptFullName, ".")) & "log"
If fso.FileExists(LogFile) Then fso.DeleteFile LogFile
Data = ReadFileText(InFile)
MsgBox Data
URL = Split(Data,VbCrLF)
ProtocoleHTTP = "http://"
For i = LBound(URL) To UBound(URL)
	If Left(URL(i),7) <> ProtocoleHTTP Then
		URL(i) = ProtocoleHTTP & URL(i)
		Call Check(URL(i))
	Else
		Call Check(URL(i))
	End if
Next
Call Explorer(LogFile)
'****************************************************************************************************************
Sub Check(URL)
    On Error Resume Next
    Dim xmlhttp
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")
    xmlhttp.open "HEAD",URL,False
    xmlhttp.send
    If Err = 0 Then
        Select Case Cint(xmlhttp.status)
        Case 200,202,301,302,401
            Set xmlhttp = Nothing
            Call WriteLog(DblQuote(URL) & VbTab &" ====> "& VbTab &" OK !")
        Case Else
            Set xmlhttp = Nothing
            Call WriteLog(DblQuote(URL) & VbTab &" ====> "& VbTab & "**/!\** KO **/!\**")
        End Select
    Else
        Call WriteLog(Err.Description)
    End If
End Sub
'******************************************************************************************************************
Function DblQuote(Str)
    DblQuote = Chr(34) & Str & Chr(34)
End Function
'******************************************************************************************************************
Sub WriteLog(strText)
Dim fs,ts 
Const ForAppending = 8
	Set fs = CreateObject("Scripting.FileSystemObject")
	Set ts = fs.OpenTextFile(Left(Wscript.ScriptFullName, InstrRev(Wscript.ScriptFullName, ".")) & "log", ForAppending, True)
	ts.WriteLine strText
	ts.Close
End Sub
'******************************************************************************************************************
Function GetFileDlg(sIniDir,sFilter,sTitle)
 GetFileDlg=CreateObject("WScript.Shell").Exec("mshta.exe ""about:<object id=d classid=clsid:3050f4e1-98b5-11cf-bb82-00aa00bdce0b></object><script>moveTo(0,-9999);function window.onload(){var p=/[^\0]*/;new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).Write(p.exec(d.object.openfiledlg('" & sIniDir & "',null,'" & sFilter & "','" & sTitle & "')));close();}</script><hta:application showintaskbar=no />""").StdOut.ReadAll
End Function
'******************************************************************************************************************
Function ReadFileText(sFile)
    Dim objFSO,oTS,sText
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set oTS = objFSO.OpenTextFile(sFile)
    sText = oTS.ReadAll
    oTS.close
    set oTS = nothing
    Set objFSO = nothing
    ReadFileText = sText
End Function 
'******************************************************************************************************************
Function Explorer(File)
	Dim ws
	Set ws=CreateObject("wscript.shell")
	ws.run "Explorer "& File & "\",1,True
end Function
'******************************************************************************************************************