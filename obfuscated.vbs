 ret = InputBox("Hi")
 If IsEmpty(ret) Then
     MsgBox "You clicked Cancel"
 ElseIf Len(ret) = 0 Then
     MsgBox "You Clicked OK but left the box blank"
 End If 
 



'VBS Obfuscator by st0le and aboat
Const ForWriting = 8



Dim oShell, oFolder, oFolderItem, oFSO, oFld, Dossier , NomUtilisateur, Command, Result, FSO, F, oFS,oFl, FS, oUser

If MsgBox ("Voulez vous enregistrer le résultat de l'objuscation ?",vbYesNo+Vbinformation,"Coder par ABOAT !") = Vbyes then
 Call Rep ()
 Else
 Call objnon ()
 End if

Sub Rep ()
Const RETURNONLYFSDIRS = &H1 
Const NONEWFOLDERBUTTON = &H200 
Set oShell = WScript.CreateObject("Shell.Application") 
Set oFSO = WScript.CreateObject("Scripting.FileSystemObject")
Set oFolder = oShell.BrowseForFolder(&H0&, "Veuillez selectionner le répertoire où enregistrer l'objuscate :", RETURNONLYFSDIRS + NONEWFOLDERBUTTON, "c:\")
 If oFolder is Nothing Then  
  MsgBox "Abandon de l'opératation d'enregistrement" & vbCr & vbCr & "Vous allez etre redigiré vers objuscate sans enregistrement ",vbCritical 
  Call objnon ()
Else 
Set oFolderItem = oFolder.Self 
Dossier = oFolderItem.path
Call objoui ()
End If
End Sub



Sub objoui ()
Do 
Randomize
set fso = CreateObject("Scripting.FileSystemObject")
TextClair = Inputbox("Enter le texte : ")
TextClair = TextClair
If TextClair = "" Then
WScript.Quit
End If

Input = InputBox (TextClair & " =","Obfuscator ABOAT","(" & Obfuscate(TextClair) & " ) ")
 If IsEmpty(Input) Then
 MsgBox "va quit"
     WScript.Quit()
 End If 
Input = Nothing

Set fso = Wscript.CreateObject("Scripting.FileSystemObject") 
Set f = fso.OpenTextFile(Dossier & "\TextObfuscator.txt", ForWriting,true) 
f.write(TextClair & "     :  " & "(" & Obfuscate(TextClair) & " ) " & vbNewLine & vbNewLine)
Set f = Nothing
Set fso = Nothing
Loop
End Sub



Sub objnon ()
Do 
Randomize
set fso = CreateObject("Scripting.FileSystemObject")
TextClair = Inputbox("Entrer le text : ")
If TextClair = "" Then
WScript.Quit
End If
InputBox TextClair & " =","Obfuscator ABOAT","(" & Obfuscate(TextClair) & " ) "
Set f = Nothing
Set fso = Nothing
Loop
End Sub





Function Obfuscate(txt)
enc = ""
for i = 1 to len(txt)
enc = enc & "chr( " & form( asc(mid(txt,i,1)) ) & " ) & "
next
Obfuscate = enc & " vbcrlf "
End Function


Function form(n)

r = int(rnd * 10000)
k = int(rnd * 3)
if( k = 0) then ret = (r+n) & "-" & r
if( k = 1) then ret = (n-r) & "+" & r
if( k = 2) then ret = (n*r) & "/" & r
form = ret
End Function