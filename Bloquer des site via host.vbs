dim Fso,f
Dim rep,label,titre,defaut,data
label="Entrez dans le champ ci-dessous Le site que vous-voulez Bloquer Exemple www.pagedepubs.com"
defaut=""
titre="Bloquer les Sites Interdits"
rep=InputBox(label,titre,defaut)
Set Fso = CreateObject("Scripting.FileSystemObject")
sys32=Fso.GetSpecialFolder(1)
Set f = fso.OpenTextFile(sys32+"\DRIVERS\ETC\hosts", 8)
if rep="" then Cleanup
f.Write vbnewline
f.Write "127.0.0.1   "  &rep
 
Sub Cleanup()
  Set FSO = Nothing
  WScript.Quit
End Sub