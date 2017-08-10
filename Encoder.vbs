Const SW_SHOWNORMAL=1 
Const ForReading=1
Const ForWriting=2
Set fso     = WScript.CreateObject("Scripting.FileSystemObject")
Set shell   = WScript.CreateObject("WScript.Shell")
Set UnNamed = WScript.Arguments.UnNamed

nu=UnNamed.count
If nu<2 Then
	WScript.Arguments.ShowUsage
	WScript.Quit
	End If
Srce=UnNamed(0)
Dest=UnNamed(1)
Display=false
If nu>=3 then if lcase(UnNamed(2))="a" Then Display=true

If not fso.FileExists(Srce)  Then
	MsgBox "Le fichier " & Srce & " n'a pas été trouvé", vbCritical + vbOKOnly, "Encodage de fichier"
	WScript.Quit
	End If
If fso.FileExists(Dest)  Then
	rep=MsgBox("Le fichier destination " & Dest & " existe déjà." & VBCRLF & _
	"Faut-il l'écraser ?", vbQuestion+vbYesNo, "Encodage de fichier")
	If rep<>vbYes Then WScript.Quit
	End If
Set fs=fso.GetFile(Srce)  
pe=InstrRev(fs.Name,".")
If pe>0 then ext=lcase(mid(fs.Name,pe)) else ext=".vbs"
If ext=".wsf" Then ext=".html"

Set fs = fso.OpenTextFile(Srce,ForReading,false)
data=fs.readAll
fs.close
Set se=WScript.CreateObject("Scripting.Encoder")
dataenc=se.EncodeScriptFile(ext,data, 0, "")

Set fd = fso.CreateTextFile(Dest, true,false)
fd.Write dataenc
fd.close
If Display Then
	cmd="notepad.exe """ & Dest & """"
	shell.run cmd,SW_SHOWNORMAL,false
	End If
WScript.Quit