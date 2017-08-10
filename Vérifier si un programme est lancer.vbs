Dim a

a = InputBox ("Quelle processus veut tu verif ?")
b = "'" & a & "'"
Set objWMI = GetObject("winmgmts:root\cimv2")
    sQuery = "Select * from Win32_process Where Name = " & b & " "
If objWMI.execquery(sQuery).Count = 1 Then
 MsgBox "Le processus " & a & " est lancer !"
 WScript.Quit
End If

MsgBox "Le processus" & a & " n'est pas lancer "
MsgBox "Suite"
WScript.Quit