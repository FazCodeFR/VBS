Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\.rootcimv2")
Set colComputers = objWMIService.ExecQuery _
("Select * from Win32_ComputerSystem")
For Each objItem in colComputers
nom = objItem.Name
Next

Nom = InputBox("Saisir le nouveau nom du PC " & nom ,"Nom","Mon_PC")

For Each objComputer in colComputers
errReturn = ObjComputer.Rename(nom)
If errReturn Then
WScript.Echo "Error N° " & errReturn & _
vbNewLine & _
"Description : " & Err.Description
End If
Next