
MsgBox stVersion

'
' Lecture version windows
'
Function stVersion() 
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colOperatingSystems = objWMIService.ExecQuery _
    ("Select * from Win32_OperatingSystem")
 
For Each objOperatingSystem in colOperatingSystems
 
	stVersion = objOperatingSystem.Caption 
Next
end function