 strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" _
    & strComputer & "\root\cimv2")
Set colSMBIOS = objWMIService.ExecQuery _
    ("Select * from Win32_SystemEnclosure")
For Each objSMBIOS in colSMBIOS
    Wscript.Echo "Part Number: " & objSMBIOS.PartNumber
    Wscript.Echo "Serial Number: " _
        & objSMBIOS.SerialNumber
    Wscript.Echo "Asset Tag: " _
        & objSMBIOS.SMBIOSAssetTag
Next
