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