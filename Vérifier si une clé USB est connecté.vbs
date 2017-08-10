'----------------------------NumSerie_Usb.vbs------------------------------
'Tester et vérifier si votre clé USB est connecté ou non, 
'si cette dernière est connectée alors le script nous donne son N° de Série.
'© Hackoo
'--------------------------------------------------------------------------
  Sub NumSerie_Usb()
  Dim NumSerie
  'Retrouver la clé Usb et son numéro de série
  Set fso = CreateObject("Scripting.FileSystemObject")
  For Each Drive In fso.Drives
  If Drive.IsReady Then
  If Drive.DriveType=1 Then
  NumSerie=fso.Drives(Drive + "\").SerialNumber
  MsgBox "La Clé Usb inséré a comme Num° de Série "&NumSerie,64,"Vérification Clé Usb © Hackoo"
  end if
  End If
  Next
  End Sub
 
'------------------------------checkUSB----------------------------
Sub checkUSB
strComputer = "."
On Error Resume Next
Set WshShell = CreateObject("Wscript.Shell")
beep = chr(007)
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_DiskDrive WHERE InterfaceType='USB'",,48)
intCount = 0
For Each drive In colItems
    If drive.mediaType <> "" Then
        intCount = intCount + 1
    End If
Next
If intCount > 0 Then
    MsgBox "Votre Clé USB Personnelle est bien Connectée !",64,"Flash Drive Check © Hackoo!"
	Call NumSerie_Usb() ' Appelle a la procédure NumSerie_Usb()
else
	WshShell.Run "cmd /c @echo " & beep, 0
	wscript.sleep 1000
	MsgBox "Votre Clé USB Personnelle n'est pas Connectée ",48,"Flash Drive Check © Hackoo !"
End If
End Sub
'---------------------------Fin du checkUSB----------------------------
Call checkUSB ' Appelle a la procédure checkUSB