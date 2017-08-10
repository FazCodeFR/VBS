strComputer = "."  
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")  
  
Set colItems = objWMIService.ExecQuery ("Select * from Win32_NetworkAdapter")  
   
   For Each objItem In colItems  
        
      strConnectionStatus = objItem.NetConnectionStatus  
    
      If strConnectionStatus = 2 Then  
        strMacAddress = objItem.MACAddress  
      End If  
   Next   
       
 Set objAdapters = objWMIService.ExecQuery ("Select * from Win32_NetworkAdapterConfiguration WHERE IPEnabled = 'True' AND MACAddress='" & strMacAddress & "'")  
   
For Each objAdapter in objAdapters  
   
   If Not IsNull(objAdapter.DefaultIPGateway) Then  
      For i = 0 To UBound(objAdapter.DefaultIPGateway)  
             FAIIP = objAdapter.DefaultIPGateway(i)
      Next  
   End If  
Next  
MsgBox FAIIP