'=================================================================================
' Wscript.Network - Map n/w drives
'=================================================================================
Dim oNetwork
Dim oDrives 

' Objects
Set oNetwork = Wscript.CreateObject("Wscript.Network")

' Mapping N/W drive
If oNetwork.UserName = "codin1" then
   oNetwork.MapNetworkDrive "K:", "\\127.0.0.1\zTemp1"
Else
   WScript.Echo "Error!" & " You are logged in as " & oNetwork.UserName
end If

' List
Set oDrives  = oNetwork.EnumNetworkDrives
If oDrives.Count = 0 Then
   Msgbox "Mapped network drives not found!"
Else
   drives_count = oDrives.Count
   For x = 0 to (drives_count -1) Step 2
      WScript.Echo "Drive " & oDrives.Item(x) & " = " & oDrives.Item(x+1)
   Next
End If