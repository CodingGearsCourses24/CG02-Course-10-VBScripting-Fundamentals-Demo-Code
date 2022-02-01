'=================================================================================
' Wscript.Network - Disconnect
'=================================================================================

Dim oNetwork

set oNetwork = Wscript.CreateObject("Wscript.Network")

if oNetwork.UserName = "codin" then

   oNetwork.RemoveNetworkDrive("Y:")

else
   WScript.Echo "Error!!!"
end if

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