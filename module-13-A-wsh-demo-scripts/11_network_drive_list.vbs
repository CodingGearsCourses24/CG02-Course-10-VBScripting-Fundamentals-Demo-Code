'==========================================================
' Wscript.Network (This is the WSH Network object)
'        Connect to and disconnect from network shares and network printers
'        Map and unmap network shares
'        Access information about the currently logged-on user
'==========================================================

Dim oNetwork
Dim oDrives 
Dim x
Dim drives_count
Dim strMessage

CreateObjects
' CheckForNetworkDrives1
CheckForNetworkDrives2
' DisplayOtherInfo
CloseObjects

'==========================================================
' Create Objects
'==========================================================
Sub CreateObjects
   set oNetwork = Wscript.CreateObject("Wscript.Network")
   set oDrives  = oNetwork.EnumNetworkDrives
End Sub

'==========================================================
' Network Drives 1
'==========================================================
Sub CheckForNetworkDrives1
   If oDrives.Count = 0 Then
      Msgbox "Mapped network drives not found!"
   Else
      drives_count = oDrives.Count
      For x = 0 to (drives_count - 1) Step 2
         WScript.Echo "Drive " & oDrives.Item(x) & " = " & oDrives.Item(x+1)
      Next
   End If
End Sub

'==========================================================
' Network Drives 2 (Uses Msgbox)
'==========================================================
Sub CheckForNetworkDrives2
   If oDrives.Count = 0 Then
      Msgbox "Mapped network drives not found!"
   Else
      drives_count = oDrives.Count
      strMessage = (drives_count / 2) & " mapped network drive(s)" & vbCRLF
      For x = 0 to (drives_count -1) Step 2
         strMessage = strMessage & oDrives (x) & " = " & oDrives.Item(x+1) & vbCRLF
      Next
      Msgbox strMessage
   End If
End Sub

'==========================================================
' Network Misc Properties
'==========================================================
Sub DisplayOtherInfo
   WScript.Echo oNetwork.ComputerName
   WScript.Echo oNetwork.UserName
   WScript.Echo oNetwork.UserDomain
End Sub
'==========================================================
' Close Objects
'==========================================================
Sub CloseObjects
    Set oNetwork = Nothing
    Set oDrives = Nothing
End Sub