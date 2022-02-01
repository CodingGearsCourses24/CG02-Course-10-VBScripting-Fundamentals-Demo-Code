'=================================================================================
' Wscript.Network - Check the logged in user
'=================================================================================

Dim oNetwork

Set oNetwork = Wscript.CreateObject("Wscript.Network")

If oNetwork.UserName = "Codin" then
   Wscript.echo "Logged in as codin. Script will abort."
Else
   Wscript.echo "You are not logged in as codin. Proceeding..."
End If
