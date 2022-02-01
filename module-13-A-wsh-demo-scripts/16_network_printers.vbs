'=================================================================================
' Wscript.Network
'=================================================================================
'
Dim oNetwork
Dim oDrives 
Dim oPrinters
Dim x
Dim printers_count
Dim strMessage


CreateObjects
CheckForPrinters
GetPrinterCount
CloseObjects

'==========================================================
' Create Objects
'==========================================================
Sub CreateObjects
   Set oNetwork = Wscript.CreateObject("Wscript.Network")
   Set oPrinters = oNetwork.EnumPrinterConnections
End Sub

'==========================================================
' Check For Printers
'==========================================================
Sub CheckForPrinters
   If oPrinters.Count = 0 Then
      Msgbox "Printers not found!"
   Else
      printers_count = oPrinters.Count
      For x = 0 to (printers_count -1) Step 2
         WScript.Echo "Printer " & oPrinters.Item(x) & " = " & oPrinters.Item(x+1)
      Next
   End If
End Sub

'==========================================================
' Printer - Count
'==========================================================
Sub GetPrinterCount
   If oPrinters.Count = 0 Then
      Msgbox "Printers not found!"
   Else
      printers_count = oPrinters.Count
      WScript.Echo "Printer Count " & (printers_count / 2)
   End If
End Sub

'=================================================================================
' Close Objects
'=================================================================================
Sub CloseObjects
    Set oNetwork = Nothing
    Set oPrinters = Nothing
End Sub