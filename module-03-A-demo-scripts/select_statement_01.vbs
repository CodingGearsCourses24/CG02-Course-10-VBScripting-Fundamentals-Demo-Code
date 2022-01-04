'=========================================================================
' Select Case
'=========================================================================

option explicit

' Declare variables
Dim cardType

cardType = "Discover"

Select Case cardType
  Case "MasterCard"
    MsgBox "Payment processed by Master Card"
  Case "Visa"
    MsgBox "Payment processed by Visa"
  Case "American Express"
    MsgBox "Payment processed by American Express"
  Case "Discover"
    MsgBox "Payment processed by Discover"
  Case else
    MsgBox "Unknown Card"
End Select