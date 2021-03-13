'=========================================================================
' Select Case
'=========================================================================

Dim number

number = 500

Select Case number
  Case 100
    MsgBox "We got 100"
  Case 200
    MsgBox "We got 200"
  Case 300, 400, 500, 600
    MsgBox "We got 300 or 400 or 500 or 600"
  Case else
    MsgBox "We got something else!"
End Select