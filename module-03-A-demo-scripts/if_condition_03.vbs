'=========================================================================
' If..Then...ElseIf
'=========================================================================

option explicit

Dim x
x = 500
	
If x = 100 Then
    ' Processing - True
    MsgBox "Message 1: The number is equal to 100!"
ElseIf x > 100 Then
    ' Processing - False
    MsgBox "Message 2: The number is greater than 100!"
Else
    MsgBox "Message 3: The number is less than 100!"
End If
