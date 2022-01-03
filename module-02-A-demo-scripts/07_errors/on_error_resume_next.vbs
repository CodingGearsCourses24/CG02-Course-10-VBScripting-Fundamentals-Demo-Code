'================================================================
' On Error Resume Next
' On Error GoTo 0
'       Enables or disables error-handling
'================================================================

option explicit


' Numeric Variables
Dim num1, num2, div

num1 = 10
num2 = 0

' Processing 1
On Error Resume Next
div = num1 / num2


MsgBox div

MsgBox "We are done!"

' Processing 2
On Error GoTo 0
div = num1 / num2