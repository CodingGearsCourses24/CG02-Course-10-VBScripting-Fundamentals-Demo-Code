'================================================================
' Err Object
'       Contains information about run-time errors. 
'       Accepts the Raise and Clear methods for generating 
'       and clearing run-time errors.
'       >> Raise & Clear
'================================================================

option explicit

' Numeric Variables
Dim num1, num2, div

num1 = 10
num2 = 0

' Processing
On Error Resume Next
div = num1 / num2

Err.Raise 101
MsgBox "Error: " & CStr(Err.Number) & " " & Err.Description

Err.Clear 'Reset
MsgBox "Error: " & CStr(Err.Number) & " " & Err.Description & "EOM"
