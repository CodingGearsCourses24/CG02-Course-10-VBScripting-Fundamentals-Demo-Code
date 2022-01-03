'=========================================================================
' Using Operators with IF Statements
'=========================================================================

option explicit

' Declare variables
Dim name1, name2
Dim x1, x2, x3, x4, x5

name1 = "John"
name2 = "Peter"

' Using <>
If name1 <> name2 Then
	MsgBox "Names did not match"
End If

x1 = 1
x2 = 1
x3 = 2
x4 = 2
x5 = 20

' Using And
If x1 = x2 And x3 = x4 Then
	MsgBox "And condition met!"
End If

' Using OR
If x1 = x2 Or x4 = x5 Then
	MsgBox "OR codition met!"
End If

