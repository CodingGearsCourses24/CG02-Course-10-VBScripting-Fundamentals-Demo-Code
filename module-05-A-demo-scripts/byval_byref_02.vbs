' ===========================================================
' Using ByRef & ByVal with parameters/arguments
' 		ByVal ==> Passed by value
' 		ByRef ==> Passed by Reference
' ===========================================================

Dim MyNumber
MyNumber = 100

MsgBox "M0: " & MyNumber

'PlayWithNumbers MyNumber
'PlayWithNumbers(MyNumber)
Call PlayWithNumbers(MyNumber)
'Call PlayWithNumbers((MyNumber))

MsgBox "M1: " & MyNumber

Function PlayWithNumbers(ByRef MyParam)
	MyParam = 25
End Function