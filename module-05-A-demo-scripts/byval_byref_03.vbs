' ===========================================================
' Using ByRef & ByVal with parameters/arguments
' 		ByVal ==> Passed by value
' 		ByRef ==> Passed by Reference
' ===========================================================

Dim mynum1, mynum2, result
mynum1 = 5
mynum2 = 7

result = AddNumbers(mynum1, mynum2)

MsgBox "M1: Result is " & result
MsgBox "M2: mynum1 = " & mynum1 & " and mynum2 = " & mynum2

Function AddNumbers(num1, num2)  '<======== ByRef, default
	AddNumbers = num1 + num2
	num1 = 11111
	num2 = 22222
End Function