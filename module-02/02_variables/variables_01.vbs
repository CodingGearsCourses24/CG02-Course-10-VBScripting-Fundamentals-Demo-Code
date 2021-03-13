' ======================================================
' Variables, Constants, Dim, &, &_
' The Dim keyword is used to declare a variable.
' The Dim keyword is short for Dimension. 
' &  --> String concatenation
' &_ --> Line Continuation 
' ======================================================

' Variables
Dim num1, num2
Dim sum

num1 = 10
num2 = 20

' Processing
sum = num1 + num2

MsgBox sum

MsgBox "The sum of " & num1 & " and " & num2 & " is " & sum

MsgBox "====> The sum of " & num1 & " and " & num2 &_
				" is " & sum