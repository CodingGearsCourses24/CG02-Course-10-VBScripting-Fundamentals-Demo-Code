'=========================================================================
' Sub Procedure --- does not return a value
' Function Procedure ------- can return a value. 
'=========================================================================

option explicit

' Variables
Dim num1, num2, total, sum

num1 = 10
num2 = 50

total  = Add(num1, num2)

WScript.Echo "The sum of the two numbers : " & total

'A Function procedure -- can return a value. 
Function Add(num1, num2)
    sum = num1 + num2
End Function