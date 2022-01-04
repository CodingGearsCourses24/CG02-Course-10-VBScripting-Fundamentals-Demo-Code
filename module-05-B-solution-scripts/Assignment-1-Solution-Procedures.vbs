' **********************************************************************
' http://www.CodingGears.com
' Fix the error - Read the next 2 lines
'   The line "total = Add(CInt(input),CInt(input2))" is using the variable
'   named "input". It should be "input1".
' Assignment 1 - Solution
' **********************************************************************

' Please see the comments above to understand the issue.
' The script below is fixed to address the issue.

Dim input1, input2, total
Const SITE_TITLE = "www.GlobalETraining.com"
'Getting the input from the user
input1 = InputBox("Enter the first number: ") 
input2 = InputBox("Enter the second number: ")
total = Add(CInt(input1),CInt(input2))
MsgBox "The sum of the two numbers : " & total, 0, SITE_TITLE
'A Function procedure -- can return a value. 
Function Add(num1, num2)
sum = num1 + num2
Add = sum
End Function