'================================================================
' Operators:
'       Arthritic Operators 	: ^, -, *, /, \, Mod, +, -, &
'       Comparison  			: =, <>, <, >, <=, >=, Is
'       Logical				    : Not, And, Or, Xor, Eqv, Imp 
'================================================================

option explicit
'On Error Resume Next

' Declare variables
Const  SITE_TITLE = "www.CodingGears.com" 

Dim num1, num2, sum, product, power

num1 = 3
num2 = 4

sum = num1 + num2
product = num1 * num2
power = num1 ^ num2


MsgBox "The sum of " & num1 & " and " & num2 & " is " & sum, 0, SITE_TITLE

MsgBox "The product of " & num1 & " and " & num2 & " is " & product, 0, SITE_TITLE

MsgBox num1 & " raised to the power of " & num2 & " is " & power, 0, SITE_TITLE