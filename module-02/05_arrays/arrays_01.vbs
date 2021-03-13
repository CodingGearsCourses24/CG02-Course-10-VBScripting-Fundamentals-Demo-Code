'=========================================================================
' Arrays
'=========================================================================

' Error Handling
option explicit

' Variables
Const  SITE_TITLE = "www.CodingGears.com" 

' Arrays
Dim arrNums(4) 

' Assigning values to the elements in A
arrNums(0) = 100
arrNums(1) = 200
arrNums(2) = 300
arrNums(3) = 400
arrNums(4) = 500

Dim num1

' The value of an element in the array is assigned to a variable
num1 = arrNums(1) 

MsgBox "Value of num1 is " & num1, 0, SITE_TITLE

MsgBox "Value of arrNums(2) is " & arrNums(2), 0, SITE_TITLE
MsgBox "Value of arrNums(3) is " & arrNums(3), 0, SITE_TITLE

