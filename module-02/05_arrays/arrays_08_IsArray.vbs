'=========================================================================
' Arrays - isArray
'=========================================================================

' Error Handling
option explicit

' Variables
Const  SITE_TITLE = "www.CodingGears.com" 
Dim num1, b, x, bnot
Dim arrNums(2)
Dim arrAnimals

arrNums(0) = 100
arrNums(1) = 200

MsgBox "Checking if arrNums is an array  : " & IsArray(arrNums) , 0, SITE_TITLE
MsgBox "Checking if num1 is an array  : " & IsArray(num1) , 0, SITE_TITLE
