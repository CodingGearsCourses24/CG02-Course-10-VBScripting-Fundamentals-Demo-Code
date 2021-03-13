'=========================================================================
' Arrays
'=========================================================================

' Error Handling
option explicit

' Variables
Const  SITE_TITLE = "www.CodingGears.com" 

' Arrays
Dim arrAnimals, animal1

' Assigning values to the elements in A
arrAnimals = Array("Lion", "Tiger", "Chicken", "Snake", "Elephant")

' The value of an element in the array is assigned to a variable
' arrAnimals(5) = "Dog"
animal1 = arrAnimals(3) 

MsgBox "The animal is " & animal1, 0, SITE_TITLE