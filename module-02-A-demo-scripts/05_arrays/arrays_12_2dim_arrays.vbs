'=========================================================================
' Arrays
'=========================================================================

' Error Handling
option explicit

' Variables
Dim PhoneBook()

ReDim PhoneBook(2, 0) '3 Cols, 1 Row
PhoneBook(0,0) = "Peter"
PhoneBook(1,0) = "Boston"
PhoneBook(2,0) = "111-111-0000"

MsgBox "1: " & PhoneBook(0,0)

ReDim preserve PhoneBook(2, 2) '3 Cols, 3 Rows 

MsgBox "2: " & PhoneBook(0,0)

ReDim preserve PhoneBook(2, 5) '3 Cols, 6 Rows 

MsgBox "3: " & PhoneBook(0,0)

ReDim preserve PhoneBook(2, 20) '3 Cols, 21 Rows 
MsgBox "3: " & PhoneBook(0,0)