'=========================================================================
' Arrays
'=========================================================================

' Error Handling
option explicit

' Variables
Dim PhoneBook()

ReDim PhoneBook(4, 25)

MsgBox LBound(PhoneBook, 1)
MsgBox UBound(PhoneBook, 1)

MsgBox LBound(PhoneBook, 2)
MsgBox UBound(PhoneBook, 2)