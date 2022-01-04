'=========================================================================
' Arrays
' ReDim - change the size of the last dimension.
'=========================================================================

' Error Handling
option explicit

' Variables
Const  SITE_TITLE = "www.CodingGears.com" 

Dim arrFriends()
ReDim arrFriends(2)
arrFriends(0) = "John"
arrFriends(1) = "Peter"
arrFriends(2) = "Jimmy"

MsgBox arrFriends(1), 0, SITE_TITLE
'MsgBox arrFriends(4), 0, SITE_TITLE

ReDim arrFriends(4)
arrFriends(3) = "Mike"
arrFriends(4) = "Kenny"

MsgBox arrFriends(4), 0, SITE_TITLE

ReDim arrFriends(9)

arrFriends(8) = "Penny"
MsgBox arrFriends(8), 0, SITE_TITLE
MsgBox arrFriends(4), 0, SITE_TITLE
