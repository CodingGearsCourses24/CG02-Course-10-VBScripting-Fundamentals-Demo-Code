'=========================================================================
' Arrays
' ReDim Preserve  - Save the data in an existing array 
'                   when you change the size of the last dimension.
'=========================================================================

' Error Handling
option explicit

' Declare variables
Const  SITE_TITLE = "www.CodingGears.com" 

' Declare an array named A with 10 elements
Dim arrFriends
arrFriends = Array("John", "Peter", "Jimmy", "Mike", "Kevin", "Tom")

ReDim arrFriends(9)
arrFriends(8) = "Gary"
arrFriends(9) = "Kelly"


MsgBox arrFriends(4), 0, SITE_TITLE
MsgBox arrFriends(5), 0, SITE_TITLE
MsgBox arrFriends(8), 0, SITE_TITLE
MsgBox arrFriends(9), 0, SITE_TITLE
