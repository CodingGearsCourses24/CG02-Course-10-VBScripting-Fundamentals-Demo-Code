'=========================================================================
' Arrays
'=========================================================================

' Error Handling
option explicit

' Variables
Dim PhoneBook(2, 4) ' 3 Columns, 5 Rows

Const  SITE_TITLE = "www.CodingGears.com"   

' Assigning values to the elements in A
' PhoneBook(columns, rows)
PhoneBook(0,0) = "Peter"
PhoneBook(1,0) = "Boston"
PhoneBook(2,0) = "111-111-0000"

PhoneBook(0,1) = "Mike"
PhoneBook(1,1) = "San Jose"
PhoneBook(2,1) = "111-111-0001"

PhoneBook(0,2) = "Sara"
PhoneBook(1,2) = "Denver"
PhoneBook(2,2) = "111-111-0002"

PhoneBook(0,2) = "Lilly"
PhoneBook(1,2) = "Houston"
PhoneBook(2,2) = "111-111-0003"

PhoneBook(0,2) = "Spark"
PhoneBook(1,2) = "Denton"
PhoneBook(2,2) = "111-111-0004"

MsgBox PhoneBook(0,0) & " lives in " & PhoneBook(1,0) & ". " & vbCrLf & "Contact No: " & PhoneBook(2,0)
MsgBox PhoneBook(0,1) & " lives in " & PhoneBook(1,1) & ". " & vbCrLf & "Contact No: " & PhoneBook(2,1)
MsgBox PhoneBook(0,2) & " lives in " & PhoneBook(1,2) & ". " & vbCrLf & "Contact No: " & PhoneBook(2,2)
