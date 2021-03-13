'=========================================================================
' Arrays
'=========================================================================

' Error Handling
option explicit

' Variables
Dim PhoneBook(2, 4) ' 3 Columns, 5 Rows
Dim colLowerIndex, colHigherIndex, rowLowerIndex, rowHigherIndex
Dim colIndex, rowIndex, searchName, matched_row, ci, found, ri
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

PhoneBook(0,3) = "Lilly"
PhoneBook(1,3) = "Houston"
PhoneBook(2,3) = "111-111-0003"

PhoneBook(0,4) = "Spark"
PhoneBook(1,4) = "Denton"
PhoneBook(2,4) = "111-111-0004"

' Getting the counts for processing
colLowerIndex = 0
colHigherIndex = UBound(PhoneBook, 1)

rowLowerIndex = 0
rowHigherIndex = UBound(PhoneBook, 2)

' Using nested for loop
for ri = rowLowerIndex to rowHigherIndex '0, 1, 2, 3, 4
    for ci = colLowerIndex to colHigherIndex '0, 1, 2
        WScript.Echo PhoneBook(ci, ri) '0,0   1,0   2,0
    Next
    WScript.Echo "-------"
Next

