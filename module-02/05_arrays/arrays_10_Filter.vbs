'=========================================================================
' Arrays - Filter
'=========================================================================

' Error Handling
option explicit

' Variables
Const  SITE_TITLE = "www.CodingGears.com" 
Dim arrAnimals
Dim matched, not_matched, item

' Filter (Filter(inputstrings,value[,include[,compare]]))
' Compare:
'   0 = vbBinaryCompare (Binary comparison)
'   1 = vbTextCompare (Textual comparison)
arrAnimals = Array("Lion", "Tiger", "Chicken", "Snake", "Elephant")

matched = Filter(arrAnimals, "a", True, 1)

MsgBox matched(0)
MsgBox matched(1)

not_matched = Filter(arrAnimals, "a", False, 1)
MsgBox not_matched(0)
MsgBox not_matched(1)
MsgBox not_matched(2)
