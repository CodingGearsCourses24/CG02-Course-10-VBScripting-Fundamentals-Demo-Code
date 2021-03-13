'=========================================================================
' Arrays - Erase
'=========================================================================

' Error Handling
option explicit

' Variables
Dim arrAnimals
arrAnimals = Array("Lion", "Tiger", "Chicken", "Snake", "Elephant")

' Erase ArrayName
MsgBox "===> Checkpoint A "
MsgBox arrAnimals(1)
Erase arrAnimals
MsgBox "===> Checkpoint B : Erased "
MsgBox arrAnimals(1)