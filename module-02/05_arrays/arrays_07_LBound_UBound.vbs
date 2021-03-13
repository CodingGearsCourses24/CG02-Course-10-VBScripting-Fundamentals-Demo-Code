'=========================================================================
' Arrays - LBound, UBound
'=========================================================================

' Error Handling
option explicit


' Declare variables
Const  SITE_TITLE = "www.CodingGears.com" 
Dim lowerIndex, upperIndex
Dim arrNums(4)

arrNums(0) = 100
arrNums(1) = 200
arrNums(2) = 300
arrNums(3) = 400
arrNums(4) = 500

lowerIndex = LBound(arrNums)
upperIndex = UBound(arrNums)

MsgBox "Lower Index is " & lowerIndex, 0, SITE_TITLE
MsgBox "Upper Index is " & upperIndex, 0, SITE_TITLE