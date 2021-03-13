'=========================================================================
' Loops - FOR
' Processing 1-Dimensional array
'=========================================================================

option explicit

' Variables
Dim total, arrNums, i

' Array
arrNums=Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11)

' Using array concepts in a FOR loop
For  i = LBound(arrNums) To UBound(arrNums) 
  total = total + arrNums(i)
  If total > 10 Then
	  WScript.Echo "Hurray! The total is greater than 10!" & VBCrLf & "Current total is " & total & "! !"
  End If
  WScript.Echo "- - -"
Next

WScript.Echo  "The sum of all the elements in the array is " & total