'=========================================================================
' Select Case
'=========================================================================
' weekday takes date as the input, and return a number.
' ----------------------------------
' Constant	 Value	Description
' ----------------------------------
' vbSunday	  1	    Sunday
' vbMonday	  2	    Monday
' vbTuesday	  3	    Tuesday
' vbWednesday	4	    Wednesday
' vbThursday	5	    Thursday
' vbFriday	  6	    Friday
' vbSaturday	7	    Saturday
'=========================================================================

' Declare variables
option explicit

Dim x
x = weekday(date)

Select Case x
  Case 1
    MsgBox "Sunday"
  Case 2
    MsgBox "Monday"
  Case 3
    MsgBox "Tuesday"
  Case 4
    MsgBox "Wednesday"
  Case 5
    MsgBox "Thursday"
  Case vbFriday
    MsgBox "Friday"
  Case else
    MsgBox "Saturday" 
End Select
