' Assignment Solution

Number = InputBox("Enter a number between 1 and 25: ") 

If Number > 25 OR Number < 1 Then
  MsgBox "You enter a number that is out of range!"
Else
	For x = 1 To Number Step 1
		MsgBox x, 0, "Number :"
	Next
End If