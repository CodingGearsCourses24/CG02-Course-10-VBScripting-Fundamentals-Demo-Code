' Assignment Solution

Number = InputBox("Enter a number between 1 and 100: ") 

If Number > 100 OR Number < 1 Then
  MsgBox "You enter a number that is out of range!"
Else
	For x = 1 To Number Step 1
		result = x Mod 10
		If result = 0 Then
			MsgBox x, 0, "Result"
		End If
	Next
End If