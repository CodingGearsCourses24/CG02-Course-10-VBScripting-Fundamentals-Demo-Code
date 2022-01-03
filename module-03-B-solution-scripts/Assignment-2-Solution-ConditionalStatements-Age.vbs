' Assignment Solution


JohnAge = InputBox("Enter John's age: ") 
KerryAge = InputBox("Enter Kerry's age: ") 


If JohnAge > KerryAge Then
	MsgBox "John is older than Kerry." 
ElseIf JohnAge < KerryAge Then
	MsgBox "John is younger than Kerry." 
ElseIf JohnAge = KerryAge Then
	MsgBox "John and Kerry are both " & JohnAge & " years old."
End If