' Assignment Solution

option explicit

Const  SITE_TITLE = "www.GlobalETraining.com" 

'**********************************************************************
'Display Message
Sub DisplayMsg(strMessage, intResult)
   MsgBox strMessage & " : " & intResult,0,SITE_TITLE
End Sub

'Add Function
Function Add(inta, intb)
	Add = inta + intb
End Function

'Subtract Function
Function Subtract(inta, intb)
	Subtract = inta - intb
End Function

'Multiply Function
Function Multiply(inta, intb)
	Multiply = inta * intb
End Function

'Divide Function
Function Divide(inta, intb)
	Divide = inta / intb
End Function

'Power Function
Function PowerOf(inta, intb)
	PowerOf = inta ^ intb
End Function
'**********************************************************************


Dim a, b, result, operation

a = CInt(InputBox("Enter the first number : ")) 
b = CInt(InputBox("Enter the second number : "))

operation = InputBox("Enter: " & vbNewLine  & "     1 for Add " & vbNewLine  & "     2 for Subtract " & vbNewLine  & "     3 for Multiply " & vbNewLine  & "     4 for Divide " & vbNewLine  & "     5 for PowerOf ") 

Select Case operation
  Case 1
    result = Add(a, b)
  Case 2
    result = Subtract(a, b)
  Case 3
    result = Multiply(a, b)
  Case 4
    result = Divide(a, b)
  Case 5
    result = PowerOf(a, b)
  Case else
    MsgBox "Your selection is not VALID!", 48, "ERROR"
End Select

If operation > 0 AND operation < 6 Then
	DisplayMsg "The result is ", result
End If