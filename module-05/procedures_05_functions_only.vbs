'=========================================================================
' Sub Procedure --- does not return a value
' Function Procedure ------- can return a value. 
'=========================================================================

' *** Simple Calculator Functions ***

option explicit

' Variables
Const  SITE_TITLE = "www.CodingGears.com" 

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