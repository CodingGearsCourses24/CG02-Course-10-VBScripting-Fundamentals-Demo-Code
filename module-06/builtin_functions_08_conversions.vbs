'=========================================================================
' Built-in Functions:
' CBool, CCur, CDate, Cint, Hex, Oct
'=========================================================================

var_num1 = 24
var_num2 = -52.25
var_num3 = 0
var_num4 = 4605

num_a = 123.5691716

MyDate = "April 15, 1998"   ' Define date.
MyTime = "5:15:32 PM"         ' Define time.

DisplayMessage CBool(var_num1), "CBool var_num1 >> "
DisplayMessage CBool(var_num2), "CBool var_num2 >> "
DisplayMessage CBool(var_num3), "CBool var_num3 >> "

DisplayMessage Hex(var_num4), "Hex >> " 
DisplayMessage Oct(var_num4), "Oct >> "

DisplayMessage CCur(num_a), "CCur >> " 
DisplayMessage CInt(num_a), "CInt >> " 

DisplayMessage CDate(MyDate), "CDate MyDate >> " 
DisplayMessage CDate(MyTime) , "CDate MyTime >> " 

 

Function DisplayMessage(message, id)
	MsgBox id & " : " & message,0,"Welcome"
End Function
