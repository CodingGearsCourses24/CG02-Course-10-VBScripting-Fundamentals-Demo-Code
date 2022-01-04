'=========================================================================
' Built-in Functions:
' IsNumeric, IsArray, IsDate
'=========================================================================

num1 = 24
num2 = "24"
num3 = "two"
array1 = Array(1,2,3,4,5)
date1 = Now
obj1 = 123
Set obj2 = CreateObject("Scripting.FileSystemObject")

DisplayMessage IsNumeric(num1), "IsNumeric num1 >> "
DisplayMessage IsNumeric(num2), "IsNumeric num2 >> "
DisplayMessage IsNumeric(num3), "IsNumeric num3 >> "

DisplayMessage IsArray(array1), "IsArray array1 >> "
DisplayMessage IsDate(date1), "IsDate date1 >> "
DisplayMessage IsObject(obj1), "IsObject obj1 >> "
DisplayMessage IsObject(obj2), "IsObject obj2 >> "

Function DisplayMessage(message, id)
	MsgBox id & " : " & message,0,"Welcome"
End Function

