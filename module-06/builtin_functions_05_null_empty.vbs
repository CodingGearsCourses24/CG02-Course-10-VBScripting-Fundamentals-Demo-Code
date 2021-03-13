' Using built-in functions - IsNUll, IsEmpty, Null, Empty
' Empty : The Empty keyword is used to indicate an uninitialized variable value.

Dim var, var_msg, var_null, var_empty

var_data = 100
var_null = Null
var_empty = Empty

result_isempty = IsEmpty(var_empty)
result_isnull = IsNull(var_empty)

DisplayMessage result_isempty, "IsEmpty"
DisplayMessage result_isnull, "IsNull'"

Function DisplayMessage(message, id)
	MsgBox id & " : " & message,0,"Welcome"
End Function
