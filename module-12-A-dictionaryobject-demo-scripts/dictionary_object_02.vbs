'=========================================================================
' Dictionary Object
'=========================================================================

Dim salaries   ' Create a variable.
Set salaries = CreateObject("Scripting.Dictionary")
salaries.Add "Kevin", "20,000.00"
salaries.Add "Sam", "30,000.00"
salaries.Add "Peter", "10,000.00"

name = "Tina"
If salaries.Exists(name) Then
    WScript.Echo "We found the employee!"
    WScript.Echo salaries(name)
Else
    WScript.Echo "Employee Not Found!!!"
End If

' Keys
WScript.Echo "Keys >>>>>"
keys = salaries.Keys
WScript.Echo "keys variable type : " & VarType(keys)

For each key in keys
    WScript.Echo "Key : " & key
Next
WScript.Echo "-----------------------"

' Items
WScript.Echo "Items >>>>>"
items = salaries.Items
WScript.Echo "Items variable type : " & VarType(items)

For each item in items
    WScript.Echo "Item : " & item
Next