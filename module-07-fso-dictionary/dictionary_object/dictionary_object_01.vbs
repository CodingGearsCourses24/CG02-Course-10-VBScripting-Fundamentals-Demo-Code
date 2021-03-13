'=========================================================================
' Dictionary Object
'=========================================================================

Dim salaries

Set salaries = CreateObject("Scripting.Dictionary")

'Add
salaries.Add "Peter", "Apple Inc"
salaries.Add "John", "Microsoft Inc"

' Count
WScript.Echo "Count         : " & salaries.count
WScript.Echo

'Add
salaries.Add "Kelly", "Facebook Inc"
salaries.Add "Tina", "Self-employed"
WScript.Echo "Added 2 more items..."

' Count
WScript.Echo "Count         : " & salaries.count
WScript.Echo

'Remove
salaries.Remove("Peter")
WScript.Echo "Removed 1 item from the dictionary obj..."

' Count
WScript.Echo "Count         : " & salaries.count ' 3
WScript.Echo

' update
WScript.Echo "Tina ? " & salaries("Tina")
salaries("Tina") = "Oracle Inc"
WScript.Echo "Tina ? " & salaries("Tina")

' RemoveAll
salaries.RemoveAll
WScript.Echo "Removed all items..."
WScript.Echo "Count after removing all : " & salaries.count


