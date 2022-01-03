' Assignment Solution

Dim arrNames(4)

arrNames(0) = "John"
arrNames(1) = "Kerry"
arrNames(2) = "Palm"
arrNames(3) = "Yola"
arrNames(4) = "David"

upperIndex = UBound(arrNames)

count = upperIndex + 1

For x = 0 To upperIndex Step 1

	MsgBox arrNames(x), 0, "Element :"

Next

MsgBox "Total number of elements in the array are " & count, 0, "Element :"

