'=========================================================================
' Loop: For Each...Next
'=========================================================================

Option Explicit

Dim arrNums, item

arrNums = Array(1, 2, 3, 4, 5)

For Each item in arrNums
    ' Do something
    WScript.Echo item
Next