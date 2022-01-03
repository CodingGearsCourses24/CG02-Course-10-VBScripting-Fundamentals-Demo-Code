Dim arr(3,4) ' 4 Rows 5 columns
 
'Filling 2 Dimenional array
For i=0 to Ubound(arr,1) '  Ubound of first dimension
    For j=0 to Ubound(arr,2) ' Ubound of Second dimension
        arr(i,j) = "Row: " & i & " , " & "Column: " & j
    Next
Next
 
'Fetching values from  2 Dimenional array
For i=0 to Ubound(arr,1) '  Ubound of first dimension
    For j=0 to Ubound(arr,2) ' Ubound of Second dimension
        WScript.Echo arr(i,j)
    Next
Next