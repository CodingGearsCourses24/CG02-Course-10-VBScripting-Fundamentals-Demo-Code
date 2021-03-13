'=========================================================================
' Loop: Do 
'=========================================================================

Dim count

count = 1
Do
    MsgBox count & " : I am inside the loop!"
    count = count + 1
    If count = 10 Then
        'WScript.Quit
        Exit Do
    End If
Loop

MsgBox count & " : I am OUTSIDE the loop!"
