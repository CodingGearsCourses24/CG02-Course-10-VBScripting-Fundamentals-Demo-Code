' ===========================================================
' Using ByRef & ByVal with parameters/arguments
' 		ByVal ==> Passed by value
' 		ByRef ==> Passed by Reference
' ===========================================================

Class Student 'Class
    Public MyStudentId
End Class

Sub ChangeStudentId (ByRef MyId) 'Method
    MyId = 55555
End Sub

Dim Student1
Set student1 = New Student
student1.MyStudentId = 11111

ChangeStudentId student1.MyStudentId

MsgBox "M1: The updated student id is " & student1.MyStudentId 'student1.MyStudentId
