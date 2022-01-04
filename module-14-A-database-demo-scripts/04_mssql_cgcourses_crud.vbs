' Database

option explicit

' Declare variables
Const adOpenStatic = 3
Const adLockOptimistic = 3
Dim objConnection, objRecordSet
Dim strSQL, tmp


CreateADOObjects
'AddSampleData
'DisplayAllCourses
'DeleteAllRecords
'DeleteWithWhere
'UpdateCourseDescription
CloseADOObjects

' Create Objects
Sub CreateADOObjects
    Set objConnection = CreateObject("ADODB.Connection")
    objConnection.ConnectionString = ConnectionString()
    objConnection.Open
    Set objRecordSet = CreateObject("ADODB.Recordset")
End Sub

' Add sample data to the courses table
Sub AddSampleData
     strSQL = "SELECT *  FROM [CodingGearsDB].[dbo].[Courses]"
     objRecordSet.Open strSQL, objConnection, adOpenStatic, adLockOptimistic

     AddRecord "CG101","Software Development Processes","Learn SDLC Processes","Active"
     AddRecord "CG102","VBScripting Fundamentals","Simple Scripting","Active"
     AddRecord "CG103","Mastering Linux Command Line","Linux for everyone","Active"
     AddRecord "CG104","Mastering Bash Shell Scripting","Automate yout daily tasks!","Active"
     AddRecord "CG105","Mastering VirtualBox","Creat virtual machines","Active"
End Sub

' Close Objects
Sub CloseADOObjects
     objConnection.Close
End Sub

' Returns the connection string
Function ConnectionString()
     ConnectionString = "Provider=MSOLEDBSQL;" &_
                        "Server=SCUBE02\SQLExpress;" &_
                        "Database=CodingGearsDB;" &_ 
                        "UID=student1;" &_
                        "PWD=pass123;"
End Function

' Displays all courses
Sub DisplayAllCourses
     strSQL = "SELECT *  FROM [CodingGearsDB].[dbo].[Courses]"
     objRecordSet.Open strSQL, objConnection, adOpenStatic, adLockOptimistic
     With objRecordSet
          Do While NOT objRecordSet.EOF
               WScript.Echo trim(objRecordSet("CourseID")) & " is " & objRecordSet("Description")
               objRecordSet.MoveNext
          Loop
     End With
End Sub

' Add Record
Function AddRecord(courseId, name, description, status)
     objRecordSet.AddNew
     objRecordSet("CourseID") = courseId
     objRecordSet("Name") = name
     objRecordSet("Description") = description
     objRecordSet("Status") = status
     objRecordSet.Update
End Function

' Update Course Description
Sub UpdateCourseDescription
     objConnection.Execute "Update [CodingGearsDB].[dbo].[Courses] SET Description = 'New Description 2' WHERE CourseID='CG105'"
End Sub

' Delete with where condition
Sub DeleteWithWhere
     objConnection.Execute "Delete FROM [CodingGearsDB].[dbo].[Courses] WHERE CourseID='CG104'"
End Sub

' Delete all records from courses
Sub DeleteAllRecords
     objConnection.Execute "Delete FROM [CodingGearsDB].[dbo].[Courses]"
End Sub