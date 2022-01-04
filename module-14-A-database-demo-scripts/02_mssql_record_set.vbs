'=================================================================================
' CodingGears.io
' Connecting to MS SQL Server
'=================================================================================

Dim objConnection

' Declare variables
Const adOpenStatic = 3 'A static cursor
Const adLockOptimistic = 3 'Lock records only when calling update

' Declaring the ADO connection object
Set objConnection = CreateObject("ADODB.Connection")

' Declaring the connection string
mssql_connection_string =  "Provider=MSOLEDBSQL;" &_
                            "Server=SCUBE02\SQLExpress;" &_
                            "Database=AdventureWorksLT2019;" &_ 
                            "UID=student1;" &_
                            "PWD=pass123;"
objConnection.ConnectionString = mssql_connection_string

' Open
objConnection.Open

' Sql Text
sqlText1 = "SELECT TOP(2) CustomerID," &_
           "FirstName, LastName, EmailAddress " &_
           "FROM SalesLT.Customer"
' Record Set
 Set objRecordSet = CreateObject("ADODB.RecordSet")

' Open Record Set
objRecordSet.open sqlText1, objConnection, adOpenStatic, adLockOptimistic

' Get Record Count
WScript.Echo "Record Count : " & objRecordSet.RecordCount
WScript.Echo

' Loop thru the record set
DO UNTIL objRecordSet.EOF
    WScript.Echo "Customer ID : " & objRecordSet.Fields("CustomerID")
    WScript.Echo "First Name : " & objRecordSet.Fields("FirstName")
    WScript.Echo "Last Name : " & objRecordSet.Fields("LastName")
    WScript.Echo "EmailAddress : " & objRecordSet.Fields("EmailAddress")
    WScript.Echo 
    objRecordSet.MoveNext
LOOP
' Closing
