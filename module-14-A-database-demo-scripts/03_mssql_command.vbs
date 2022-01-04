'=================================================================================
' CodingGears.io
' Connecting to MS SQL Server
'=================================================================================

Dim objConnection

' SQL Statement
sqlText1 = "SELECT TOP (2) CustomerID," &_
           "FirstName, LastName, EmailAddress " &_
           "FROM SalesLT.Customer"

' Declaring the ADO connection object
Set objConnection = CreateObject("ADODB.Connection")
Set objCommand = CreateObject("ADODB.Command")

' Declaring the connection string
mssql_connection_string =  "Provider=MSOLEDBSQL;" &_
                            "Server=SCUBE02\SQLExpress;" &_
                            "Database=AdventureWorksLT2019;" &_ 
                            "UID=student1;" &_
                            "PWD=pass123;"
                            
objConnection.ConnectionString = mssql_connection_string
objConnection.Open

objCommand.ActiveConnection = objConnection
objCommand.CommandText = sqlText1
objCommand.Prepared = True

Set rs1 = objCommand.Execute  

do until rs1.EOF
    for each x in rs1.Fields
        WScript.Echo x.name & " : " & x.value
    next
    rs1.MoveNext
    WScript.Echo "-----------------------"
Loop

objConnection.Close