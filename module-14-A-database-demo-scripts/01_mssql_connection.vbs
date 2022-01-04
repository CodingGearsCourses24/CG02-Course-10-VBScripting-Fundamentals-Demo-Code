'=================================================================================
' CodingGears.io
' Connecting to MS SQL Server
'=================================================================================

Dim objConnection

' Declaring the ADO connection object
Set objConnection = CreateObject("ADODB.Connection")

' Declaring the connection string
' Microsoft OLE DB Driver for SQL Server (MSOLEDBSQL)
mssql_connection_string1 =  "Provider=MSOLEDBSQL;" &_
                            "Server=SCUBE02\SQLExpress;" &_
                            "Database=AdventureWorksLT2019;" &_ 
                            "UID=student1;" &_
                            "PWD=pass123;"

mssql_connection_string2 =  "Provider=MSOLEDBSQL;" &_
                            "Server=SCUBE02\SQLExpress;" &_
                            "Database=AdventureWorksLT2019;" &_ 
                            "Trusted_Connection=yes;"

objConnection.ConnectionString = mssql_connection_string2

WScript.Echo "Status before opening : " & objConnection.State

objConnection.Open
WScript.Echo "Status after opening : " & objConnection.State

objConnection.Close
WScript.Echo "Status after closing : " & objConnection.State