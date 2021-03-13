'------------------------------------------------------------
' Read CSV file 
'------------------------------------------------------------

'------------------------------------------------------------
' Variables
'------------------------------------------------------------
Option Explicit

Dim oFSO, oFile, textstream
Dim myfile1, line, array1
Dim myfile1_path
Dim name, phone

Const tmp_data_folder1="D:\VBScripts_Files\tmp_data"

myfile1="AddressBook.txt"
myfile1_path = tmp_data_folder1 & "\" & myfile1

Const OpenFileForReading = 1 
Const OpenFileForWriting = 2 
Const OpenFileForAppending = 8

'------------------------------------------------------------
' Procedure Call
'------------------------------------------------------------
Call ReadCsvFile

'------------------------------------------------------------
' Using FileSystemObject.OpenTextFile  Read CSV file
'------------------------------------------------------------
Sub ReadCsvFile
	Set oFSO = CreateObject("Scripting.FileSystemObject")

	Set textstream = oFSO.OpenTextFile(myfile1_path, OpenFileForReading) 
	
	Do Until textstream.AtEndOfStream
		line = textstream.ReadLine
		array1 = split(line, ",")
		name = array1(0)
		phone = array1(1)
		MsgBox name & "'s phone number is " & phone, 0, "Reading..."
	Loop

	Set oFSO = Nothing
End Sub