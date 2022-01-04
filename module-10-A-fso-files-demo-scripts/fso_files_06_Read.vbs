'------------------------------------------------------------
'
' You need a textstream to read file. 
' You can get a textstream using the FSO or File object.
'
'			FileSystemObject.OpenTextFile(filename*, mode, create, format)
' 			File.OpenAsTextStream(mode, format) 
'
'			Mode: 1 for Reading
'				  2 for Writing
'				  8 for Appending	
'
'			Format : -1 to open file as Unicode
'				  	  0 to open file as ASCII
'				  	 -2 to open file as the system default
'	
' We will Explore:
'			TextStream Object Methods:
'			Read ==> Read a specified number of characters.
'			ReadLine ==> Read one line.
'			ReadAll ==> Read the entire text file.
'			AtEndOfStream ==> Reads upto the end of file
'------------------------------------------------------------

'------------------------------------------------------------
' Variables
'------------------------------------------------------------
Option Explicit

Dim oFSO, oFile, textstream
Dim myfile1, myfile2, myfile3, myfile1_path, myfile2_path, myfile3_path
Dim line, line_no, chars
Dim tmp1

Const tmp_data_folder="D:\VBScripts_Files\tmp_data"

myfile1="Cars.txt"
myfile2="Cities.txt"
myfile3="US_States.txt"

myfile1_path = tmp_data_folder & "\" & myfile1
myfile2_path = tmp_data_folder & "\" & myfile2
myfile3_path = tmp_data_folder & "\" & myfile3

Const OpenFileForReading = 1 
Const OpenFileForWriting = 2 
Const OpenFileForAppending = 8 

'------------------------------------------------------------
' Procedure Call
'------------------------------------------------------------
'Call ReadFile1_LineByLine
'Call ReadFile2_CharByChar
'Call ReadFile3_EntireContents
Call ReadFile4_Using_FileObject

'------------------------------------------------------------
' Using FileSystemObject.OpenTextFile  Line By Line (ReadLine)
'------------------------------------------------------------
Sub ReadFile1_LineByLine
	Set oFSO = CreateObject("Scripting.FileSystemObject")

	Set textstream = oFSO.OpenTextFile(myfile1_path, OpenFileForReading) 
	line_no = 0
	Do Until textstream.AtEndOfStream
		line_no = line_no + 1
		line = textstream.ReadLine
		MsgBox line_no & " : " & line, 0, "Reading..."
	Loop

	Set oFSO = Nothing
End Sub

'------------------------------------------------------------
' Using FileSystemObject.OpenTextFile  Char by Char (Read)
'------------------------------------------------------------
Sub ReadFile2_CharByChar
	Set oFSO = CreateObject("Scripting.FileSystemObject")

	Set textstream = oFSO.OpenTextFile(myfile2_path, OpenFileForReading) 
	
	Do Until textstream.AtEndOfStream
		chars = textstream.Read(1)
		MsgBox chars, 0, "Reading..."
	Loop

	Set oFSO = Nothing
End Sub

'------------------------------------------------------------
' Using FileSystemObject.OpenTextFile  Complete File (ReadAll)
'------------------------------------------------------------
Sub ReadFile3_EntireContents
	Set oFSO = CreateObject("Scripting.FileSystemObject")

	Set textstream = oFSO.OpenTextFile(myfile3_path, OpenFileForReading) 
	
	line = textstream.ReadAll
	MsgBox line, 0, "Reading..."
		
	Set oFSO = Nothing
End Sub

'------------------------------------------------------------
' Using File.OpenAsTextStream(mode, format) 
'------------------------------------------------------------
Sub ReadFile4_Using_FileObject
	Set oFSO = CreateObject("Scripting.FileSystemObject")

	Set oFile = oFSO.GetFile(myfile3_path) 
	
	Set textstream = oFile.OpenAsTextStream(OpenFileForReading,-2)
	
	line = textstream.ReadAll
	MsgBox line, 0, "Reading..."
		
	Set oFSO = Nothing
End Sub