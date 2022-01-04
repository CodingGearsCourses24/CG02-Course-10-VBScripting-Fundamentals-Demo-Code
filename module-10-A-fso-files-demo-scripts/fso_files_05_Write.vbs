'------------------------------------------------------------
' We will Explore:(* = Required)
'			FileSystemObject.CreateTextFile(filename*, overwrite, unicode)
'							overwrite: True or False
'									True --> Overwrite the file.
'									False --> Don't Overwrite the file. (Default)
'			FileSystemObject.OpenTextFile(filename*, mode, create, format)
'
' 			File.OpenAsTextStream(mode, format) 
'
'							Mode: 1 for Reading
'								  2 for Writing
'								  8 for Appending	
'
'			Format : -1 to open file as Unicode
'				  	  0 to open file as ASCII
'				  	 -2 to open file as the system default
'
'			Write ==> Writes without a trailing newline character
'			WriteLine ==> Writes with a trailing newline character
'			WriteBlankLines ==> Writes blank line
'------------------------------------------------------------

'------------------------------------------------------------
' Variables
'------------------------------------------------------------
Option Explicit

Dim oFSO, oFile, textstream
Dim myfile1, myfile2, myfile3, myfile4, myfile5
Dim myfile1_path, myfile2_path, myfile3_path, myfile4_path, myfile5_path

Const tmp_folder1="D:\VBScripts_Files\tmp1"
Const tmp_folder2="D:\VBScripts_Files\tmp2"

myfile1="mytestfile1.txt"
myfile2="mytestfile2.txt"
myfile3="mytestfile3.txt"
myfile4="mytestfile4.txt"
myfile5="mytestfile5.txt"

myfile1_path = tmp_folder1 & "\" & myfile1
myfile2_path = tmp_folder1 & "\" & myfile2
myfile3_path = tmp_folder1 & "\" & myfile3
myfile4_path = tmp_folder1 & "\" & myfile4
myfile5_path = tmp_folder1 & "\" & myfile5

Const OpenFileForReading = 1 
Const OpenFileForWriting = 2 
Const OpenFileForAppending = 8 

'************************************************************
' Procedure Call
'************************************************************
Call WriteToFile1

'------------------------------------------------------------
' Using FileSystemObject.CreateTextFile 
'------------------------------------------------------------
Sub WriteToFile1
	Set oFSO = CreateObject("Scripting.FileSystemObject")

	Set textstream = oFSO.CreateTextFile(myfile1_path, true) 
	
	textstream.Write("This is my line 1")
	textstream.Write("This is my line 2")
	textstream.WriteLine("This is my line 3")
	textstream.WriteLine("This is my line 4")
	textstream.WriteBlankLines 2
	textstream.WriteLine("This is my line 5")
	textstream.Close
	
	Set oFSO = Nothing
End Sub

'------------------------------------------------------------
' Using FileSystemObject.OpenTextFile 
'------------------------------------------------------------
Sub WriteToFile2
	Set oFSO = CreateObject("Scripting.FileSystemObject")

	Set textstream = oFSO.OpenTextFile(myfile2_path, OpenFileForWriting, true) 
	
	textstream.Write("This is my line 21")
	textstream.Write("This is my line 22")
	textstream.WriteLine("This is my line 23")
	textstream.WriteLine("This is my line 24")
	textstream.WriteBlankLines 2
	textstream.WriteLine("This is my line 25")
	textstream.Close
	
	Set textstream = Nothing
	Set oFSO = Nothing
End Sub

'------------------------------------------------------------
' Using File.OpenAsTextStream  - Writing
'------------------------------------------------------------
Sub WriteToFile3
	Set oFSO = CreateObject("Scripting.FileSystemObject")

	If oFSO.FileExists(myfile3_path) Then
		MsgBox "File found. We can proceed...", 0, "Alert :"
	Else
		MsgBox "File not found. File will be created...", 0, "Alert :"
		oFSO.CreateTextFile(myfile3_path)
	End if
	
	Set oFile = oFSO.GetFile(myfile3_path) 
	
	Set textstream = oFile.OpenAsTextStream(OpenFileForWriting)
	
	textstream.Write("This is my line 31")
	textstream.Write("This is my line 32")
	textstream.WriteLine("This is my line 33")
	textstream.WriteLine("This is my line 34")
	textstream.WriteBlankLines 2
	textstream.WriteLine("This is my line 35")
	textstream.Close
	
	Set oFile = Nothing
	Set oFSO = Nothing
End Sub

'------------------------------------------------------------
' Using FileSystemObject.OpenAsTextStream  - Appending
'------------------------------------------------------------
Sub WriteToFile4
	Set oFSO = CreateObject("Scripting.FileSystemObject")

	If oFSO.FileExists(myfile4_path) Then
		MsgBox "File found. We can proceed...", 0, "Alert :"
	Else
		MsgBox "File not found. File will be created...", 0, "Alert :"
		oFSO.CreateTextFile(myfile4_path)
	End if
	
	Set oFile = oFSO.GetFile(myfile4_path) 
	
	Set textstream = oFile.OpenAsTextStream(OpenFileForAppending)
	
	textstream.WriteLine("This is my line 41")
	textstream.WriteLine("This is my line 42")
	textstream.WriteLine("This is my line 43")
	textstream.WriteLine("This is my line 44")
	textstream.WriteLine("This is my line 45")
	textstream.WriteLine("------------------")
	textstream.Close
	
	Set oFile = Nothing
	Set oFSO = Nothing
End Sub

'------------------------------------------------------------
' Using FileSystemObject.OpenAsTextStream  - Appending
'------------------------------------------------------------
Sub WriteToFile5
	Set oFSO = CreateObject("Scripting.FileSystemObject")

	If oFSO.FileExists(myfile5_path) Then
		MsgBox "File found. We can proceed...", 0, "Alert :"
	Else
		MsgBox "File not found. File will be created...", 0, "Alert :"
		oFSO.CreateTextFile(myfile5_path)
	End if
	
	Set oFile = oFSO.GetFile(myfile5_path) 
	
	Set textstream = oFile.OpenAsTextStream(OpenFileForReading)
	
	textstream.WriteLine("This is my line 51")
	textstream.WriteLine("This is my line 52")
	textstream.WriteLine("This is my line 53")
	textstream.WriteLine("This is my line 54")
	textstream.WriteLine("This is my line 55")
	textstream.WriteLine("------------------")
	textstream.Close
	
	Set oFile = Nothing
	Set oFSO = Nothing
End Sub