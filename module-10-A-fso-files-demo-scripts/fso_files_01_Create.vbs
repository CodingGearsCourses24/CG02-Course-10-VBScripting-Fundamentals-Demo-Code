'------------------------------------------------------------
' We will Explore:(* = Required)
'				CreateTextFile(filename*, overwrite, unicode)
'				OpenTextFile(filename*, mode, create, format)
'							Mode: 1 for Reading
'								  2 for Writing
'								  8 for Appending	
'		  Objects: 
' 				FileSystemObject or File object.
'------------------------------------------------------------

'------------------------------------------------------------
' Variables
'------------------------------------------------------------
Option Explicit

Dim oFSO,textStream
Dim my_file1, my_file2, my_file3, my_file1_path, my_file2_path, my_file3_path

Const ForWriting = 2
Const tmp_folder1="D:\VBScripts_Files\tmp1"

my_file1="MyNewFile1.txt"
my_file1_path = tmp_folder1 & "\" & my_file1

my_file2="MyNewFile2.txt"
my_file2_path = tmp_folder1 & "\" & my_file2

my_file3="MyNewFile3.txt"
my_file3_path = tmp_folder1 & "\" & my_file3

'************************************************************
' Procedure Call
'************************************************************
Call CreateFile3

'************************************************************
'Sub-routines & Functions
'************************************************************

'------------------------------------------------------------
'Using CreateTextFile method of the FSO object
'------------------------------------------------------------
Sub CreateFile1
	Set oFSO = CreateObject("Scripting.FileSystemObject")

	oFSO.CreateTextFile(my_file1_path) 

	Set oFSO = Nothing
End Sub

'------------------------------------------------------------
' Using OpenTextFile method of the FSO object
'------------------------------------------------------------
Sub CreateFile2
	Set oFSO = CreateObject("Scripting.FileSystemObject")

	Set textStream = oFSO.OpenTextFile(my_file2_path, ForWriting, True)
	
	Set oFSO = Nothing
End Sub

'------------------------------------------------------------
' Using CreateTextFile method of the FSO object
'------------------------------------------------------------
Sub CreateFile3
	Set oFSO = CreateObject("Scripting.FileSystemObject")

	If oFSO.FileExists(my_file3_path) Then
		MsgBox my_file3_path & " already exists! ", 0, "Alert :"
	Else
		oFSO.CreateTextFile(my_file3_path)
	End if
	
	Set oFSO = Nothing
End Sub