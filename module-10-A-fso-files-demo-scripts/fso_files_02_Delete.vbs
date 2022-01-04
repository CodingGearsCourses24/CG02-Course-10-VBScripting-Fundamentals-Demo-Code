'------------------------------------------------------------
' We will Explore:(* = Required)
'				FileSystemObject.DeleteFile( filepath*, force)
' 				File.Delete (force)
'------------------------------------------------------------

'------------------------------------------------------------
' Variables
'------------------------------------------------------------
Option Explicit

Dim oFSO, oFile, textStream
Dim my_file1, my_file2, my_file3, my_file1_path, my_file2_path, my_file3_path

Const ForWriting = 2
Const tmp_folder1="D:\VBScripts_Files\tmp1"

my_file1="MyNewFile1.txt"
my_file1_path = tmp_folder1 & "\" & my_file1

my_file2="MyNewFile2.txt"
my_file2_path = tmp_folder1 & "\" & my_file2

my_file3="MyNewFile3.txt"
my_file3_path = tmp_folder1 & "\" & my_file3

'------------------------------------------------------------
' Procedure Call
'------------------------------------------------------------
Call DeleteFile3

'------------------------------------------------------------
' Using File.Delete method
'------------------------------------------------------------
Sub DeleteFile1
	Set oFSO = CreateObject("Scripting.FileSystemObject")

	Set oFile = oFSO.GetFile(my_file1_path) 
	
	oFile.Delete
	
	Set oFile = Nothing
	Set oFSO = Nothing
End Sub

'------------------------------------------------------------
' Using FileSystemObject.DeleteFile method
'------------------------------------------------------------
Sub DeleteFile2
	Set oFSO = CreateObject("Scripting.FileSystemObject")

	oFSO.DeleteFile(my_file2_path) 
	
	Set oFile = Nothing
	Set oFSO = Nothing
End Sub

'------------------------------------------------------------
' Using FileSystemObject.DeleteFile method
'------------------------------------------------------------
Sub DeleteFile3
	Set oFSO = CreateObject("Scripting.FileSystemObject")

	If oFSO.FileExists(my_file3_path) Then
		Set oFile = oFSO.Getfile(my_file3_path)
		oFile.Delete ()
	Else
		MsgBox my_file3_path & " " & "not found.", 0, "Alert :"
	End if
	
	Set oFile = Nothing
	Set oFSO = Nothing
End Sub