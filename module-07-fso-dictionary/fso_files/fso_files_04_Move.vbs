'------------------------------------------------------------
' We will Explore:(* = Required)
'				FileSystemObject.MoveFile( source*, destination*)
' 				File.Move(destination*)
'------------------------------------------------------------

'------------------------------------------------------------
' Variables
'------------------------------------------------------------
Option Explicit

Dim oFSO, oFile
Dim source_file, target_file 
Dim source_file_path, target_file_path

Const tmp_folder1="D:\VBScripts_Files\tmp1"
Const tmp_folder2="D:\VBScripts_Files\tmp2"

'************************************************************
' Procedure Call
'************************************************************
Call MoveFile3

'------------------------------------------------------------
' Using File.Move method
'------------------------------------------------------------
Sub MoveFile1
	source_file="cars.txt"
	source_file_path = tmp_folder1 & "\" & source_file

	target_file="cars_move1.txt"
	target_file_path = tmp_folder2 & "\" & target_file
	
	Set oFSO = CreateObject("Scripting.FileSystemObject")

	Set oFile = oFSO.GetFile(source_file_path) 
	
	oFile.Move(target_file_path)
	
	Set oFile = Nothing
	Set oFSO = Nothing
End Sub

'------------------------------------------------------------
' Using FileSystemObject.MoveFile method
'------------------------------------------------------------
Sub MoveFile2
	source_file="cities.txt"
	source_file_path = tmp_folder1 & "\" & source_file

	target_file="cities_move2.txt"
	target_file_path = tmp_folder2 & "\" & target_file
	
	Set oFSO = CreateObject("Scripting.FileSystemObject")

	oFSO.MoveFile source_file_path, target_file_path
	
	Set oFile = Nothing
	Set oFSO = Nothing
End Sub

'------------------------------------------------------------
' Using FileSystemObject.DeleteFile method
'------------------------------------------------------------
Sub MoveFile3
	source_file="countries.txt"
	source_file_path = tmp_folder1 & "\" & source_file

	target_file="countries_move3.txt"
	target_file_path = tmp_folder2 & "\" & target_file
	
	Set oFSO = CreateObject("Scripting.FileSystemObject")
	
	If oFSO.FileExists(source_file_path) Then
		oFSO.MoveFile source_file_path, target_file_path
	Else
		MsgBox source_file_path & " " & "not found.", 0, "Alert :"
	End if
	
	Set oFile = Nothing
	Set oFSO = Nothing
End Sub