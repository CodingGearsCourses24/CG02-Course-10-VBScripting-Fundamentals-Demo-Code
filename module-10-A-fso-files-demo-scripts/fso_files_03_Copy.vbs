'------------------------------------------------------------
' We will Explore:(* = Required)
'				FileSystemObject.CopyFile(source*, destination*, overwrite )
' 				File.Copy
'------------------------------------------------------------

'------------------------------------------------------------
' Variables
'------------------------------------------------------------
Option Explicit

Dim oFSO, oFile
Dim source_file, target_file
Dim source_file_path, target_file_path

Const ForWriting = 2
Const tmp_folder1="D:\VBScripts_Files\tmp1"
Const tmp_folder2="D:\VBScripts_Files\tmp2"

source_file="cities.txt"
source_file_path = tmp_folder1 & "\" & source_file

'************************************************************
' Procedure Call
'************************************************************
Call CopyFile3

'------------------------------------------------------------
' Using File.Copy method
'------------------------------------------------------------
Sub CopyFile1
	target_file="cities_copy1.txt"
	
	target_file_path = tmp_folder2 & "\" & target_file
	
	Set oFSO = CreateObject("Scripting.FileSystemObject")

	Set oFile = oFSO.GetFile(source_file_path) 
	
	oFile.Copy(target_file_path)
	
	Set oFile = Nothing
	Set oFSO = Nothing
End Sub

'------------------------------------------------------------
' Using FileSystemObject.CopyFile method
'------------------------------------------------------------
Sub CopyFile2
	target_file="cities_copy2.txt"
	
	target_file_path = tmp_folder2 & "\" & target_file
	
	Set oFSO = CreateObject("Scripting.FileSystemObject")

	oFSO.CopyFile source_file_path, target_file_path, True
	
	Set oFile = Nothing
	Set oFSO = Nothing
End Sub

'------------------------------------------------------------
' Using FileSystemObject.DeleteFile method
'------------------------------------------------------------
Sub CopyFile3
	target_file="cities_copy3.txt"
	
	target_file_path = tmp_folder2 & "\" & target_file
	
	Set oFSO = CreateObject("Scripting.FileSystemObject")
	
	If oFSO.FileExists(source_file_path) Then
		oFSO.CopyFile source_file_path, target_file_path, True
	Else
		MsgBox source_file_path & " " & "not found.", 0, "Alert :"
	End if
	
	Set oFile = Nothing
	Set oFSO = Nothing
End Sub