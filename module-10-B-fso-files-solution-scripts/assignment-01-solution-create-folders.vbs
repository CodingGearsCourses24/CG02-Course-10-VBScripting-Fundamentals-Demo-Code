' **********************************************************************
' http://www.CodingGears.com
' FSO - Read folder list from a file & create them
' Assignment 2 - Solution
' **********************************************************************

' Constants
Const OpenFileForReading = 1 
Const OpenFileForWriting = 2 
Const OpenFileForAppending = 8 

' Vars
Dim folders_dict, oFSO

' Objects
Set folders_dict = CreateObject("Scripting.Dictionary")
Set oFSO = CreateObject("Scripting.FileSystemObject")

' Parent Name
parent_folder_name = InputBox("Enter the parent folder name: ", SITE_TITLE, "Enter Input Here", 1000, 5000) 

' File with folder list
file_folder_list = GetFileWithFolderList()
scriptdir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)

oFSO.CreateFolder(parent_folder_name)
CreateFolders(file_folder_list)

' Read file
Sub CreateFolders(file_folder_list)
	Set textstream = oFSO.OpenTextFile(file_folder_list, OpenFileForReading) 
	line_no = 0
	Do Until textstream.AtEndOfStream
		line_no = line_no + 1
		line = textstream.ReadLine
		oFSO.CreateFolder(scriptdir + "\" + parent_folder_name + "\" + line )
	Loop
	Set oFSO = Nothing
End Sub

Function GetFileWithFolderList()
    script_name = Wscript.ScriptName
	arr_script_name = Split(script_name, ".")
	script_name_no_ext = arr_script_name(0)
	GetFileWithFolderList = script_name_no_ext + ".txt"
End Function