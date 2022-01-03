'------------------------------------------------------------
' We will Explore:
'				FolderExists - Checking the existence of a folder
'				Nothing - What is this?
'------------------------------------------------------------

Dim oFSO, strDirectory

strDirectory = "D:\VBScripts_tmp\MyFolder2"

Set oFSO = CreateObject("Scripting.FileSystemObject")
	
If oFSO.FolderExists(strDirectory) Then
	msgbox strDirectory & " already created "
else
	oFSO.CreateFolder(strDirectory)
end if

Set oFSO = Nothing