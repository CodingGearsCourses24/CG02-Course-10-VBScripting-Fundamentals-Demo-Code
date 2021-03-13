'------------------------------------------------------------
' We will Explore:
'				DeleteFolder
' 				FolderExists
'------------------------------------------------------------

Dim objFSO, strDirectory

strDirectory = "D:\VBScripts_Folders\tmp\MyFolderData-copy2"

Set objFSO = CreateObject("Scripting.FileSystemObject")
	
If objFSO.FolderExists(strDirectory) Then
	objFSO.DeleteFolder(strDirectory)
	msgbox strDirectory & " -- Deleted! "
else
	msgbox strDirectory & " folder does not exist! ", 0, "Alert!"
end if

Set objFSO = Nothing