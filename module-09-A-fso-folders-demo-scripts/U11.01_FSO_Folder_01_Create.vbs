'------------------------------------------------------------
' We will Explore:
'		CreateFolder
'------------------------------------------------------------

Dim oFSO, strFolder

strFolder = "D:\VBScripts_Folders\tmp\MyFolder2"

Set oFSO = CreateObject("Scripting.FileSystemObject")

if oFSO.FolderExists(strFolder) Then
	MsgBox strFolder & " already exists.", 0, "Result:"
else
	oFSO.CreateFolder(strFolder)
end If

Set oFSO = Nothing