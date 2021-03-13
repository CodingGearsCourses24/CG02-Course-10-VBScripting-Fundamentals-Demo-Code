'------------------------------------------------------------
' We will Explore:
'				DeleteFolder(folder to be deleted)
'------------------------------------------------------------

Dim objFSO, strFolder

strFolder="D:\VBScripts_Folders\tmp\MyFolderData-copy1"

Set objFSO = CreateObject("Scripting.FileSystemObject")

objFSO.DeleteFolder(strFolder)

Set objFSO = Nothing