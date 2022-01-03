'------------------------------------------------------------
' We will Explore:
'				MoveFolder <source folder>, <destination folder>
'------------------------------------------------------------

Dim oFSO, sourceFolder, targetFolder

sourceFolder="D:\VBScripts_Folders\tmp\MyFolderData"
targetFolder="D:\VBScripts_Folders\tmp2\MyFolderData"

Set oFSO=createobject("Scripting.Filesystemobject")

oFSO.MoveFolder sourceFolder, targetFolder

set oFSO = Nothing

MsgBox "Done",0,"Alert:"