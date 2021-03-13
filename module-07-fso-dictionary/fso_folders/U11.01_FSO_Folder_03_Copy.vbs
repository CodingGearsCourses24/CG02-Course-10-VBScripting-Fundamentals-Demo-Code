'------------------------------------------------------------
' We will Explore:
'				CopyFolder source, destination[, overwrite]
'------------------------------------------------------------

Dim oFSO, sourceFolder, destFolder

Set oFSO=createobject("Scripting.Filesystemobject")

sourceFolder="D:\VBScripts_Folders\tmp\MyFolderData"
destFolder="D:\VBScripts_Folders\tmp\MyFolderData-copy3"

oFSO.CopyFolder sourceFolder, destFolder, True

Set oFSO = Nothing