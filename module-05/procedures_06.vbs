'=========================================================================
' Using functions from external vbscript file
'=========================================================================

option explicit

'**********************************************************************
'Sub to read external vbs file
Sub Include(extVBScriptFile)
	Dim objFso, objExtFile
	Dim strfileContent, strScriptDir
  
	strScriptDir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
	
	Set objFso = CreateObject("Scripting.FileSystemObject")
	Set objExtFile = objFso.OpenTextFile(strScriptDir & "\" & extVBScriptFile, 1)
	strfileContent = objExtFile.ReadAll
	objExtFile.Close
	ExecuteGlobal strfileContent
	Set objFso = Nothing
	Set objExtFile = Nothing
End Sub
'**********************************************************************

Include "procedures_05_functions_only.vbs"

' Variables
Dim a, b, result

a = 10
b = 8

result = Add(a, b)

DisplayMsg "The result is", result