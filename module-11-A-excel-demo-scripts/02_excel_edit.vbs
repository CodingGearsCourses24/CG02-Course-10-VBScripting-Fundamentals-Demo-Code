'---------------------------------------------------------------------
' We will Explore:
'	- [Excel.Application] -> [Workbook] -> [Sheets] -> [Rows, Columns, Cells]
'	- Editing Excel Document
'	- We will use Sheet object here.
' 	-		Example: Range("A1").Value = "John Smith"
'	- We need to get to the desired cell, and then set the value.
'---------------------------------------------------------------------

Option Explicit

Dim objExcel
Dim objWorkbook
Dim objWorkSheet
Dim ObjSheet

Dim sheetName
Dim strDirectory
Dim excelDocName

Dim nametmp, x

strDirectory = "D:\GitBucket\vbscripting-course\02-indev\excel\tmp_data"
excelDocName = "Sample3.xlsx"
sheetName = "Data1"

'Excel Object
Set objExcel = CreateObject("Excel.Application")

'Excel Document Visible Property
objExcel.visible = false

'Workbook Object
Set objWorkbook = objExcel.Workbooks.Open(strDirectory & "\" &  excelDocName) 

'Set the value of the Particular Cell 
'objExcel.Sheets(sheetName).Range("A3").Value = "John Smith"
'objExcel.Sheets(sheetName).Range("B3").Value = "45"
'objExcel.Sheets(sheetName).Range("C3").Value = "John.Smith@gmail.com"

'objExcel.Sheets(sheetName).Range("A4").Value = "Terry Woola"
'objExcel.Sheets(sheetName).Range("B4").Value = "50"
'objExcel.Sheets(sheetName).Range("C4").Value = "twoola@gmail.com"

'objExcel.Sheets(sheetName).Range("A5").Value = "Larry Port"
'objExcel.Sheets(sheetName).Range("B5").Value = "49"
'objExcel.Sheets(sheetName).Range("C5").Value = "lport@gmail.com"

'Update email address for Larry Port
for x = 1 to 5 step +1
	nametmp = objExcel.Sheets(sheetName).Cells(x,1).Value
	if nametmp = "Larry Port" Then
		objExcel.Sheets(sheetName).Cells(x,3).Value = "lport1234@gmail.com"
	End If
Next

msgbox "Editing Completed!              ", 0, "Status..."

'Save, Close & Quit
objWorkbook.Save
objWorkbook.Close 
objExcel.Quit

'Release Memory
Set objExcel = Nothing
Set objWorkbook = Nothing